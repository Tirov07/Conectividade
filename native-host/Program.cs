using System.Collections.Concurrent;
using System.Text.Json;
using System.Xml;
using System.Xml.Serialization;

namespace ConectividadeHost;

/// <summary>
/// Native host da Conectividade Canadá.
///
/// Faz a ponte entre:
///   - o DISCO (pastas C:\sga\req e C:\sga\resp que o SGA usa), e
///   - a EXTENSÃO do Chrome (que executa a consulta na sessão logada).
///
/// Loop principal:
///   1. Um FileSystemWatcher vigia C:\sga\req por novos arquivos *.xml.
///   2. Para cada envio (Operacao 6), extrai placa/chassi/renavam e manda
///      { tipo:"consultar_selagem", id, placa, ... } para a extensão.
///   3. A extensão responde { tipo:"resultado_selagem", id, status, dados }.
///   4. Monta o ConsultaSelagemResposta e grava o XML em C:\sga\resp.
///
/// A comunicação com a extensão usa Native Messaging (stdin/stdout). O Chrome
/// inicia este processo automaticamente quando a extensão chama
/// connectNative(); não é necessário "abrir" nada.
/// </summary>
public static class Program
{
    // Pastas padrão. Podem ser sobrescritas por variáveis de ambiente
    // CONECT_REQ_DIR / CONECT_RESP_DIR (útil para testes).
    private static readonly string ReqDir =
        Environment.GetEnvironmentVariable("CONECT_REQ_DIR") ?? @"C:\sga\req";
    private static readonly string RespDir =
        Environment.GetEnvironmentVariable("CONECT_RESP_DIR") ?? @"C:\sga\resp";

    // Pedidos aguardando resposta da extensão: id -> arquivo de origem.
    private static readonly ConcurrentDictionary<string, PedidoPendente> Pendentes = new();

    // Arquivos de envio em processamento (dedupe do watcher / sobrescritas).
    // chave = caminho completo do arquivo em C:\sga\req; valor = quando começou.
    // Se um pedido travar (extensão morreu, etc.), destravamos após este tempo.
    private static readonly ConcurrentDictionary<string, DateTime> EmProcessamento = new();
    private static readonly TimeSpan TempoMaxProcessamento = TimeSpan.FromSeconds(90);

    public static void Main()
    {
        Log.Info($"Native host iniciado. REQ={ReqDir} RESP={RespDir}");

        Directory.CreateDirectory(ReqDir);
        Directory.CreateDirectory(RespDir);

        IniciarWatcher();

        // Processa arquivos que já estavam na pasta antes de abrir (fila pendente).
        foreach (var arquivo in Directory.GetFiles(ReqDir, "*.xml"))
            ProcessarEnvio(arquivo);

        // Loop de leitura das mensagens da extensão (bloqueante até o Chrome
        // fechar o stream).
        while (true)
        {
            JsonElement? msg;
            try
            {
                msg = NativeMessaging.ReadMessage();
            }
            catch (Exception ex)
            {
                Log.Error($"Erro ao ler mensagem da extensão: {ex.Message}");
                break;
            }

            if (msg == null)
            {
                Log.Info("Stream da extensão fechado. Encerrando host.");
                break;
            }

            try
            {
                TratarMensagemDaExtensao(msg.Value);
            }
            catch (Exception ex)
            {
                Log.Error($"Erro ao tratar mensagem: {ex.Message}");
            }
        }
    }

    // ---- Watcher das pastas -------------------------------------------------

    private static void IniciarWatcher()
    {
        var watcher = new FileSystemWatcher(ReqDir, "*.xml")
        {
            NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite,
            EnableRaisingEvents = true
        };

        // O SGA grava sempre <USUARIO>.XML (ex.: ADM.XML). Dependendo de como
        // ele escreve, o watcher pode disparar Created (arquivo novo) OU
        // Changed (sobrescrita de um arquivo de mesmo nome). Tratamos os dois.
        watcher.Created += (_, e) => ProcessarEnvio(e.FullPath);
        watcher.Changed += (_, e) => ProcessarEnvio(e.FullPath);
        // Mantém o watcher vivo pelo tempo de vida do processo.
        GC.KeepAlive(watcher);
    }

    // ---- Processa um arquivo de envio (C:\sga\req) --------------------------

    private static void ProcessarEnvio(string caminho)
    {
        // Como o nome é fixo por usuário (ADM.XML), o watcher pode disparar
        // mais de uma vez para o mesmo arquivo (Created + Changed) e o SGA pode
        // sobrescrevê-lo. Garante um único processamento por arquivo enquanto
        // a consulta correspondente não terminar.
        var agora = DateTime.UtcNow;
        if (!EmProcessamento.TryAdd(caminho, agora))
        {
            // Já havia um processamento em curso. Só ignora se for recente;
            // se ficou preso além do limite, destrava e reprocessa.
            if (EmProcessamento.TryGetValue(caminho, out var inicio)
                && agora - inicio < TempoMaxProcessamento)
            {
                Log.Info($"'{Path.GetFileName(caminho)}' já em processamento; ignorando disparo duplicado.");
                return;
            }
            Log.Error($"'{Path.GetFileName(caminho)}' estava preso; destravando e reprocessando.");
            EmProcessamento[caminho] = agora;
        }

        ConsultaSelagemEnvio envio;
        try
        {
            EsperarArquivoLiberar(caminho);
            envio = Desserializar<ConsultaSelagemEnvio>(caminho);
        }
        catch (Exception ex)
        {
            Log.Error($"Falha ao ler envio '{caminho}': {ex.Message}");
            EmProcessamento.TryRemove(caminho, out _);
            return;
        }

        Log.Info($"Envio lido: Operacao={envio.Operacao} Placa={envio.Placa} ({Path.GetFileName(caminho)})");

        // Começamos só com a Operação 6 (consulta de selagem por placa).
        if (envio.Operacao != 6)
        {
            Log.Info($"Operação {envio.Operacao} ainda não implementada. Ignorando por ora.");
            EmProcessamento.TryRemove(caminho, out _);
            return;
        }

        var id = Guid.NewGuid().ToString("N");
        Pendentes[id] = new PedidoPendente
        {
            ArquivoOrigem = caminho,
            Placa = envio.Placa
        };

        NativeMessaging.WriteMessage(new
        {
            tipo = "consultar_selagem",
            id,
            placa = envio.Placa ?? "",
            chassi = envio.Chassi ?? "",
            renavam = envio.Renavam ?? ""
        });

        Log.Info($"Pedido {id} enviado à extensão (placa {envio.Placa}).");
    }

    // ---- Trata a resposta da extensão --------------------------------------

    private static void TratarMensagemDaExtensao(JsonElement msg)
    {
        if (!msg.TryGetProperty("tipo", out var tipoEl)) return;
        var tipo = tipoEl.GetString();

        if (tipo == "pong")
        {
            Log.Info("pong recebido da extensão.");
            return;
        }

        if (tipo != "resultado_selagem") return;

        var id = msg.GetProperty("id").GetString() ?? "";
        if (!Pendentes.TryRemove(id, out var pendente))
        {
            Log.Error($"Resultado recebido para id desconhecido: {id}");
            return;
        }

        var status = msg.TryGetProperty("status", out var st) && st.GetBoolean();
        var message = msg.TryGetProperty("message", out var mm) ? mm.GetString() ?? "" : "";

        var resposta = new ConsultaSelagemResposta
        {
            Status = status,
            Message = message
        };

        if (status && msg.TryGetProperty("dados", out var dados) && dados.ValueKind == JsonValueKind.Object)
            PreencherDados(resposta, dados);

        GravarResposta(resposta, pendente);
    }

    private static void PreencherDados(ConsultaSelagemResposta r, JsonElement d)
    {
        if (d.TryGetProperty("veiculo", out var v))
        {
            r.Veiculo.Renavam = S(v, "renavam");
            r.Veiculo.Placa = S(v, "placa");
            r.Veiculo.Marca = S(v, "marca");
            r.Veiculo.Modelo = S(v, "modelo");
            r.Veiculo.Ano = S(v, "ano");
            r.Veiculo.Aro = S(v, "aro");
            r.Veiculo.Pneu = S(v, "pneu");
            r.Veiculo.Chassi = S(v, "chassi");
        }
        if (d.TryGetProperty("equipamento", out var e))
        {
            r.Equipamento.Marca = S(e, "marca");
            r.Equipamento.Modelo = S(e, "modelo");
            r.Equipamento.NroSerie = S(e, "nroSerie");
            r.Equipamento.ConstK = S(e, "constK");
            r.Equipamento.Redutor = S(e, "redutor");
        }
        if (d.TryGetProperty("selagem", out var s))
        {
            r.Selagem.DataSelagem = S(s, "dataSelagem");
            r.Selagem.PostoSelagem = S(s, "postoSelagem");
            r.Selagem.Protocolo = S(s, "protocolo");
            r.Selagem.SelosAdesivos = S(s, "selosAdesivos");
            r.Selagem.SelosPlasticos = S(s, "selosPlasticos");
            r.Selagem.Situacao = S(s, "situacao");
        }
        if (d.TryGetProperty("ensaio", out var en))
        {
            r.Ensaio.UFEnsaio = S(en, "ufEnsaio");
            r.Ensaio.PostoEnsaio = S(en, "postoEnsaio");
            r.Ensaio.DataEnsaio = S(en, "dataEnsaio");
            r.Ensaio.VelMax = S(en, "velMax");
            r.Ensaio.PadraoUtilizado = S(en, "padraoUtilizado");
        }
        if (d.TryGetProperty("resultado", out var re))
        {
            r.Resultado.UFResultado = S(re, "ufResultado");
            r.Resultado.RespTecnico = S(re, "respTecnico");
            r.Resultado.DataEmissao = S(re, "dataEmissao");
            r.Resultado.DataValidade = S(re, "dataValidade");
            r.Resultado.NumCertificado = S(re, "numCertificado");
            r.Resultado.TipoCertificado = S(re, "tipoCertificado");
        }
        if (d.TryGetProperty("gru", out var g))
        {
            r.GRU.NumGRU = S(g, "numGru");
            r.GRU.DataPag = S(g, "dataPag");
            r.GRU.ValorPag = S(g, "valorPag");
            r.GRU.ValorGru = S(g, "valorGru");
        }
    }

    private static string S(JsonElement obj, string prop) =>
        obj.TryGetProperty(prop, out var el) && el.ValueKind == JsonValueKind.String
            ? el.GetString() ?? ""
            : "";

    // ---- Grava o XML de resposta em C:\sga\resp -----------------------------

    private static void GravarResposta(ConsultaSelagemResposta resposta, PedidoPendente pendente)
    {
        // O SGA usa SEMPRE o mesmo nome por usuário nas duas pastas:
        //   C:\sga\req\ADM.XML  ->  C:\sga\resp\ADM.XML
        // Por isso o arquivo de resposta mantém EXATAMENTE o mesmo nome de
        // arquivo do envio (inclusive a extensão, ex.: .XML maiúsculo).
        var nomeArquivo = Path.GetFileName(pendente.ArquivoOrigem);
        var destino = Path.Combine(RespDir, nomeArquivo);

        try
        {
            Serializar(resposta, destino);
            Log.Info($"Resposta gravada: {destino} (STATUS={resposta.Status})");

            // Remove o arquivo de requisição já processado.
            try { File.Delete(pendente.ArquivoOrigem); } catch { /* best-effort */ }
        }
        catch (Exception ex)
        {
            Log.Error($"Falha ao gravar resposta '{destino}': {ex.Message}");
        }
        finally
        {
            // Libera o arquivo para um próximo pedido do mesmo usuário.
            EmProcessamento.TryRemove(pendente.ArquivoOrigem, out _);
        }
    }

    // ---- Serialização XML ---------------------------------------------------

    private static T Desserializar<T>(string caminho)
    {
        var ser = new XmlSerializer(typeof(T));
        using var fs = File.OpenRead(caminho);
        return (T)ser.Deserialize(fs)!;
    }

    private static void Serializar<T>(T obj, string caminho)
    {
        var ser = new XmlSerializer(typeof(T));
        var settings = new XmlWriterSettings { Indent = true, Encoding = new System.Text.UTF8Encoding(false) };

        // Namespace padrão "www.canainfo.com.br" sem prefixo (igual ao exemplo).
        var ns = new XmlSerializerNamespaces();
        ns.Add("", "www.canainfo.com.br");

        var tmp = caminho + ".tmp";
        using (var writer = XmlWriter.Create(tmp, settings))
            ser.Serialize(writer, obj, ns);

        // Escrita atômica: grava .tmp e renomeia, pra o SGA nunca ler pela metade.
        if (File.Exists(caminho)) File.Delete(caminho);
        File.Move(tmp, caminho);
    }

    private static void EsperarArquivoLiberar(string caminho, int tentativas = 20)
    {
        for (int i = 0; i < tentativas; i++)
        {
            try
            {
                using var fs = File.Open(caminho, FileMode.Open, FileAccess.Read, FileShare.None);
                return; // conseguiu abrir exclusivo => liberado
            }
            catch (IOException)
            {
                Thread.Sleep(100);
            }
        }
    }

    private sealed class PedidoPendente
    {
        public string ArquivoOrigem { get; set; } = "";
        public string Placa { get; set; } = "";
    }
}
