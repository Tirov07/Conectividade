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
    // Tempo de "destravamento" de segurança: só reprocessa um mesmo arquivo se
    // ele ficou preso além disto (ex.: host travou). Precisa ser MAIOR que o
    // pior lote possível (Op. 13 com ~500 placas × ~10s ≈ 83 min), senão o
    // watcher poderia re-enfileirar o lote no meio e causar retrabalho enorme.
    private static readonly TimeSpan TempoMaxProcessamento = TimeSpan.FromHours(3);

    // Fila de arquivos de envio a processar, UM DE CADA VEZ. A extensão só faz
    // uma consulta por vez (mesma aba do Chrome), então serializamos aqui.
    private static readonly BlockingCollection<string> FilaEnvios = new();

    // Sincroniza a espera da resposta da extensão para o pedido atual.
    private static readonly object RespostaLock = new();
    private static ResultadoConsulta? RespostaBrutaAtual;
    private static ManualResetEventSlim? SinalResposta;

    // Arquivos processados há pouco: usado para ignorar eventos "fantasma" do
    // watcher que chegam logo depois de deletarmos o arquivo de requisição
    // (o próprio delete/escrita pode gerar um evento Changed atrasado).
    // chave = caminho; valor = quando terminamos de processar.
    private static readonly ConcurrentDictionary<string, DateTime> ProcessadosRecentemente = new();
    private static readonly TimeSpan PeriodoDeGraca = TimeSpan.FromSeconds(5);

    public static void Main()
    {
        Log.Info($"Native host iniciado. REQ={ReqDir} RESP={RespDir}");

        Directory.CreateDirectory(ReqDir);
        Directory.CreateDirectory(RespDir);

        IniciarWatcher();

        // Thread que consome a fila e processa um envio por vez.
        var consumidora = new Thread(ConsumirFila) { IsBackground = true };
        consumidora.Start();

        // Enfileira arquivos que já estavam na pasta antes de abrir.
        foreach (var arquivo in Directory.GetFiles(ReqDir, "*.xml"))
            EnfileirarEnvio(arquivo);

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
        watcher.Created += (_, e) => EnfileirarEnvio(e.FullPath);
        watcher.Changed += (_, e) => EnfileirarEnvio(e.FullPath);
        // Mantém o watcher vivo pelo tempo de vida do processo.
        GC.KeepAlive(watcher);
    }

    /// <summary>
    /// Coloca um arquivo de envio na fila de processamento.
    ///
    /// Por que uma fila: a extensão só consegue fazer UMA consulta por vez
    /// (usa a mesma aba do Chrome). Quando o SGA grava dois arquivos quase
    /// juntos (ex.: ADM.xml e o temporário ADMTemp.xml), mandar os dois de uma
    /// vez à extensão fazia um deles dar timeout. A fila garante processamento
    /// sequencial: um pedido só começa quando o anterior termina.
    ///
    /// Também ignoramos arquivos temporários do SGA (ex.: *Temp.xml), que são
    /// apenas rascunhos enquanto ele escreve o arquivo real.
    /// </summary>
    private static void EnfileirarEnvio(string caminho)
    {
        var nome = Path.GetFileNameWithoutExtension(caminho);

        // Ignora o arquivo temporário do SGA (rascunho). O arquivo "oficial"
        // é o <USUARIO>.xml (ex.: ADM.xml), não o <USUARIO>Temp.xml.
        if (nome.EndsWith("Temp", StringComparison.OrdinalIgnoreCase))
        {
            Log.Info($"Ignorando arquivo temporário do SGA: {Path.GetFileName(caminho)}");
            return;
        }

        // Ignora eventos "fantasma" logo após um processamento (o próprio
        // delete/escrita gera um evento Changed atrasado para o mesmo nome).
        var agora = DateTime.UtcNow;
        if (ProcessadosRecentemente.TryGetValue(caminho, out var quando)
            && agora - quando < PeriodoDeGraca)
        {
            return;
        }

        // Limpeza leve: remove entradas de período de graça já vencidas.
        foreach (var kv in ProcessadosRecentemente)
            if (agora - kv.Value > PeriodoDeGraca)
                ProcessadosRecentemente.TryRemove(kv.Key, out _);

        FilaEnvios.Add(caminho);
    }

    // ---- Consumidor da fila: processa um envio de cada vez ------------------

    private static void ConsumirFila()
    {
        foreach (var caminho in FilaEnvios.GetConsumingEnumerable())
        {
            try
            {
                ProcessarEnvio(caminho);
            }
            catch (Exception ex)
            {
                Log.Error($"Erro ao processar '{caminho}': {ex.Message}");
            }
        }
    }

    // ---- Processa um arquivo de envio (C:\sga\req) --------------------------
    // Roda na thread da fila. É BLOQUEANTE: envia o pedido à extensão e espera
    // a resposta (ou timeout) antes de retornar, garantindo que o próximo item
    // da fila só comece depois que este terminar.

    private static void ProcessarEnvio(string caminho)
    {
        // Dedupe: o watcher pode disparar Created + Changed para o mesmo arquivo.
        var agora = DateTime.UtcNow;
        if (!EmProcessamento.TryAdd(caminho, agora))
        {
            if (EmProcessamento.TryGetValue(caminho, out var inicio)
                && agora - inicio < TempoMaxProcessamento)
            {
                Log.Info($"'{Path.GetFileName(caminho)}' já em processamento; ignorando disparo duplicado.");
                return;
            }
            Log.Error($"'{Path.GetFileName(caminho)}' estava preso; destravando e reprocessando.");
            EmProcessamento[caminho] = agora;
        }

        try
        {
            // O arquivo pode ter sido deletado entre o enfileiramento e agora
            // (ex.: evento duplicado após o delete do processamento anterior).
            if (!File.Exists(caminho))
            {
                return; // nada a fazer; sai em silêncio.
            }

            // Detecta a operação pela tag raiz do XML, sem assumir o tipo.
            int operacao;
            string rootTag;
            try
            {
                EsperarArquivoLiberar(caminho);
                (operacao, rootTag) = DetectarOperacao(caminho);
            }
            catch (FileNotFoundException)
            {
                return; // sumiu no meio do caminho; ignora sem poluir o log.
            }
            catch (Exception ex)
            {
                Log.Error($"Falha ao ler envio '{caminho}': {ex.Message}");
                return;
            }

            // Roteia por operação.
            if (rootTag == "ConsultaLotePlacasEnvio" || operacao == 13)
            {
                ProcessarLote(caminho);
            }
            else if (rootTag == "ConsultaSelagemEnvio" || operacao == 6)
            {
                ProcessarSelagem(caminho);
            }
            else
            {
                Log.Info($"Operação {operacao} (raiz '{rootTag}') ainda não implementada. Ignorando por ora.");
            }
        }
        finally
        {
            EmProcessamento.TryRemove(caminho, out _);
            // Marca o período de graça: eventos do watcher para este mesmo
            // arquivo nos próximos segundos (gerados pelo delete/escrita) serão
            // ignorados no EnfileirarEnvio.
            ProcessadosRecentemente[caminho] = DateTime.UtcNow;
        }
    }

    /// <summary>
    /// Descobre a operação e a tag raiz do XML sem assumir o tipo concreto.
    /// </summary>
    private static (int operacao, string rootTag) DetectarOperacao(string caminho)
    {
        var doc = new XmlDocument();
        doc.Load(caminho);
        var root = doc.DocumentElement?.LocalName ?? "";
        var opNós = doc.GetElementsByTagName("Operacao");
        int op = opNós.Count > 0 && int.TryParse((opNós[0]?.InnerText ?? "").Trim(), out var n) ? n : 0;
        return (op, root);
    }

    /// <summary>
    /// Consulta UMA placa via extensão e devolve o resultado (status + dados
    /// brutos em JSON). É o tijolo reutilizado tanto pela Op. 6 (uma placa)
    /// quanto pela Op. 13 (lote). BLOQUEANTE: espera a extensão responder.
    /// </summary>
    private static ResultadoConsulta ConsultarUmaPlaca(string placa, string renavam, string chassi, int timeoutSeg = 75)
    {
        var sinal = new ManualResetEventSlim(false);
        lock (RespostaLock)
        {
            RespostaBrutaAtual = null;
            SinalResposta = sinal;
        }

        var id = Guid.NewGuid().ToString("N");
        Pendentes[id] = new PedidoPendente { ArquivoOrigem = "", Placa = placa };

        NativeMessaging.WriteMessage(new
        {
            tipo = "consultar_selagem",
            id,
            placa = placa ?? "",
            chassi = chassi ?? "",
            renavam = renavam ?? ""
        });

        bool respondeu = sinal.Wait(TimeSpan.FromSeconds(timeoutSeg));

        ResultadoConsulta resultado;
        lock (RespostaLock)
        {
            SinalResposta = null;
            resultado = RespostaBrutaAtual ?? new ResultadoConsulta
            {
                Status = false,
                Message = "Tempo esgotado aguardando a resposta da consulta."
            };
        }
        Pendentes.TryRemove(id, out _);

        if (!respondeu)
            Log.Error($"Consulta da placa '{placa}' expirou sem resposta da extensão.");

        return resultado;
    }

    // ---- Operação 6: consulta de selagem (uma placa) ------------------------

    private static void ProcessarSelagem(string caminho)
    {
        var envio = LerEnvio(caminho);
        Log.Info($"Op.6 Selagem | {DescreverCriterios(envio)} ({Path.GetFileName(caminho)})");

        var r = ConsultarUmaPlaca(envio.Placa, envio.Renavam, envio.Chassi);

        var resposta = new ConsultaSelagemResposta { Status = r.Status, Message = r.Message };
        if (r.Status && r.Dados.HasValue)
            PreencherDados(resposta, r.Dados.Value);

        GravarResposta(resposta, new PedidoPendente { ArquivoOrigem = caminho, Placa = envio.Placa ?? "" });
    }

    // ---- Operação 13: consulta de lote de placas ----------------------------

    private static void ProcessarLote(string caminho)
    {
        var envio = LerLote(caminho);
        Log.Info($"Op.13 Lote | {envio.Lote.Count} placas ({Path.GetFileName(caminho)})");

        var resposta = new ConsultaLotePlacasResposta { Status = true };
        int ok = 0, falhou = 0, n = 0;
        int total = envio.Lote.Count;

        foreach (var item in envio.Lote)
        {
            n++;
            Log.Info($"Op.13 consultando placa {n}/{total}: {item.Placa} (cliente {item.CodCliente})");

            // Avisa a extensão para mostrar o progresso (ex.: 30/120) na tela.
            EnviarProgressoLote(n, total, item.Placa);

            var r = ConsultarUmaPlaca(item.Placa, "", "");
            var consulta = MontarItemLote(item, r);
            resposta.Consultas.Add(consulta);

            if (consulta.StatusConsulta) ok++; else falhou++;
        }

        // Sinaliza que o lote terminou (a extensão esconde o indicador).
        EnviarProgressoLote(total, total, "", concluido: true);

        resposta.Message = $"CONSULTA DE LOTE FINALIZADA. {ok} ok, {falhou} sem dados.";
        Log.Info($"Op.13 lote concluído: {ok} ok, {falhou} sem dados de {total}.");

        GravarRespostaLote(resposta, caminho);
    }

    /// <summary>
    /// Envia à extensão o progresso do lote para exibição na tela (overlay +
    /// badge). É só informativo; não espera resposta.
    /// </summary>
    private static void EnviarProgressoLote(int atual, int total, string placa, bool concluido = false)
    {
        try
        {
            NativeMessaging.WriteMessage(new
            {
                tipo = "progresso_lote",
                atual,
                total,
                placa = placa ?? "",
                concluido
            });
        }
        catch { /* progresso é best-effort; nunca derruba o lote */ }
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
        if (!Pendentes.ContainsKey(id))
        {
            Log.Error($"Resultado recebido para id desconhecido/expirado: {id}");
            return;
        }

        var status = msg.TryGetProperty("status", out var st) && st.GetBoolean();
        var message = msg.TryGetProperty("message", out var mm) ? mm.GetString() ?? "" : "";
        var naoEncontrado = msg.TryGetProperty("naoEncontrado", out var ne) && ne.ValueKind == JsonValueKind.True;

        // Quando o site respondeu "Nenhum resultado encontrado.", padronizamos a
        // Message no formato EXATO que o SGA espera (texto em maiúsculas, com o
        // "×" do botão fechar e a quebra de linha), independente de variações
        // de capitalização da página.
        if (naoEncontrado)
            message = "×\nNENHUM RESULTADO ENCONTRADO.";

        var resultado = new ResultadoConsulta { Status = status, Message = message };
        if (status && msg.TryGetProperty("dados", out var dados) && dados.ValueKind == JsonValueKind.Object)
            resultado.Dados = dados.Clone(); // guarda o JSON bruto para quem consumir

        // Entrega o resultado para a thread da fila (que está esperando) e a
        // acorda. A montagem/gravação do XML é feita lá, de forma sequencial.
        lock (RespostaLock)
        {
            RespostaBrutaAtual = resultado;
            SinalResposta?.Set();
        }
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
            var placaRet = resposta.Veiculo?.Placa ?? "";
            var detalhe = resposta.Status
                ? $"Placa={placaRet} Protocolo={resposta.Selagem?.Protocolo}"
                : $"Message={resposta.Message}";
            Log.Info($"Resposta gravada: {destino} (STATUS={resposta.Status} | {detalhe})");

            // Remove o arquivo de requisição já processado.
            try { File.Delete(pendente.ArquivoOrigem); } catch { /* best-effort */ }
        }
        catch (Exception ex)
        {
            Log.Error($"Falha ao gravar resposta '{destino}': {ex.Message}");
        }
        // (O EmProcessamento é liberado no finally de ProcessarEnvio.)
    }

    // ---- Serialização XML ---------------------------------------------------

    /// <summary>
    /// Lê o XML de envio do SGA de forma TOLERANTE ao formato legado.
    ///
    /// O SGA gera um XML que o XmlSerializer padrão do .NET rejeita por dois
    /// motivos:
    ///   1. O namespace é "www.canainfo.com.br" (sem http://), o que confunde
    ///      o casamento rígido de elementos.
    ///   2. Campos booleanos vêm como "True"/"False" (maiúsculo), mas o XML
    ///      padrão exige "true"/"false" minúsculo — isso quebrava na leitura
    ///      do elemento <ConsultaGRU>.
    ///
    /// Por isso lemos os campos manualmente por NOME LOCAL do elemento
    /// (ignorando o namespace) e tratamos o booleano de forma case-insensitive.
    /// </summary>
    private static ConsultaSelagemEnvio LerEnvio(string caminho)
    {
        var doc = new XmlDocument();
        doc.Load(caminho);

        string Pegar(string nomeLocal)
        {
            // Busca o elemento pelo nome local, independente do namespace.
            var nós = doc.GetElementsByTagName(nomeLocal);
            return nós.Count > 0 ? (nós[0]?.InnerText ?? "").Trim() : "";
        }

        bool PegarBool(string nomeLocal)
        {
            var v = Pegar(nomeLocal);
            return v.Equals("true", StringComparison.OrdinalIgnoreCase)
                || v.Equals("1", StringComparison.OrdinalIgnoreCase)
                || v.Equals("sim", StringComparison.OrdinalIgnoreCase);
        }

        int PegarInt(string nomeLocal)
        {
            var v = Pegar(nomeLocal);
            return int.TryParse(v, out var n) ? n : 0;
        }

        return new ConsultaSelagemEnvio
        {
            Operacao = PegarInt("Operacao"),
            Url = Pegar("Url"),
            Placa = Pegar("Placa"),
            Renavam = Pegar("Renavam"),
            Chassi = Pegar("Chassi"),
            ConsultaGRU = PegarBool("ConsultaGRU"),
            UrlConsultaGRU = Pegar("UrlConsultaGRU")
        };
    }

    /// <summary>
    /// Monta uma descrição legível dos critérios de busca para o log, mostrando
    /// apenas os campos efetivamente preenchidos (placa, renavam e/ou chassi).
    /// Assim o log deixa claro se a consulta foi por placa OU por renavam.
    /// </summary>
    private static string DescreverCriterios(ConsultaSelagemEnvio e)
    {
        var partes = new List<string>();
        if (!string.IsNullOrWhiteSpace(e.Placa)) partes.Add($"Placa={e.Placa}");
        if (!string.IsNullOrWhiteSpace(e.Renavam)) partes.Add($"Renavam={e.Renavam}");
        if (!string.IsNullOrWhiteSpace(e.Chassi)) partes.Add($"Chassi={e.Chassi}");
        return partes.Count > 0 ? string.Join(" ", partes) : "(sem criterios)";
    }

    // ---- Leitura/montagem/gravação do LOTE (Op. 13) -------------------------

    /// <summary>
    /// Lê o envio de lote (ConsultaLotePlacasEnvio) de forma tolerante ao
    /// formato legado do SGA (namespace sem http://).
    /// </summary>
    private static ConsultaLotePlacasEnvio LerLote(string caminho)
    {
        var doc = new XmlDocument();
        doc.Load(caminho);

        string Pegar(XmlNode? no, string nomeLocal)
        {
            if (no == null) return "";
            foreach (XmlNode filho in no.ChildNodes)
                if (filho.LocalName == nomeLocal)
                    return (filho.InnerText ?? "").Trim();
            return "";
        }

        var envio = new ConsultaLotePlacasEnvio
        {
            Operacao = 13
        };

        foreach (XmlNode no in doc.GetElementsByTagName("Lote"))
        {
            var id = 0;
            if (no.Attributes?["Id"]?.Value is string s) int.TryParse(s, out id);
            envio.Lote.Add(new LoteItemEnvio
            {
                Id = id,
                Placa = Pegar(no, "Placa"),
                CodCliente = Pegar(no, "CodCliente")
            });
        }

        return envio;
    }

    /// <summary>
    /// Monta um item de resposta do lote a partir do resultado cru da consulta
    /// de uma placa. Mapeia os mesmos dados da Op. 6 para o layout do lote.
    /// </summary>
    private static LoteItemResposta MontarItemLote(LoteItemEnvio item, ResultadoConsulta r)
    {
        var c = new LoteItemResposta
        {
            Id = item.Id,
            CodCliente = item.CodCliente,
            StatusConsulta = r.Status
        };
        // Placa sempre volta (mesmo sem dados, para o SGA casar pelo Id/placa).
        c.Veiculo.Placa = item.Placa;

        if (!r.Status || !r.Dados.HasValue)
            return c; // sem dados: StatusConsulta=False, campos vazios.

        var d = r.Dados.Value;

        if (d.TryGetProperty("veiculo", out var v))
        {
            c.Veiculo.Renavam = S(v, "renavam");
            c.Veiculo.Placa = S(v, "placa").Length > 0 ? S(v, "placa") : item.Placa;
            c.Veiculo.Marca = S(v, "marca");
            c.Veiculo.Modelo = S(v, "modelo");
            c.Veiculo.Ano = S(v, "ano");
            c.Veiculo.Aro = S(v, "aro");
            c.Veiculo.Pneu = S(v, "pneu");
        }
        if (d.TryGetProperty("selagem", out var s))
        {
            c.Selagem.DataSelagem = S(s, "dataSelagem");
            c.Selagem.PostoSelagem = S(s, "postoSelagem");
        }
        if (d.TryGetProperty("ensaio", out var en))
        {
            c.Ensaio.UFEnsaio = S(en, "ufEnsaio");
            c.Ensaio.PostoEnsaio = S(en, "postoEnsaio");
            c.Ensaio.DataEnsaio = S(en, "dataEnsaio");
            c.Ensaio.VelMax = S(en, "velMax");
            c.Ensaio.PadraoUtilizado = S(en, "padraoUtilizado");
        }
        if (d.TryGetProperty("equipamento", out var eq))
        {
            c.Equipamento.Marca = S(eq, "marca");
            c.Equipamento.Modelo = S(eq, "modelo");
            c.Equipamento.NroSerie = S(eq, "nroSerie");
        }
        if (d.TryGetProperty("resultado", out var re))
        {
            c.Resultado.UFResultado = S(re, "ufResultado");
            c.Resultado.RespTecnico = S(re, "respTecnico");
            c.Resultado.DataEmissao = S(re, "dataEmissao");
            c.Resultado.DataValidade = S(re, "dataValidade");
            c.Resultado.NumCertificado = S(re, "numCertificado");
            c.Resultado.TipoCertificado = S(re, "tipoCertificado");

            // Certificado1 e Certificado2 — hoje repetem a mesma informação do
            // resultado (confirmar o porquê com o responsável da Conectividade
            // atual; por ora seguimos o padrão de repetir).
            c.Certificado1.NroDocto1 = c.Resultado.NumCertificado;
            c.Certificado1.Tipo1 = c.Resultado.TipoCertificado;
            c.Certificado2.NroDocto2 = c.Resultado.NumCertificado;
            c.Certificado2.Tipo2 = c.Resultado.TipoCertificado;
        }

        return c;
    }

    private static void GravarRespostaLote(ConsultaLotePlacasResposta resposta, string caminhoOrigem)
    {
        var nomeArquivo = Path.GetFileName(caminhoOrigem);
        var destino = Path.Combine(RespDir, nomeArquivo);
        try
        {
            SerializarLote(resposta, destino);
            Log.Info($"Resposta de lote gravada: {destino} ({resposta.Consultas.Count} placas)");
            try { File.Delete(caminhoOrigem); } catch { /* best-effort */ }
        }
        catch (Exception ex)
        {
            Log.Error($"Falha ao gravar resposta de lote '{destino}': {ex.Message}");
        }
    }

    private static void SerializarLote(ConsultaLotePlacasResposta obj, string caminho)
    {
        var ser = new XmlSerializer(typeof(ConsultaLotePlacasResposta));
        var settings = new XmlWriterSettings { Indent = true, Encoding = new System.Text.UTF8Encoding(false) };
        var ns = new XmlSerializerNamespaces();
        ns.Add("", "www.canainfo.com.br");

        var tmp = caminho + ".tmp";
        using (var writer = XmlWriter.Create(tmp, settings))
            ser.Serialize(writer, obj, ns);

        if (File.Exists(caminho)) File.Delete(caminho);
        File.Move(tmp, caminho);
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

    /// <summary>
    /// Resultado cru de UMA consulta à extensão: status, mensagem e o bloco
    /// de dados em JSON (quando houve sucesso). Reaproveitado por Op. 6 e 13.
    /// </summary>
    private sealed class ResultadoConsulta
    {
        public bool Status { get; set; }
        public string Message { get; set; } = "";
        public JsonElement? Dados { get; set; }
    }
}
