using System.Xml.Serialization;

namespace ConectividadeHost;

// Modelos que espelham EXATAMENTE os XMLs trocados com o SGA.
// Namespace padrão de todos os elementos: "www.canainfo.com.br".

// ---------------------------------------------------------------------------
// ENVIO — arquivo que o SGA grava em C:\sga\req
// ---------------------------------------------------------------------------

[XmlRoot("ConsultaSelagemEnvio", Namespace = "www.canainfo.com.br")]
public class ConsultaSelagemEnvio
{
    [XmlElement("Operacao")]
    public int Operacao { get; set; }

    [XmlElement("Url")]
    public string Url { get; set; } = "";

    [XmlElement("Placa")]
    public string Placa { get; set; } = "";

    [XmlElement("Renavam")]
    public string Renavam { get; set; } = "";

    [XmlElement("Chassi")]
    public string Chassi { get; set; } = "";

    [XmlElement("ConsultaGRU")]
    public bool ConsultaGRU { get; set; }

    [XmlElement("UrlConsultaGRU")]
    public string UrlConsultaGRU { get; set; } = "";
}

// ---------------------------------------------------------------------------
// RETORNO — arquivo que o host grava em C:\sga\resp
// ---------------------------------------------------------------------------

[XmlRoot("ConsultaSelagemResposta", Namespace = "www.canainfo.com.br")]
public class ConsultaSelagemResposta
{
    // IMPORTANTE: o SGA grava e espera o STATUS como "True"/"False" com a
    // primeira letra MAIÚSCULA (veja o XML de retorno real). Um bool do .NET
    // seria serializado como "true"/"false" minúsculo e o SGA poderia não
    // reconhecer. Por isso guardamos como string e controlamos o valor exato.
    [XmlIgnore]
    public bool Status { get; set; }

    [XmlElement("STATUS")]
    public string StatusTexto
    {
        get => Status ? "True" : "False";
        set => Status = value?.Trim().Equals("true", StringComparison.OrdinalIgnoreCase) ?? false;
    }

    [XmlElement("Veiculo")]
    public VeiculoDto Veiculo { get; set; } = new();

    [XmlElement("Equipamento")]
    public EquipamentoDto Equipamento { get; set; } = new();

    [XmlElement("Selagem")]
    public SelagemDto Selagem { get; set; } = new();

    [XmlElement("Ensaio")]
    public EnsaioDto Ensaio { get; set; } = new();

    [XmlElement("Resultado")]
    public ResultadoDto Resultado { get; set; } = new();

    [XmlElement("GRU")]
    public GruDto GRU { get; set; } = new();

    [XmlElement("Message")]
    public string Message { get; set; } = "";
}

public class VeiculoDto
{
    [XmlElement("Renavam")] public string Renavam { get; set; } = "";
    [XmlElement("Placa")] public string Placa { get; set; } = "";
    [XmlElement("Marca")] public string Marca { get; set; } = "";
    [XmlElement("Modelo")] public string Modelo { get; set; } = "";
    [XmlElement("Ano")] public string Ano { get; set; } = "";
    [XmlElement("Aro")] public string Aro { get; set; } = "";
    [XmlElement("Pneu")] public string Pneu { get; set; } = "";
    [XmlElement("Chassi")] public string Chassi { get; set; } = "";
}

public class EquipamentoDto
{
    [XmlElement("Marca")] public string Marca { get; set; } = "";
    [XmlElement("Modelo")] public string Modelo { get; set; } = "";
    [XmlElement("NroSerie")] public string NroSerie { get; set; } = "";
    [XmlElement("ConstK")] public string ConstK { get; set; } = "";
    [XmlElement("Redutor")] public string Redutor { get; set; } = "";
}

public class SelagemDto
{
    [XmlElement("DataSelagem")] public string DataSelagem { get; set; } = "";
    [XmlElement("PostoSelagem")] public string PostoSelagem { get; set; } = "";
    [XmlElement("Protocolo")] public string Protocolo { get; set; } = "";
    [XmlElement("SelosAdesivos")] public string SelosAdesivos { get; set; } = "";
    [XmlElement("SelosPlasticos")] public string SelosPlasticos { get; set; } = "";
    [XmlElement("Situacao")] public string Situacao { get; set; } = "";
}

public class EnsaioDto
{
    [XmlElement("UFEnsaio")] public string UFEnsaio { get; set; } = "";
    [XmlElement("PostoEnsaio")] public string PostoEnsaio { get; set; } = "";
    [XmlElement("DataEnsaio")] public string DataEnsaio { get; set; } = "";
    [XmlElement("VelMax")] public string VelMax { get; set; } = "";
    [XmlElement("PadraoUtilizado")] public string PadraoUtilizado { get; set; } = "";
}

public class ResultadoDto
{
    [XmlElement("UFResultado")] public string UFResultado { get; set; } = "";
    [XmlElement("RespTecnico")] public string RespTecnico { get; set; } = "";
    [XmlElement("DataEmissao")] public string DataEmissao { get; set; } = "";
    [XmlElement("DataValidade")] public string DataValidade { get; set; } = "";
    [XmlElement("NumCertificado")] public string NumCertificado { get; set; } = "";
    [XmlElement("TipoCertificado")] public string TipoCertificado { get; set; } = "";
}

public class GruDto
{
    [XmlElement("NumGRU")] public string NumGRU { get; set; } = "";
    [XmlElement("DataPag")] public string DataPag { get; set; } = "";
    [XmlElement("ValorPag")] public string ValorPag { get; set; } = "";
    [XmlElement("ValorGru")] public string ValorGru { get; set; } = "";
}

// ===========================================================================
// OPERAÇÃO 13 — CONSULTA DE LOTE DE PLACAS
// ===========================================================================
// Consulta várias placas de uma vez. Reusa o MESMO fluxo da consulta de
// selagem (tela /selagens/consultar -> "ver" -> detalhes), mas para cada placa
// do lote, e monta um retorno com N blocos <Consulta>.
//
// Observação: o <UrlConsulta> (tela de certificados) do envio é IGNORADO de
// propósito — essa tela passou a exigir captcha por placa. A DataValidade sai
// do campo "Informações do resultado da verificação" da tela de selagem.

// ---- ENVIO (C:\sga\req) ---------------------------------------------------

[XmlRoot("ConsultaLotePlacasEnvio", Namespace = "www.canainfo.com.br")]
public class ConsultaLotePlacasEnvio
{
    [XmlElement("Operacao")]
    public int Operacao { get; set; }

    [XmlElement("Url")]
    public string Url { get; set; } = "";

    // Tela de certificados — ignorada (hoje exige captcha). Mantida só para
    // compatibilidade com o XML que o SGA gera.
    [XmlElement("UrlConsulta")]
    public string UrlConsulta { get; set; } = "";

    [XmlElement("Lote")]
    public List<LoteItemEnvio> Lote { get; set; } = new();
}

public class LoteItemEnvio
{
    [XmlAttribute("Id")]
    public int Id { get; set; }

    [XmlElement("Placa")]
    public string Placa { get; set; } = "";

    [XmlElement("CodCliente")]
    public string CodCliente { get; set; } = "";
}

// ---- RETORNO (C:\sga\resp) ------------------------------------------------

[XmlRoot("ConsultaLotePlacasResposta", Namespace = "www.canainfo.com.br")]
public class ConsultaLotePlacasResposta
{
    [XmlElement("Consulta")]
    public List<LoteItemResposta> Consultas { get; set; } = new();

    // STATUS e Message GERAIS do lote (vêm depois dos <Consulta> no XML).
    [XmlIgnore]
    public bool Status { get; set; }

    [XmlElement("STATUS")]
    public string StatusTexto
    {
        get => Status ? "True" : "False";
        set => Status = value?.Trim().Equals("true", StringComparison.OrdinalIgnoreCase) ?? false;
    }

    [XmlElement("Message")]
    public string Message { get; set; } = "";
}

public class LoteItemResposta
{
    [XmlAttribute("Id")]
    public int Id { get; set; }

    [XmlElement("CodCliente")]
    public string CodCliente { get; set; } = "";

    [XmlElement("Veiculo")]
    public LoteVeiculoDto Veiculo { get; set; } = new();

    [XmlElement("Selagem")]
    public LoteSelagemDto Selagem { get; set; } = new();

    [XmlElement("Ensaio")]
    public EnsaioDto Ensaio { get; set; } = new();

    [XmlElement("Equipamento")]
    public LoteEquipamentoDto Equipamento { get; set; } = new();

    [XmlElement("Certificado1")]
    public Certificado1Dto Certificado1 { get; set; } = new();

    [XmlElement("Certificado2")]
    public Certificado2Dto Certificado2 { get; set; } = new();

    [XmlElement("Resultado")]
    public ResultadoDto Resultado { get; set; } = new();

    // Status individual DESTA placa (True se achou dados, False se não).
    [XmlIgnore]
    public bool StatusConsulta { get; set; }

    [XmlElement("StatusConsulta")]
    public string StatusConsultaTexto
    {
        get => StatusConsulta ? "True" : "False";
        set => StatusConsulta = value?.Trim().Equals("true", StringComparison.OrdinalIgnoreCase) ?? false;
    }
}

// O veículo do lote NÃO traz Chassi (diferente da Op. 6).
public class LoteVeiculoDto
{
    [XmlElement("Renavam")] public string Renavam { get; set; } = "";
    [XmlElement("Placa")] public string Placa { get; set; } = "";
    [XmlElement("Marca")] public string Marca { get; set; } = "";
    [XmlElement("Modelo")] public string Modelo { get; set; } = "";
    [XmlElement("Ano")] public string Ano { get; set; } = "";
    [XmlElement("Aro")] public string Aro { get; set; } = "";
    [XmlElement("Pneu")] public string Pneu { get; set; } = "";
}

// A selagem do lote traz só Data e Posto.
public class LoteSelagemDto
{
    [XmlElement("DataSelagem")] public string DataSelagem { get; set; } = "";
    [XmlElement("PostoSelagem")] public string PostoSelagem { get; set; } = "";
}

// O equipamento do lote traz só Marca/Modelo/NroSerie (sem ConstK/Redutor).
public class LoteEquipamentoDto
{
    [XmlElement("Marca")] public string Marca { get; set; } = "";
    [XmlElement("Modelo")] public string Modelo { get; set; } = "";
    [XmlElement("NroSerie")] public string NroSerie { get; set; } = "";
}

public class Certificado1Dto
{
    [XmlElement("NroDocto1")] public string NroDocto1 { get; set; } = "";
    [XmlElement("DataDocto1")] public string DataDocto1 { get; set; } = "";
    [XmlElement("ValidadeDocto1")] public string ValidadeDocto1 { get; set; } = "";
    [XmlElement("Tipo1")] public string Tipo1 { get; set; } = "";
}

public class Certificado2Dto
{
    [XmlElement("NroDocto2")] public string NroDocto2 { get; set; } = "";
    [XmlElement("DataDocto2")] public string DataDocto2 { get; set; } = "";
    [XmlElement("ValidadeDocto2")] public string ValidadeDocto2 { get; set; } = "";
    [XmlElement("Tipo2")] public string Tipo2 { get; set; } = "";
}

// ===========================================================================
// OPERAÇÃO 3 — INFORMAR SELAGEM (ESCRITA)
// ===========================================================================
// Grava uma selagem oficial no Inmetro e consome uma GRU real. Por isso o
// preenchimento é feito com MUITO cuidado e, no começo, SEM clicar em "Enviar"
// (modo seguro). Fluxo na tela /selagens/informar:
//   Etapa 1: buscar o veículo (placa/renavam/chassi/GRU)
//   Etapa 2: preencher marca -> modelo -> nº série (cada um dispara AJAX e
//            recarrega o form), depois ConstK, TipoRegistro, Escala, Ano, PAM,
//            selos (plásticos/adesivos com tipo+número), radio do selo da
//            caixa/sensor, velocidade e o atesto; então Enviar.

[XmlRoot("InformarSelagemEnvio", Namespace = "www.canainfo.com.br")]
public class InformarSelagemEnvio
{
    [XmlElement("Operacao")] public int Operacao { get; set; }
    [XmlElement("Url")] public string Url { get; set; } = "";
    [XmlElement("UrlConsulta")] public string UrlConsulta { get; set; } = "";
    [XmlElement("TipoPosto")] public string TipoPosto { get; set; } = "";
    [XmlElement("UrlDownload")] public string UrlDownload { get; set; } = "";

    [XmlElement("Equipamento")] public InformarEquipamentoDto Equipamento { get; set; } = new();
    [XmlElement("Veiculo")] public InformarVeiculoDto Veiculo { get; set; } = new();
    [XmlElement("Selagem")] public InformarSelagemDto Selagem { get; set; } = new();
}

public class InformarEquipamentoDto
{
    [XmlElement("NroSerie")] public string NroSerie { get; set; } = "";
    [XmlElement("Marca")] public string Marca { get; set; } = "";
    [XmlElement("Modelo")] public string Modelo { get; set; } = "";
    [XmlElement("Redutor")] public string Redutor { get; set; } = "";
    [XmlElement("ConstK")] public string ConstK { get; set; } = "";
    [XmlElement("TipoRegistro")] public string TipoRegistro { get; set; } = "";
    [XmlElement("EscalaVelocidade")] public string EscalaVelocidade { get; set; } = "";
    [XmlElement("AnoInstrumento")] public string AnoInstrumento { get; set; } = "";
    [XmlElement("PAM")] public string PAM { get; set; } = "";
}

public class InformarVeiculoDto
{
    [XmlElement("Placa")] public string Placa { get; set; } = "";
    [XmlElement("Renavam")] public string Renavam { get; set; } = "";
    [XmlElement("Chassi")] public string Chassi { get; set; } = "";
    [XmlElement("Pneu")] public string Pneu { get; set; } = "";
    [XmlElement("Aro")] public string Aro { get; set; } = "";
}

public class InformarSelagemDto
{
    [XmlElement("GRUSelagem")] public string GRUSelagem { get; set; } = "";
    [XmlElement("Velocidade")] public string Velocidade { get; set; } = "";
    // Selos vêm como listas separadas por espaços, cada selo = tipo(3)+numero.
    // Ex.: "180H012259785  180H012259798" -> dois adesivos.
    [XmlElement("SelosAdesivos")] public string SelosAdesivos { get; set; } = "";
    [XmlElement("SelosPlasticos")] public string SelosPlasticos { get; set; } = "";
    [XmlElement("SeloPlasticoCaixaSensor")] public string SeloPlasticoCaixaSensor { get; set; } = "";
}

// ---- RETORNO da Op. 3 (C:\sga\resp) ---------------------------------------
// Confirmado por exemplo real de ERRO (GRU já utilizada). O retorno de SUCESSO
// (com protocolo?) ainda será enviado pelo cliente para calibrar os campos.

[XmlRoot("InformarSelagemResposta", Namespace = "www.canainfo.com.br")]
public class InformarSelagemResposta
{
    [XmlIgnore]
    public bool Status { get; set; }

    [XmlElement("STATUS")]
    public string StatusTexto
    {
        get => Status ? "True" : "False";
        set => Status = value?.Trim().Equals("true", StringComparison.OrdinalIgnoreCase) ?? false;
    }

    // Protocolo e Situacao só aparecem no SUCESSO (ex.: Protocolo gerado,
    // Situacao "Ativa"). No ERRO o XML traz apenas STATUS + Message, então
    // ShouldSerialize omite esses elementos quando vazios.
    [XmlElement("Protocolo")]
    public string Protocolo { get; set; } = "";
    public bool ShouldSerializeProtocolo() => !string.IsNullOrEmpty(Protocolo);

    [XmlElement("Situacao")]
    public string Situacao { get; set; } = "";
    public bool ShouldSerializeSituacao() => !string.IsNullOrEmpty(Situacao);

    [XmlElement("Message")]
    public string Message { get; set; } = "";
}
