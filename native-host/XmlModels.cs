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
