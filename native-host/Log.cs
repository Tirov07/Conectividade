namespace ConectividadeHost;

/// <summary>
/// Logger simples em arquivo, com RESET automático por tamanho.
///
/// ATENÇÃO: no Native Messaging o stdout é o canal de comunicação com o Chrome
/// (não pode ser usado para logs!) e o stderr pode interferir. Por isso os logs
/// vão para um arquivo em %LOCALAPPDATA%\CanadaConectividade\host.log.
///
/// O log serve só para conferir se o host está funcionando — não precisa de
/// histórico longo. Para não crescer indefinidamente com muitas consultas por
/// dia, quando o host.log passa de <see cref="TamanhoMaxBytes"/> ele é
/// simplesmente ZERADO e recomeça do zero (sem guardar arquivos antigos).
/// </summary>
public static class Log
{
    private static readonly string LogDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "CanadaConectividade");

    private static readonly string LogFile = Path.Combine(LogDir, "host.log");
    private static readonly object Lock = new();

    // Limite de tamanho; ao ultrapassar, o log é zerado. Pode ser sobrescrito
    // pela variável de ambiente CONECT_LOG_MAX_MB (padrão: 2 MB).
    private static readonly long TamanhoMaxBytes = LerLimiteMb() * 1024L * 1024L;

    public static void Info(string msg) => Escrever("INFO", msg);
    public static void Error(string msg) => Escrever("ERRO", msg);

    private static void Escrever(string nivel, string msg)
    {
        try
        {
            Directory.CreateDirectory(LogDir);
            var linha = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{nivel}] {msg}{Environment.NewLine}";
            lock (Lock)
            {
                ResetarSePassouDoLimite();
                File.AppendAllText(LogFile, linha);
            }
        }
        catch
        {
            // Nunca deixar o log derrubar o host.
        }
    }

    /// <summary>
    /// Se o host.log passou do limite, apaga tudo e recomeça do zero.
    /// Chamado já dentro do lock. Não guarda histórico antigo (de propósito:
    /// o log é só para verificar o funcionamento, não para auditoria).
    /// </summary>
    private static void ResetarSePassouDoLimite()
    {
        try
        {
            var info = new FileInfo(LogFile);
            if (info.Exists && info.Length >= TamanhoMaxBytes)
            {
                // Recomeça o arquivo, deixando uma marca de quando foi zerado.
                File.WriteAllText(
                    LogFile,
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [INFO] (log reiniciado ao atingir o limite de tamanho){Environment.NewLine}");
            }
        }
        catch
        {
            // Se falhar (arquivo em uso, etc.), no pior caso o log continua
            // crescendo — nunca derruba o host por causa disso.
        }
    }

    private static int LerLimiteMb()
    {
        var v = Environment.GetEnvironmentVariable("CONECT_LOG_MAX_MB");
        return int.TryParse(v, out var mb) && mb > 0 ? mb : 2;
    }
}
