namespace ConectividadeHost;

/// <summary>
/// Logger simples em arquivo.
///
/// ATENÇÃO: no Native Messaging o stdout é o canal de comunicação com o Chrome
/// (não pode ser usado para logs!) e o stderr pode interferir. Por isso os logs
/// vão para um arquivo em %LOCALAPPDATA%\CanadaConectividade\host.log.
/// </summary>
public static class Log
{
    private static readonly string LogDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "CanadaConectividade");

    private static readonly string LogFile = Path.Combine(LogDir, "host.log");
    private static readonly object Lock = new();

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
                File.AppendAllText(LogFile, linha);
            }
        }
        catch
        {
            // Nunca deixar o log derrubar o host.
        }
    }
}
