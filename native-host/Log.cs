namespace ConectividadeHost;

/// <summary>
/// Logger simples em arquivo, COM ROTAÇÃO POR TAMANHO.
///
/// ATENÇÃO: no Native Messaging o stdout é o canal de comunicação com o Chrome
/// (não pode ser usado para logs!) e o stderr pode interferir. Por isso os logs
/// vão para um arquivo em %LOCALAPPDATA%\CanadaConectividade\host.log.
///
/// Como as consultas acontecem muitas vezes ao dia, o arquivo cresceria sem
/// parar. Para evitar isso, quando o host.log passa de <see cref="TamanhoMaxBytes"/>
/// ele é "rotacionado": host.log.1 vira host.log.2, host.log vira host.log.1 e
/// um host.log novo começa do zero. Mantemos no máximo <see cref="MaxArquivos"/>
/// arquivos antigos. Assim o total em disco é limitado e sempre há histórico
/// recente para depurar.
/// </summary>
public static class Log
{
    private static readonly string LogDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "CanadaConectividade");

    private static readonly string LogFile = Path.Combine(LogDir, "host.log");
    private static readonly object Lock = new();

    // Limite de tamanho por arquivo (5 MB) e quantos arquivos antigos manter.
    // Pode ser sobrescrito pela variável de ambiente CONECT_LOG_MAX_MB.
    private static readonly long TamanhoMaxBytes = LerLimiteMb() * 1024L * 1024L;
    private const int MaxArquivos = 2; // host.log.1 e host.log.2

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
                RotacionarSeNecessario();
                File.AppendAllText(LogFile, linha);
            }
        }
        catch
        {
            // Nunca deixar o log derrubar o host.
        }
    }

    /// <summary>
    /// Se o host.log passou do limite, renomeia em cadeia:
    ///   host.log.2 (apagado) &lt;- host.log.1 &lt;- host.log
    /// e deixa o host.log livre para recomeçar.
    /// Chamado já dentro do lock.
    /// </summary>
    private static void RotacionarSeNecessario()
    {
        try
        {
            var info = new FileInfo(LogFile);
            if (!info.Exists || info.Length < TamanhoMaxBytes)
                return;

            // Apaga o mais antigo (host.log.2).
            var maisAntigo = $"{LogFile}.{MaxArquivos}";
            if (File.Exists(maisAntigo)) File.Delete(maisAntigo);

            // Desloca os intermediários: host.log.1 -> host.log.2, etc.
            for (int i = MaxArquivos - 1; i >= 1; i--)
            {
                var origem = $"{LogFile}.{i}";
                var destino = $"{LogFile}.{i + 1}";
                if (File.Exists(origem)) File.Move(origem, destino, overwrite: true);
            }

            // host.log atual vira host.log.1.
            File.Move(LogFile, $"{LogFile}.1", overwrite: true);
        }
        catch
        {
            // Se a rotação falhar (arquivo em uso, etc.), no pior caso o log
            // continua crescendo — nunca derruba o host por causa disso.
        }
    }

    private static int LerLimiteMb()
    {
        var v = Environment.GetEnvironmentVariable("CONECT_LOG_MAX_MB");
        return int.TryParse(v, out var mb) && mb > 0 ? mb : 5;
    }
}
