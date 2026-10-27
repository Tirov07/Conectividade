using System.Text;
using System.Text.Json;

namespace ConectividadeHost;

/// <summary>
/// Implementa o protocolo de Native Messaging do Chrome sobre stdin/stdout.
///
/// Cada mensagem é precedida por 4 bytes (uint32, byte order nativo da máquina =
/// little-endian no Windows) com o tamanho em bytes do JSON UTF-8 que vem a
/// seguir. Referência: documentação oficial de Native Messaging do Chrome.
/// </summary>
public static class NativeMessaging
{
    private static readonly Stream StdIn = Console.OpenStandardInput();
    private static readonly Stream StdOut = Console.OpenStandardOutput();
    private static readonly object WriteLock = new();

    /// <summary>
    /// Lê uma mensagem do Chrome. Retorna null quando o stream fecha
    /// (Chrome encerrou a extensão/host).
    /// </summary>
    public static JsonElement? ReadMessage()
    {
        var lengthBytes = ReadExactly(4);
        if (lengthBytes == null) return null;

        uint length = BitConverter.ToUInt32(lengthBytes, 0);
        if (length == 0) return null;

        var payload = ReadExactly((int)length);
        if (payload == null) return null;

        var json = Encoding.UTF8.GetString(payload);
        using var doc = JsonDocument.Parse(json);
        return doc.RootElement.Clone();
    }

    /// <summary>Envia um objeto como mensagem JSON para o Chrome.</summary>
    public static void WriteMessage(object message)
    {
        var json = JsonSerializer.Serialize(message);
        var payload = Encoding.UTF8.GetBytes(json);
        var lengthBytes = BitConverter.GetBytes((uint)payload.Length);

        lock (WriteLock)
        {
            StdOut.Write(lengthBytes, 0, 4);
            StdOut.Write(payload, 0, payload.Length);
            StdOut.Flush();
        }
    }

    private static byte[]? ReadExactly(int count)
    {
        var buffer = new byte[count];
        int read = 0;
        while (read < count)
        {
            int n = StdIn.Read(buffer, read, count - read);
            if (n == 0) return null; // EOF
            read += n;
        }
        return buffer;
    }
}
