using System.IO;
using System.Text;
using System.Text.Json;

namespace VbaStudio.Core.Dap;

public static class DapProtocol
{
    private const int MaxContentLength = 64 * 1024 * 1024; // 64 MB - generous, but bounds allocation

    // Reads one Content-Length-framed DAP request from input. Best-effort: any malformed or
    // truncated input returns null (treated as a clean disconnect by the caller) rather than
    // throwing - a DAP client closing the pipe mid-message must not crash the server.
    public static DapRequest? ReadRequest(Stream input)
    {
        var headerLine = ReadHeaderLine(input);
        if (string.IsNullOrEmpty(headerLine))
        {
            return null;
        }

        const string prefix = "Content-Length: ";
        if (!headerLine.StartsWith(prefix) || !int.TryParse(headerLine.Substring(prefix.Length), out var contentLength))
        {
            return null;
        }

        if (contentLength < 0 || contentLength > MaxContentLength)
        {
            return null;
        }

        // Consume any further header lines (e.g. Content-Type, which real DAP clients may send)
        // until the blank separator line - DAP permits headers beyond Content-Length.
        while (true)
        {
            var line = ReadHeaderLine(input);
            if (line == null)
            {
                return null;
            }

            if (line == "")
            {
                break;
            }
        }

        var buffer = new byte[contentLength];
        var totalRead = 0;
        while (totalRead < contentLength)
        {
            var read = input.Read(buffer, totalRead, contentLength - totalRead);
            if (read == 0)
            {
                return null;
            }

            totalRead += read;
        }

        try
        {
            var json = Encoding.UTF8.GetString(buffer);
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            if (!root.TryGetProperty("seq", out var seqEl) || seqEl.ValueKind != JsonValueKind.Number || !seqEl.TryGetInt32(out var seq))
            {
                return null;
            }

            if (!root.TryGetProperty("command", out var commandEl) || commandEl.ValueKind != JsonValueKind.String)
            {
                return null;
            }

            var command = commandEl.GetString() ?? "";
            JsonElement? arguments = null;
            if (root.TryGetProperty("arguments", out var argsEl) && argsEl.ValueKind != JsonValueKind.Null)
            {
                arguments = argsEl.Clone();
            }

            return new DapRequest(seq, command, arguments);
        }
        catch (JsonException)
        {
            return null;
        }
    }

    // Serializes message (a DapResponse or DapEvent) using its own runtime type - System.Text.Json
    // serializes object-typed values by their actual runtime type, not the declared "object", so
    // this correctly emits every property the caller's specific record type declares. Flushes
    // immediately: a real stdout pipe must not buffer a DAP message the client is waiting on.
    public static void WriteMessage(Stream output, object message)
    {
        var json = JsonSerializer.Serialize(message, message.GetType());
        var jsonBytes = Encoding.UTF8.GetBytes(json);
        var header = $"Content-Length: {jsonBytes.Length}\r\n\r\n";
        var headerBytes = Encoding.ASCII.GetBytes(header);
        output.Write(headerBytes, 0, headerBytes.Length);
        output.Write(jsonBytes, 0, jsonBytes.Length);
        output.Flush();
    }

    private static string? ReadHeaderLine(Stream input)
    {
        var sb = new StringBuilder();
        var lastWasCr = false;
        int b;
        while ((b = input.ReadByte()) != -1)
        {
            if (b == '\r')
            {
                lastWasCr = true;
                continue;
            }

            if (b == '\n' && lastWasCr)
            {
                return sb.ToString();
            }

            lastWasCr = false;
            sb.Append((char)b);
        }

        return sb.Length > 0 ? sb.ToString() : null;
    }
}
