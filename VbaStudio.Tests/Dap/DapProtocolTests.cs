using System.IO;
using System.Text;
using System.Text.Json;
using VbaStudio.Core.Dap;
using Xunit;

namespace VbaStudio.Tests.Dap;

public class DapProtocolTests
{
    [Fact]
    public void WriteMessage_ThenReadRequest_RoundTripsBasicRequest()
    {
        var stream = new MemoryStream();
        var original = new DapRequest(1, "initialize", null);

        DapProtocol.WriteMessage(stream, original);
        stream.Position = 0;

        var result = DapProtocol.ReadRequest(stream);

        Assert.NotNull(result);
        Assert.Equal(1, result!.Seq);
        Assert.Equal("initialize", result.Command);
        Assert.Null(result.Arguments);
    }

    [Fact]
    public void WriteMessage_IncludesTypeField()
    {
        var stream = new MemoryStream();
        var request = new DapRequest(1, "initialize", null);

        DapProtocol.WriteMessage(stream, request);
        stream.Position = 0;
        var text = new StreamReader(stream, Encoding.UTF8).ReadToEnd();

        Assert.Contains("\"type\":\"request\"", text.Replace(" ", ""));
    }

    [Fact]
    public void ReadRequest_RequestWithNestedArguments_SurvivesRoundTrip()
    {
        var stream = new MemoryStream();
        var json = "{\"seq\":5,\"type\":\"request\",\"command\":\"setBreakpoints\"," +
                    "\"arguments\":{\"source\":{\"path\":\"src/Modules/modWork.bas\"}," +
                    "\"breakpoints\":[{\"line\":10},{\"line\":12}]}}";
        WriteRaw(stream, json);
        stream.Position = 0;

        var result = DapProtocol.ReadRequest(stream);

        Assert.NotNull(result);
        Assert.Equal(5, result!.Seq);
        Assert.Equal("setBreakpoints", result.Command);
        Assert.NotNull(result.Arguments);

        var args = result.Arguments!.Value.Deserialize<SetBreakpointsArguments>()!;
        Assert.Equal("src/Modules/modWork.bas", args.Source.Path);
        Assert.Equal(2, args.Breakpoints.Count);
        Assert.Equal(10, args.Breakpoints[0].Line);
        Assert.Equal(12, args.Breakpoints[1].Line);
    }

    [Fact]
    public void ReadRequest_MalformedContentLengthHeader_ReturnsNull()
    {
        var stream = new MemoryStream();
        var bytes = Encoding.ASCII.GetBytes("NotContentLength: 5\r\n\r\nhello");
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 0;

        var result = DapProtocol.ReadRequest(stream);

        Assert.Null(result);
    }

    [Fact]
    public void ReadRequest_EmptyStream_ReturnsNull()
    {
        var stream = new MemoryStream();

        var result = DapProtocol.ReadRequest(stream);

        Assert.Null(result);
    }

    [Fact]
    public void ReadRequest_TruncatedBody_ReturnsNull()
    {
        var stream = new MemoryStream();
        var bytes = Encoding.ASCII.GetBytes("Content-Length: 100\r\n\r\n{\"seq\":1}");
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 0;

        var result = DapProtocol.ReadRequest(stream);

        Assert.Null(result);
    }

    private static void WriteRaw(MemoryStream stream, string json)
    {
        var jsonBytes = Encoding.UTF8.GetBytes(json);
        var header = $"Content-Length: {jsonBytes.Length}\r\n\r\n";
        var headerBytes = Encoding.ASCII.GetBytes(header);
        stream.Write(headerBytes, 0, headerBytes.Length);
        stream.Write(jsonBytes, 0, jsonBytes.Length);
    }
}
