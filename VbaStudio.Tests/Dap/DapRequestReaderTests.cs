using System.Collections.Generic;
using System.IO;
using System.Threading;
using VbaStudio.Core.Dap;
using Xunit;

namespace VbaStudio.Tests.Dap;

public class DapRequestReaderTests
{
    private static MemoryStream StreamOf(params DapRequest[] requests)
    {
        var stream = new MemoryStream();
        foreach (var request in requests)
        {
            DapProtocol.WriteMessage(stream, request);
        }

        stream.Position = 0;
        return stream;
    }

    [Fact]
    public void Read_NotPaused_OrdinaryCommand_IsDequeuedNormally()
    {
        var stream = StreamOf(new DapRequest(1, "threads", null));
        var reader = new DapRequestReader(stream, () => false, _ => Assert.Fail("should not reach the running-stop handler"));

        reader.Start();
        var result = reader.Read();

        Assert.NotNull(result);
        Assert.Equal("threads", result!.Command);
    }

    [Fact]
    public void Read_EmptyStream_ReturnsNull()
    {
        var stream = StreamOf();
        var reader = new DapRequestReader(stream, () => false, _ => Assert.Fail("should not reach the running-stop handler"));

        reader.Start();
        var result = reader.Read();

        Assert.Null(result);
    }

    [Fact]
    public void Read_DisconnectWhilePaused_IsDequeuedNormally()
    {
        var stream = StreamOf(new DapRequest(1, "disconnect", null));
        var reader = new DapRequestReader(stream, () => true, _ => Assert.Fail("paused - must not use the running-stop handler"));

        reader.Start();
        var result = reader.Read();

        Assert.NotNull(result);
        Assert.Equal("disconnect", result!.Command);
    }

    [Theory]
    [InlineData("disconnect")]
    [InlineData("terminate")]
    public void Read_StopCommandWhileNotPaused_IsHandledInlineNotQueued(string command)
    {
        var stream = StreamOf(new DapRequest(1, command, null), new DapRequest(2, "threads", null));
        var handled = new List<DapRequest>();
        var reader = new DapRequestReader(stream, () => false, req => handled.Add(req));

        reader.Start();
        var result = reader.Read();

        // The stop command never reaches the queue - the next dequeue is the request behind it.
        Assert.NotNull(result);
        Assert.Equal("threads", result!.Command);
        var handledRequest = Assert.Single(handled);
        Assert.Equal(command, handledRequest.Command);
    }

    [Fact]
    public void Read_StopCommandWhileNotPaused_DoesNotBlockSubsequentReads()
    {
        // Regression guard for the bug this class exists to fix: a stop request must never sit
        // unread while the (simulated) consumer is "busy" - here modeled by isPaused always false.
        var stream = StreamOf(
            new DapRequest(1, "terminate", null),
            new DapRequest(2, "setBreakpoints", null));
        var handledCount = 0;
        var reader = new DapRequestReader(stream, () => false, _ => Interlocked.Increment(ref handledCount));

        reader.Start();
        var result = reader.Read();

        Assert.NotNull(result);
        Assert.Equal("setBreakpoints", result!.Command);
        Assert.Equal(1, handledCount);
    }
}
