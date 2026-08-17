// VbaStudio.Tests/Debug/ProbeServerTests.cs
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using VbaStudio.Core.Debug;
using VbaStudio.Core.Instrumentation;
using Xunit;

namespace VbaStudio.Tests.Debug;

public class ProbeServerTests
{
    private static int _nextPort = 18700;

    private static int GetTestPort() => Interlocked.Increment(ref _nextPort);

    [Fact]
    public async Task Probe_KnownProbeId_BuildsCorrectProbeEventAndReturnsContinue()
    {
        var port = GetTestPort();
        var probeSites = new Dictionary<int, ProbeSite>
        {
            [7] = new ProbeSite(7, "modWork", 42),
        };

        ProbeEvent? captured = null;
        using var server = new ProbeServer(port, probeSites, e =>
        {
            captured = e;
            return ProbeCommand.Continue;
        });
        server.Start();

        using var client = new HttpClient();
        var body = "{\"probe_id\":7,\"vars\":[{\"n\":\"i\",\"t\":\"Long\",\"v\":\"3\"}]}";
        var response = await client.PostAsync(
            $"http://localhost:{port}/probe/", new StringContent(body, Encoding.UTF8, "application/json"));
        var responseText = await response.Content.ReadAsStringAsync();

        server.Stop();

        Assert.NotNull(captured);
        Assert.Equal(7, captured!.ProbeId);
        Assert.Equal("modWork", captured.ModuleName);
        Assert.Equal(42, captured.OriginalLine);
        var variable = Assert.Single(captured.Variables);
        Assert.Equal("i", variable.Name);
        Assert.Equal("Long", variable.Type);
        Assert.Equal("3", variable.Value);
        Assert.Contains("\"cmd\":\"continue\"", responseText);
    }

    [Fact]
    public async Task Probe_OnProbeReturnsAbort_ReturnsAbortResponse()
    {
        var port = GetTestPort();
        var probeSites = new Dictionary<int, ProbeSite> { [1] = new ProbeSite(1, "modWork", 10) };

        using var server = new ProbeServer(port, probeSites, _ => ProbeCommand.Abort);
        server.Start();

        using var client = new HttpClient();
        var body = "{\"probe_id\":1,\"vars\":[]}";
        var response = await client.PostAsync(
            $"http://localhost:{port}/probe/", new StringContent(body, Encoding.UTF8, "application/json"));
        var responseText = await response.Content.ReadAsStringAsync();

        server.Stop();

        Assert.Contains("\"cmd\":\"abort\"", responseText);
    }

    [Fact]
    public async Task Probe_OnProbeThrows_DefaultsToAbortWithoutHanging()
    {
        var port = GetTestPort();
        var probeSites = new Dictionary<int, ProbeSite> { [1] = new ProbeSite(1, "modWork", 10) };
        var logs = new List<string>();

        using var server = new ProbeServer(
            port, probeSites, _ => throw new InvalidOperationException("boom"), logs.Add);
        server.Start();

        using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };
        var body = "{\"probe_id\":1,\"vars\":[]}";
        var response = await client.PostAsync(
            $"http://localhost:{port}/probe/", new StringContent(body, Encoding.UTF8, "application/json"));
        var responseText = await response.Content.ReadAsStringAsync();

        server.Stop();

        Assert.Contains("\"cmd\":\"abort\"", responseText);
        Assert.Contains(logs, l => l.Contains("boom"));
    }

    [Fact]
    public async Task Probe_UnknownProbeId_DefaultsToAbort()
    {
        var port = GetTestPort();
        var probeSites = new Dictionary<int, ProbeSite> { [1] = new ProbeSite(1, "modWork", 10) };

        using var server = new ProbeServer(port, probeSites, _ => ProbeCommand.Continue);
        server.Start();

        using var client = new HttpClient();
        var body = "{\"probe_id\":999,\"vars\":[]}";
        var response = await client.PostAsync(
            $"http://localhost:{port}/probe/", new StringContent(body, Encoding.UTF8, "application/json"));
        var responseText = await response.Content.ReadAsStringAsync();

        server.Stop();

        Assert.Contains("\"cmd\":\"abort\"", responseText);
    }

    [Fact]
    public async Task Probe_MalformedBody_DefaultsToAbortAndListenerStaysAlive()
    {
        var port = GetTestPort();
        var probeSites = new Dictionary<int, ProbeSite> { [1] = new ProbeSite(1, "modWork", 10) };
        var calls = 0;

        using var server = new ProbeServer(port, probeSites, _ =>
        {
            calls++;
            return ProbeCommand.Continue;
        });
        server.Start();

        using var client = new HttpClient();

        var malformedResponse = await client.PostAsync(
            $"http://localhost:{port}/probe/", new StringContent("not json", Encoding.UTF8, "application/json"));
        var malformedText = await malformedResponse.Content.ReadAsStringAsync();
        Assert.Contains("\"cmd\":\"abort\"", malformedText);
        Assert.Equal(0, calls);

        var validBody = "{\"probe_id\":1,\"vars\":[]}";
        var validResponse = await client.PostAsync(
            $"http://localhost:{port}/probe/", new StringContent(validBody, Encoding.UTF8, "application/json"));
        var validText = await validResponse.Content.ReadAsStringAsync();

        server.Stop();

        Assert.Contains("\"cmd\":\"continue\"", validText);
        Assert.Equal(1, calls);
    }

    [Fact]
    public void Loop_ResponseWriteFailureAfterStopCalledMidRequest_DoesNotCrashProcess()
    {
        var port = GetTestPort();
        var probeSites = new Dictionary<int, ProbeSite> { [1] = new ProbeSite(1, "modWork", 10) };

        ProbeServer? serverRef = null;
        var server = new ProbeServer(port, probeSites, _ =>
        {
            // Simulate the client vanishing exactly as the response is about to be written -
            // Stop() closes the listener out from under the in-flight response write, which
            // must not crash the process (would previously kill the whole test process, not
            // just fail this assertion).
            serverRef?.Stop();
            return ProbeCommand.Continue;
        });
        serverRef = server;
        server.Start();

        using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
        var body = "{\"probe_id\":1,\"vars\":[]}";
        try
        {
            client.PostAsync($"http://localhost:{port}/probe/", new StringContent(body, Encoding.UTF8, "application/json"))
                .GetAwaiter().GetResult();
        }
        catch
        {
            // The client-side request may itself fail once the listener is stopped mid-response -
            // that's an expected, acceptable outcome here. The actual assertion is that reaching
            // this point at all (the test process is still alive) proves the fix worked.
        }

        server.Dispose();
        Assert.True(true); // Reaching this line proves the background thread's exception was caught, not left to crash the process.
    }
}
