// VbaStudio.Core/Debug/ProbeServer.cs
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Threading;
using VbaStudio.Core.Instrumentation;

namespace VbaStudio.Core.Debug;

public sealed record ProbeVariable(string Name, string Type, string Value);

public sealed record ProbeEvent(int ProbeId, string ModuleName, int OriginalLine, IReadOnlyList<ProbeVariable> Variables);

public enum ProbeCommand { Continue, Abort }

public sealed class ProbeServer : IDisposable
{
    private readonly HttpListener _listener;
    private readonly IReadOnlyDictionary<int, ProbeSite> _probeSites;
    private readonly Func<ProbeEvent, ProbeCommand> _onProbe;
    private readonly Action<string>? _log;
    private Thread? _thread;
    private volatile bool _running;

    public ProbeServer(
        int port,
        IReadOnlyDictionary<int, ProbeSite> probeSites,
        Func<ProbeEvent, ProbeCommand> onProbe,
        Action<string>? log = null)
    {
        _probeSites = probeSites;
        _onProbe = onProbe;
        _log = log;
        _listener = new HttpListener();
        _listener.Prefixes.Add($"http://localhost:{port}/probe/");
    }

    public void Start()
    {
        _listener.Start();
        _running = true;
        _thread = new Thread(Loop) { IsBackground = true };
        _thread.Start();
    }

    public void Stop()
    {
        _running = false;
        _listener.Stop();
        _thread?.Join(TimeSpan.FromSeconds(5));
        _thread = null;
    }

    public void Dispose()
    {
        Stop();
        _listener.Close();
    }

    private void Loop()
    {
        while (_running)
        {
            HttpListenerContext context;
            try
            {
                context = _listener.GetContext();
            }
            catch (HttpListenerException)
            {
                // Listener was Stop()'d while GetContext() was blocking - expected shutdown path.
                break;
            }
            catch (ObjectDisposedException)
            {
                break;
            }

            try
            {
                HandleRequest(context);
            }
            catch (Exception ex)
            {
                // The response-writing section of HandleRequest runs outside its own internal
                // try/catch (which only covers body-reading/parsing/dispatch). If the client has
                // already vanished (VBA-side timeout, Excel crash, or Stop() racing an in-flight
                // write against Dispose()'s listener Close()), OutputStream.Write/Close can throw.
                // An unhandled exception on a background thread would terminate the whole process,
                // which is worse than the single-probe hang this server exists to avoid - so catch,
                // log, and keep serving the next request instead of letting it propagate.
                _log?.Invoke($"ProbeServer: unhandled exception in HandleRequest, continuing loop: {ex.Message}");
            }
        }
    }

    private void HandleRequest(HttpListenerContext context)
    {
        ProbeCommand command;
        try
        {
            using var reader = new StreamReader(context.Request.InputStream, Encoding.UTF8);
            var body = reader.ReadToEnd();
            var probeEvent = ParseProbeEvent(body);
            command = _onProbe(probeEvent);
        }
        catch (Exception ex)
        {
            // A crashed delegate, an unknown probe_id, or malformed JSON must never leave VBA's
            // blocking HTTP call waiting forever - fail closed (abort) rather than fail hung.
            _log?.Invoke($"ProbeServer: request handling failed, defaulting to abort: {ex.Message}");
            command = ProbeCommand.Abort;
        }

        var cmdText = command == ProbeCommand.Abort ? "abort" : "continue";
        var responseBytes = Encoding.UTF8.GetBytes($"{{\"cmd\":\"{cmdText}\"}}");
        context.Response.ContentType = "application/json";
        context.Response.ContentLength64 = responseBytes.Length;
        context.Response.OutputStream.Write(responseBytes, 0, responseBytes.Length);
        context.Response.OutputStream.Close();
    }

    private ProbeEvent ParseProbeEvent(string body)
    {
        using var doc = JsonDocument.Parse(body);
        var root = doc.RootElement;
        var probeId = root.GetProperty("probe_id").GetInt32();

        var variables = new List<ProbeVariable>();
        if (root.TryGetProperty("vars", out var varsElement))
        {
            foreach (var v in varsElement.EnumerateArray())
            {
                var name = v.GetProperty("n").GetString() ?? "";
                var type = v.TryGetProperty("t", out var t) ? t.GetString() ?? "" : "";
                var value = v.GetProperty("v").GetString() ?? "";
                variables.Add(new ProbeVariable(name, type, value));
            }
        }

        if (!_probeSites.TryGetValue(probeId, out var site))
        {
            throw new InvalidOperationException($"Unknown probe_id {probeId} - not in the instrumented ProbeSite map.");
        }

        return new ProbeEvent(probeId, site.ModuleName, site.OriginalLine, variables);
    }
}
