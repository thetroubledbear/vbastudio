// VbaStudio.DapServer/DapSession.cs
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using VbaStudio.Core.Dap;
using VbaStudio.Core.Debug;
using VbaStudio.Core.Model;
using VbaStudio.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace VbaStudio.DapServer;

public sealed class DapSession
{
    private const int ThreadId = 1;
    private const int FrameId = 1;
    private const int ScopeVariablesReference = 1;

    private readonly Excel.Application _excel;
    private readonly Excel.Workbook _workbook;
    private readonly string _shadowPath;
    private readonly Stream _output;
    private readonly object _outputLock = new();
    private int _nextSeq = 1;

    private string? _launchModule;
    private string? _launchEntryPoint;
    private IReadOnlyDictionary<string, string> _moduleFilePaths = new Dictionary<string, string>();

    private readonly object _breakpointsLock = new();
    private readonly HashSet<(string Module, int Line)> _breakpoints = new();

    private readonly ManualResetEventSlim _continueSignal = new(false);
    private ProbeCommand _pendingCommand = ProbeCommand.Continue;
    private volatile ProbeEvent? _currentProbe;

    public DapSession(Excel.Application excel, Excel.Workbook workbook, string shadowPath, Stream output)
    {
        _excel = excel;
        _workbook = workbook;
        _shadowPath = shadowPath;
        _output = output;
    }

    public void HandleRequest(DapRequest request)
    {
        switch (request.Command)
        {
            case "initialize":
                HandleInitialize(request);
                break;
            case "launch":
                HandleLaunch(request);
                break;
            case "setBreakpoints":
                HandleSetBreakpoints(request);
                break;
            case "threads":
                HandleThreads(request);
                break;
            case "configurationDone":
                HandleConfigurationDone(request);
                break;
            case "disconnect":
            case "terminate":
                HandleDisconnect(request);
                break;
            case "stackTrace":
                HandleStackTrace(request);
                break;
            case "scopes":
                HandleScopes(request);
                break;
            case "variables":
                HandleVariables(request);
                break;
            case "continue":
                HandleContinue(request);
                break;
            default:
                SendErrorResponse(request, $"Unknown or not-yet-implemented command: {request.Command}");
                break;
        }
    }

    private void HandleInitialize(DapRequest request)
    {
        var capabilities = new Capabilities(SupportsConfigurationDoneRequest: true);
        SendResponse(request, capabilities);
        SendEvent("initialized", null);
    }

    private void HandleLaunch(DapRequest request)
    {
        var args = request.Arguments!.Value.Deserialize<LaunchArguments>()!;
        var dotIndex = args.EntryPoint.LastIndexOf('.');
        _launchModule = args.EntryPoint.Substring(0, dotIndex);
        _launchEntryPoint = args.EntryPoint;

        var access = new ExcelVbaProjectAccess(_workbook.VBProject);
        _moduleFilePaths = access.ReadAll().ToDictionary(
            m => m.Name,
            m => Path.Combine("src", m.Kind.SourceFolder(), m.Name + m.Kind.FileExtension()));

        SendResponse(request, null);
    }

    private void HandleSetBreakpoints(DapRequest request)
    {
        var args = request.Arguments!.Value.Deserialize<SetBreakpointsArguments>()!;
        var moduleName = Path.GetFileNameWithoutExtension(args.Source.Path ?? "");

        var verifiedBreakpoints = new List<DapBreakpoint>();
        lock (_breakpointsLock)
        {
            _breakpoints.RemoveWhere(bp => bp.Module == moduleName);
            foreach (var bp in args.Breakpoints)
            {
                _breakpoints.Add((moduleName, bp.Line));
                verifiedBreakpoints.Add(new DapBreakpoint(Verified: true, Line: bp.Line));
            }
        }

        SendResponse(request, new SetBreakpointsResponseBody(verifiedBreakpoints));
    }

    private void HandleThreads(DapRequest request)
    {
        SendResponse(request, new ThreadsResponseBody(new[] { new DapThread(ThreadId, "Main") }));
    }

    private void HandleConfigurationDone(DapRequest request)
    {
        SendResponse(request, null);

        var thread = new Thread(RunDebugSession) { IsBackground = true };
        thread.Start();
    }

    private void HandleDisconnect(DapRequest request)
    {
        SendResponse(request, null);
        _pendingCommand = ProbeCommand.Abort;
        _continueSignal.Set();
    }

    private void RunDebugSession()
    {
        // An unhandled exception on this background thread would crash the entire process -
        // confirmed the hard way in M5b's own Task 1 fix (ProbeServer's response-write path).
        // Never let a session failure escape uncaught here.
        try
        {
            var debugSession = new DebugSession();
            debugSession.Run(
                _excel, _workbook, _shadowPath, _launchModule!, _launchEntryPoint!,
                OnProbe,
                log => SendEvent("output", new OutputEventBody("console", log + "\n")));

            SendEvent("terminated", null);
        }
        catch (Exception ex)
        {
            SendEvent("output", new OutputEventBody("stderr", $"Debug session failed: {ex.Message}\n"));
            SendEvent("terminated", null);
        }
    }

    private ProbeCommand OnProbe(ProbeEvent probeEvent)
    {
        bool isBreakpoint;
        lock (_breakpointsLock)
        {
            isBreakpoint = _breakpoints.Contains((probeEvent.ModuleName, probeEvent.OriginalLine));
        }

        if (!isBreakpoint)
        {
            return ProbeCommand.Continue;
        }

        _currentProbe = probeEvent;
        _continueSignal.Reset();

        SendEvent("stopped", new StoppedEventBody("breakpoint", ThreadId, null));

        _continueSignal.Wait();
        return _pendingCommand;
    }

    private void HandleStackTrace(DapRequest request)
    {
        var probe = _currentProbe;
        if (probe == null)
        {
            SendResponse(request, new StackTraceResponseBody(Array.Empty<DapStackFrame>()));
            return;
        }

        var path = _moduleFilePaths.TryGetValue(probe.ModuleName, out var p) ? p : probe.ModuleName;
        var frame = new DapStackFrame(FrameId, probe.ModuleName, new DapSource(path), probe.OriginalLine, 1);
        SendResponse(request, new StackTraceResponseBody(new[] { frame }));
    }

    private void HandleScopes(DapRequest request)
    {
        var scope = new DapScope("Locals", ScopeVariablesReference, Expensive: false);
        SendResponse(request, new ScopesResponseBody(new[] { scope }));
    }

    private void HandleVariables(DapRequest request)
    {
        var probe = _currentProbe;
        var variables = probe?.Variables
            .Select(v => new DapVariable(v.Name, v.Value, v.Type, 0))
            .ToArray() ?? Array.Empty<DapVariable>();
        SendResponse(request, new VariablesResponseBody(variables));
    }

    private void HandleContinue(DapRequest request)
    {
        SendResponse(request, new ContinueResponseBody(true));
        _pendingCommand = ProbeCommand.Continue;
        _continueSignal.Set();
    }

    private void SendResponse(DapRequest request, object? body)
    {
        var response = new DapResponse(NextSeq(), request.Seq, true, request.Command, body);
        WriteMessage(response);
    }

    private void SendErrorResponse(DapRequest request, string message)
    {
        var response = new DapResponse(NextSeq(), request.Seq, false, request.Command, null, message);
        WriteMessage(response);
    }

    private void SendEvent(string eventName, object? body)
    {
        var evt = new DapEvent(NextSeq(), eventName, body);
        WriteMessage(evt);
    }

    private void WriteMessage(object message)
    {
        lock (_outputLock)
        {
            DapProtocol.WriteMessage(_output, message);
        }
    }

    private int NextSeq() => Interlocked.Increment(ref _nextSeq);
}
