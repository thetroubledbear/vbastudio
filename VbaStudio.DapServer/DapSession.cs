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

    private ProbeEvent? _currentProbe;

    public DapSession(Excel.Application excel, Excel.Workbook workbook, string shadowPath, Stream output)
    {
        _excel = excel;
        _workbook = workbook;
        _shadowPath = shadowPath;
        _output = output;
    }

    // Runs the whole session on the CALLING thread - the same [STAThread] thread that created
    // _excel/_workbook. Application.Run (called deep inside DebugSession.Run, below) must be
    // invoked on that same STA thread: confirmed live that the earlier design - giving
    // DebugSession.Run its own Thread.Start(), with no ApartmentState.STA set - made Excel's
    // Application.Run fail intermittently (a spurious VBA "Out of memory" error, or a
    // NullReferenceException deep in COM marshaling) on an unpredictable subset of runs. The
    // mechanism: an MTA thread calling into an STA-owned RCW needs the owning STA thread's
    // message pump to service the cross-apartment call, and this process's STA thread was
    // blocked in a plain, non-pumping input.ReadByte() at the time - so the call had nothing to
    // service it. Calling Application.Run directly on the STA thread removes the cross-apartment
    // hop entirely (this is NOT a "never touch a second thread" rule - ProbeServer's own
    // pre-existing callback thread, and DialogWatcher's polling thread, both already run
    // concurrently with Application.Run without issue, since neither makes a blocking call INTO
    // the STA thread's own RCWs while that thread sits non-pumping).
    // So the whole session - reading DAP requests before the run starts, running
    // DebugSession.Run, and reading further DAP requests (stackTrace/scopes/variables/continue)
    // while a probe is paused mid-run via OnProbe's own nested read loop below - all happens on
    // this one thread. This exactly mirrors VbaStudio.Spike's own already-proven-live `debug`
    // console mode, which reads Console.ReadLine() synchronously inside its onProbe callback on
    // ProbeServer's callback thread - never a second, independently-spawned thread.
    public void RunMessageLoop(Stream input)
    {
        while (true)
        {
            var request = DapProtocol.ReadRequest(input);
            if (request == null)
            {
                break;
            }

            if (request.Command == "configurationDone")
            {
                HandleConfigurationDone(request, input);
                break;
            }

            HandleRequest(request);
        }
    }

    private void HandleRequest(DapRequest request)
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
                // Normally intercepted by RunMessageLoop before reaching here (it needs the
                // input stream to hand off to OnProbe's own read loop). Reached only if a client
                // sends configurationDone a second time - respond, but don't re-run the session.
                SendResponse(request, null);
                break;
            case "disconnect":
            case "terminate":
                // Only meaningful here if it arrives before configurationDone (nothing is running
                // yet). Once running, disconnect/terminate is intercepted inline by OnProbe's own
                // read loop, but only while a probe is actually paused - see OnProbe below. If the
                // client disconnects while the procedure is running but no breakpoint is currently
                // hit, nobody is reading stdin at that moment (this thread is blocked inside
                // Application.Run) - the message sits unread until the next pause or until the run
                // finishes on its own. Accepted for M6a's scope (exit gate is "breakpoint,
                // variables pane", not mid-flight cancellation); a real client would time out and
                // kill the process rather than hang indefinitely.
                SendResponse(request, null);
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

    private void HandleConfigurationDone(DapRequest request, Stream input)
    {
        SendResponse(request, null);
        RunDebugSession(input);
    }

    private void RunDebugSession(Stream input)
    {
        // An unhandled exception here would crash the entire process - confirmed the hard way in
        // M5b's own Task 1 fix (ProbeServer's response-write path). Never let a session failure
        // escape uncaught here.
        try
        {
            var debugSession = new DebugSession();
            debugSession.Run(
                _excel, _workbook, _shadowPath, _launchModule!, _launchEntryPoint!,
                probeEvent => OnProbe(probeEvent, input),
                log => SendEvent("output", new OutputEventBody("console", log + "\n")));

            SendEvent("terminated", null);
        }
        catch (Exception ex)
        {
            SendEvent("output", new OutputEventBody("stderr", $"Debug session failed: {ex.Message}\n"));
            SendEvent("terminated", null);
        }
    }

    // Invoked on ProbeServer's own callback thread (the same thread class Spike's console
    // `debug` mode already drives its onProbe callback from - see the note on RunMessageLoop).
    // When paused at a breakpoint, this reads and answers further DAP requests directly and
    // synchronously - stackTrace/scopes/variables/setBreakpoints - until a continue or
    // disconnect/terminate arrives, at which point it returns control to DebugSession.Run.
    private ProbeCommand OnProbe(ProbeEvent probeEvent, Stream input)
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
        SendEvent("stopped", new StoppedEventBody("breakpoint", ThreadId, null));

        while (true)
        {
            var request = DapProtocol.ReadRequest(input);
            if (request == null)
            {
                return ProbeCommand.Abort;
            }

            if (request.Command == "continue")
            {
                SendResponse(request, new ContinueResponseBody(true));
                return ProbeCommand.Continue;
            }

            if (request.Command == "disconnect" || request.Command == "terminate")
            {
                SendResponse(request, null);
                return ProbeCommand.Abort;
            }

            HandleRequest(request);
        }
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
