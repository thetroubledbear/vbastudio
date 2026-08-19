// VbaStudio.DapServer/DapSession.cs
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using VbaStudio.Core.Dap;
using VbaStudio.Core.Debug;
using VbaStudio.Core.Instrumentation;
using VbaStudio.Core.Model;
using VbaStudio.Core.Parsing;
using VbaStudio.Core.Win32;
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
    private readonly IWin32Windows _windows = new Win32Windows();
    private int _nextSeq = 1;

    private string? _launchModule;
    private string? _launchEntryPoint;
    private bool _running;
    private IReadOnlyDictionary<string, string> _moduleFilePaths = new Dictionary<string, string>();
    private IReadOnlyList<ProbeSite> _launchProbeSites = Array.Empty<ProbeSite>();
    private DapRequestReader? _reader;

    private readonly object _breakpointsLock = new();
    private readonly HashSet<(string Module, int Line)> _breakpoints = new();

    // Read by DapRequestReader's own background thread (to decide whether a "disconnect"/
    // "terminate" needs the immediate Ctrl+Break path) concurrently with writes from OnProbe on
    // ProbeServer's callback thread - volatile keeps that check honest across threads.
    private volatile ProbeEvent? _currentProbe;

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
    // the STA thread's own RCWs while that thread sits non-pumping). DapRequestReader's own
    // background thread (below) follows the same rule: it only ever does raw stream IO on
    // `input`, never a COM call, so it is safe to run concurrently with Application.Run too -
    // every actual request handler (HandleRequest, OnProbe's body) still only ever runs on
    // whichever thread dequeues it via _reader.Read(): this STA thread before the run starts, or
    // ProbeServer's callback thread while a probe is paused mid-run (OnProbe's own loop below).
    // This exactly mirrors VbaStudio.Spike's own already-proven-live `debug` console mode, which
    // reads Console.ReadLine() synchronously inside its onProbe callback on ProbeServer's
    // callback thread - never a thread that also touches _excel/_workbook's COM objects.
    public void RunMessageLoop(Stream input)
    {
        // One reader thread for the whole session, started here and never restarted - both this
        // loop and OnProbe's own paused loop below consume through it via _reader.Read() instead
        // of touching `input` directly, so a stop request can be caught and acted on immediately
        // even while the main thread is blocked inside Application.Run. See DapRequestReader.
        _reader = new DapRequestReader(input, () => _currentProbe != null, HandleStopWhileRunning);
        _reader.Start();

        while (true)
        {
            var request = _reader.Read();
            if (request == null)
            {
                break;
            }

            if (request.Command == "configurationDone")
            {
                HandleConfigurationDone(request);
                break;
            }

            try
            {
                HandleRequest(request);
            }
            catch (Exception ex)
            {
                SendErrorResponse(request, ex.Message);
            }
        }
    }

    // Invoked on DapRequestReader's own background thread - concurrently with the main thread
    // sitting blocked inside Application.Run - whenever "disconnect"/"terminate" arrives while no
    // breakpoint is currently paused. Nobody else is going to read further input soon, so this
    // must act immediately rather than queue. Sending the interrupt is pure Win32 key injection,
    // never a call into any STA-owned COM RCW, so it is safe from a non-owning thread - see the
    // threading note on RunMessageLoop's own caller and DapRequestReader's doc comment.
    private void HandleStopWhileRunning(DapRequest request)
    {
        _log?.Invoke(
            $"DapSession: '{request.Command}' received while a procedure is running with no " +
            "breakpoint paused - sending Ctrl+Break to interrupt Excel.");
        SendResponse(request, null);

        try
        {
            _windows.SendCtrlBreak((IntPtr)_excel.Hwnd);
        }
        catch (Exception ex)
        {
            _log?.Invoke($"DapSession: failed to send the interrupt signal: {ex.Message}");
        }

        // Not blocking here for Application.Run to actually return - RunDebugSession's own flow
        // sends "terminated" once DebugSession.Run does, whether that's because the interrupt
        // above worked (DialogWatcher, already running, dismisses the resulting "code execution
        // interrupted" dialog by clicking End) or the procedure simply finished on its own in the
        // meantime.
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
                // Reached here only if it arrives before configurationDone - nothing is running
                // yet. While paused at a breakpoint, OnProbe's own loop intercepts disconnect/
                // terminate before calling HandleRequest (see OnProbe below). Once running with
                // no breakpoint paused, DapRequestReader routes it to HandleStopWhileRunning
                // instead of ever reaching here, since nothing on this thread will call Read()
                // again until the blocking Application.Run call returns.
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
            case "continue":
                SendErrorResponse(request, "Not currently paused at a breakpoint.");
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
        // The "initialized" event is deliberately NOT sent here. It tells the client "send your
        // configuration requests now", and setBreakpoints - the first thing the client sends in
        // response - can only report genuine verified status once _launchProbeSites is populated,
        // which only happens in HandleLaunch. Sending it here races launch: a client that wins
        // that race gets verified:false for every breakpoint. Emitting it at the end of
        // HandleLaunch instead makes the ordering guaranteed rather than incidental (the standard
        // pattern for adapters that need launch arguments before answering configuration
        // requests).
    }

    private void HandleLaunch(DapRequest request)
    {
        if (_running)
        {
            SendErrorResponse(request, "launch is not valid after the debug session has started.");
            return;
        }

        var args = request.Arguments!.Value.Deserialize<LaunchArguments>()!;

        if (!string.Equals(_workbook.FullName, args.Program, StringComparison.OrdinalIgnoreCase))
        {
            SendErrorResponse(request, $"launch targets '{args.Program}', but the attached Excel instance has '{_workbook.FullName}' active. Open the target workbook and make it the active window before starting the debug session.");
            return;
        }

        var dotIndex = args.EntryPoint.LastIndexOf('.');
        _launchModule = args.EntryPoint.Substring(0, dotIndex);
        _launchEntryPoint = args.EntryPoint;

        var workbookDir = Path.GetDirectoryName(_workbook.FullName) ?? ".";
        var access = new ExcelVbaProjectAccess(_workbook.VBProject);
        var modules = access.ReadAll();
        _moduleFilePaths = modules.ToDictionary(
            m => m.Name,
            m => Path.Combine(workbookDir, "src", m.Kind.SourceFolder(), m.Name + m.Kind.FileExtension()));

        _launchProbeSites = ComputeLaunchProbeSites(modules, _launchModule, args.EntryPoint);

        SendResponse(request, null);

        // See HandleInitialize: emitted here, not there, so that _launchProbeSites is already
        // populated by the time the client starts sending setBreakpoints.
        SendEvent("initialized", null);
    }

    private void HandleSetBreakpoints(DapRequest request)
    {
        var args = request.Arguments!.Value.Deserialize<SetBreakpointsArguments>()!;
        var moduleName = Path.GetFileNameWithoutExtension(args.Source.Path ?? "").ToUpperInvariant();

        var verifiedBreakpoints = BreakpointVerifier.ComputeVerifiedBreakpoints(
            _launchProbeSites, moduleName, args.Breakpoints.Select(bp => bp.Line).ToList());

        // Every requested line goes into the set, verified or not - "verified" affects only the
        // status reported back to the client. An unverified line has no probe placed at it, so it
        // can never fire regardless of set membership; filtering the set buys nothing and costs a
        // silent-drop failure mode if verification was ever computed against incomplete state.
        lock (_breakpointsLock)
        {
            _breakpoints.RemoveWhere(bp => bp.Module == moduleName);
            foreach (var bp in verifiedBreakpoints.Where(bp => bp.Line.HasValue))
            {
                _breakpoints.Add((moduleName, bp.Line!.Value));
            }
        }

        SendResponse(request, new SetBreakpointsResponseBody(verifiedBreakpoints));
    }

    // Instrumentation is a pure text transform (no COM, no shadow workbook) - running it here,
    // against the live project's current source, lets setBreakpoints report genuine
    // verified/unverified status instead of the unconditional verified:true M6a shipped with.
    // DebugSession.Run (triggered later, by configurationDone) re-instruments independently
    // against the shadow copy for the actual run; the two calls are deliberately not shared
    // state - both are cheap and deterministic over the same source, so there's no staleness
    // risk worth the coupling. Failure here (bad module/procedure name) is swallowed: the
    // launch handshake must not break over a verification preview, and the real error surfaces
    // identically at configurationDone exactly as it did before this change.
    private static IReadOnlyList<ProbeSite> ComputeLaunchProbeSites(
        IReadOnlyList<VbaModule> modules, string launchModule, string entryPointQualifiedName)
    {
        try
        {
            var targetModule = modules.FirstOrDefault(
                m => string.Equals(m.Name, launchModule, StringComparison.OrdinalIgnoreCase));
            if (targetModule == null)
            {
                return Array.Empty<ProbeSite>();
            }

            var moduleSymbols = VbaParser.ParseModule(targetModule.Code, targetModule.Name);
            var procedureName = entryPointQualifiedName.Substring(entryPointQualifiedName.LastIndexOf('.') + 1);
            var procedure = moduleSymbols.Procedures.FirstOrDefault(
                p => string.Equals(p.Name, procedureName, StringComparison.OrdinalIgnoreCase));
            if (procedure == null)
            {
                return Array.Empty<ProbeSite>();
            }

            return Instrumenter.Instrument(targetModule.Code, procedure, targetModule.Name).ProbeSites;
        }
        catch
        {
            return Array.Empty<ProbeSite>();
        }
    }

    private void HandleThreads(DapRequest request)
    {
        SendResponse(request, new ThreadsResponseBody(new[] { new DapThread(ThreadId, "Main") }));
    }

    private void HandleConfigurationDone(DapRequest request)
    {
        _running = true;
        SendResponse(request, null);
        RunDebugSession();
    }

    private void RunDebugSession()
    {
        // An unhandled exception here would crash the entire process - confirmed the hard way in
        // M5b's own Task 1 fix (ProbeServer's response-write path). Never let a session failure
        // escape uncaught here.
        try
        {
            var debugSession = new DebugSession();
            var result = debugSession.Run(
                _excel, _workbook, _shadowPath, _launchModule!, _launchEntryPoint!,
                OnProbe,
                log => SendEvent("output", new OutputEventBody("console", log + "\n")));

            if (!result.Run.Success)
            {
                var d = result.Run.Diagnostic;
                var message = d == null
                    ? "Run failed with no diagnostic detail."
                    : $"Run failed: module={d.Module ?? "?"} line={d.Line?.ToString() ?? "?"} message={d.Message}";
                SendEvent("output", new OutputEventBody("stderr", message + "\n"));
            }

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
    // Reads through _reader.Read() (not the raw input stream) - it is the single dedicated
    // reader for the whole session; see DapRequestReader and RunMessageLoop.
    private ProbeCommand OnProbe(ProbeEvent probeEvent)
    {
        bool isBreakpoint;
        lock (_breakpointsLock)
        {
            isBreakpoint = _breakpoints.Contains((probeEvent.ModuleName.ToUpperInvariant(), probeEvent.OriginalLine));
        }

        if (!isBreakpoint)
        {
            return ProbeCommand.Continue;
        }

        _currentProbe = probeEvent;
        SendEvent("stopped", new StoppedEventBody("breakpoint", ThreadId, null));

        while (true)
        {
            var request = _reader!.Read();
            if (request == null)
            {
                _currentProbe = null;
                return ProbeCommand.Abort;
            }

            if (request.Command == "continue")
            {
                SendResponse(request, new ContinueResponseBody(true));
                _currentProbe = null;
                return ProbeCommand.Continue;
            }

            if (request.Command == "disconnect" || request.Command == "terminate")
            {
                SendResponse(request, null);
                _currentProbe = null;
                return ProbeCommand.Abort;
            }

            try
            {
                HandleRequest(request);
            }
            catch (Exception ex)
            {
                SendErrorResponse(request, ex.Message);
            }
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
