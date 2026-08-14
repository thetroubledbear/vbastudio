using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace VbaStudio.Core.Dap;

public sealed record Capabilities(
    [property: JsonPropertyName("supportsConfigurationDoneRequest")] bool SupportsConfigurationDoneRequest);

public sealed record LaunchArguments(
    [property: JsonPropertyName("program")] string Program,
    [property: JsonPropertyName("entryPoint")] string EntryPoint);

public sealed record DapSource(
    [property: JsonPropertyName("path")] string? Path);

public sealed record DapSourceBreakpoint(
    [property: JsonPropertyName("line")] int Line);

public sealed record SetBreakpointsArguments(
    [property: JsonPropertyName("source")] DapSource Source,
    [property: JsonPropertyName("breakpoints")] IReadOnlyList<DapSourceBreakpoint> Breakpoints);

public sealed record DapBreakpoint(
    [property: JsonPropertyName("verified")] bool Verified,
    [property: JsonPropertyName("line")] int? Line);

public sealed record SetBreakpointsResponseBody(
    [property: JsonPropertyName("breakpoints")] IReadOnlyList<DapBreakpoint> Breakpoints);

public sealed record DapThread(
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("name")] string Name);

public sealed record ThreadsResponseBody(
    [property: JsonPropertyName("threads")] IReadOnlyList<DapThread> Threads);

public sealed record DapStackFrame(
    [property: JsonPropertyName("id")] int Id,
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("source")] DapSource Source,
    [property: JsonPropertyName("line")] int Line,
    [property: JsonPropertyName("column")] int Column);

public sealed record StackTraceResponseBody(
    [property: JsonPropertyName("stackFrames")] IReadOnlyList<DapStackFrame> StackFrames);

public sealed record DapScope(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("variablesReference")] int VariablesReference,
    [property: JsonPropertyName("expensive")] bool Expensive);

public sealed record ScopesResponseBody(
    [property: JsonPropertyName("scopes")] IReadOnlyList<DapScope> Scopes);

public sealed record DapVariable(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("value")] string Value,
    [property: JsonPropertyName("type")] string? Type,
    [property: JsonPropertyName("variablesReference")] int VariablesReference);

public sealed record VariablesResponseBody(
    [property: JsonPropertyName("variables")] IReadOnlyList<DapVariable> Variables);

public sealed record ContinueResponseBody(
    [property: JsonPropertyName("allThreadsContinued")] bool AllThreadsContinued);

public sealed record StoppedEventBody(
    [property: JsonPropertyName("reason")] string Reason,
    [property: JsonPropertyName("threadId")] int ThreadId,
    [property: JsonPropertyName("description")] string? Description);

public sealed record OutputEventBody(
    [property: JsonPropertyName("category")] string Category,
    [property: JsonPropertyName("output")] string Output);
