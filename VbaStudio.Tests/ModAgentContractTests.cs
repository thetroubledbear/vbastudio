using System.IO;
using Xunit;

namespace VbaStudio.Tests;

public class ModAgentContractTests
{
    private static string ReadModAgentSource()
    {
        var path = Path.Combine("vba", "modAgent.bas");
        return File.ReadAllText(path);
    }

    [Fact]
    public void ModAgent_ModuleIsNamedAgent()
    {
        // Instrumenter.cs hardcodes calls to "Agent.Probe" - if this module's VB_Name isn't
        // literally "Agent", VBA cannot resolve those calls at all (confirmed live: this exact
        // mismatch produced "Compile error: Variable not defined" against real Excel).
        var source = ReadModAgentSource();
        Assert.Contains("Attribute VB_Name = \"Agent\"", source);
    }

    [Fact]
    public void ModAgent_ProbeDoesNotUseParamArray()
    {
        // Every call site (Instrumenter's generated "Agent.Probe <id>, Array(...)") passes the
        // Array(...) result as a SINGLE argument. A ParamArray parameter would nest that single
        // argument as one element instead of spreading it, producing a runtime Type mismatch -
        // confirmed live against real Excel. Probe must take a plain Variant, not ParamArray.
        var source = ReadModAgentSource();
        Assert.DoesNotContain("ParamArray", source);
    }

    [Fact]
    public void ModAgent_PortMatchesDebugSessionAgentPort()
    {
        // DebugSession.AgentPort and modAgent.bas's AGENT_PORT constant must agree - there is no
        // config-injection mechanism into VBA source, so both sides hardcode the same literal.
        var source = ReadModAgentSource();
        Assert.Contains("AGENT_PORT As Long = 8731", source);
    }
}
