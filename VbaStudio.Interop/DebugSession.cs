// VbaStudio.Interop/DebugSession.cs
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using VbaStudio.Core.Debug;
using VbaStudio.Core.Instrumentation;
using VbaStudio.Core.Model;
using VbaStudio.Core.Parsing;
using Excel = Microsoft.Office.Interop.Excel;

namespace VbaStudio.Interop;

public sealed record DebugResult(RunResult Run, IReadOnlyList<ProbeEvent> ProbesFired);

public sealed class DebugSession
{
    // Must match modAgent.bas's own hardcoded POST target port exactly - there is no
    // config-injection mechanism into VBA source, so both sides hardcode the same literal.
    private const int AgentPort = 8731;

    public DebugResult Run(
        Excel.Application excel,
        Excel.Workbook workbook,
        string shadowPath,
        string moduleName,
        string entryPointQualifiedName,
        Func<ProbeEvent, ProbeCommand> onProbe,
        Action<string>? log = null)
    {
        var shadowWorkbook = ExcelShadowWorkbook.CreateFromOpen(workbook, shadowPath);
        try
        {
            var access = new ExcelVbaProjectAccess(shadowWorkbook.VBProject);

            var modules = access.ReadAll();
            var targetModule = modules.FirstOrDefault(
                m => string.Equals(m.Name, moduleName, StringComparison.OrdinalIgnoreCase))
                ?? throw new InvalidOperationException($"Module '{moduleName}' not found in the shadow project.");

            var moduleSymbols = VbaParser.ParseModule(targetModule.Code, targetModule.Name);
            var procedureName = entryPointQualifiedName.Substring(entryPointQualifiedName.LastIndexOf('.') + 1);
            var procedure = moduleSymbols.Procedures.FirstOrDefault(
                p => string.Equals(p.Name, procedureName, StringComparison.OrdinalIgnoreCase))
                ?? throw new InvalidOperationException(
                    $"Procedure '{procedureName}' not found in module '{moduleName}'.");

            var instrumentResult = Instrumenter.Instrument(targetModule.Code, procedure, targetModule.Name);
            var instrumentedModule = targetModule with { Code = instrumentResult.InstrumentedSource };
            access.Write(instrumentedModule);

            var agentPath = Path.Combine(AppContext.BaseDirectory, "vba", "modAgent.bas");
            var agentSource = File.ReadAllText(agentPath, Encoding.GetEncoding(1252));
            var agentModule = new VbaModule("Agent", ModuleKind.Standard, agentSource, ".bas");
            access.Write(agentModule);

            var probeSites = instrumentResult.ProbeSites.ToDictionary(site => site.ProbeId);
            var probesFired = new List<ProbeEvent>();

            ProbeCommand WrappedOnProbe(ProbeEvent probeEvent)
            {
                probesFired.Add(probeEvent);
                return onProbe(probeEvent);
            }

            using var server = new ProbeServer(AgentPort, probeSites, WrappedOnProbe, log);
            server.Start();

            RunResult runResult;
            try
            {
                var runner = new Runner(excel, shadowWorkbook.VBProject, log);
                runResult = runner.Run(entryPointQualifiedName);
            }
            finally
            {
                server.Stop();
            }

            return new DebugResult(runResult, probesFired);
        }
        finally
        {
            shadowWorkbook.Close(false);
        }
    }
}
