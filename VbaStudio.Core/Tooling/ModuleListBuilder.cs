using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;
using VbaStudio.Core.Model;
using VbaStudio.Core.Parsing;

namespace VbaStudio.Core.Tooling;

public sealed record ModuleListing(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("procedures")] IReadOnlyList<string> Procedures);

public sealed record ModuleListResult(
    [property: JsonPropertyName("workbookPath")] string WorkbookPath,
    [property: JsonPropertyName("modules")] IReadOnlyList<ModuleListing> Modules);

public static class ModuleListBuilder
{
    // Application.Run("Module.Procedure") - the only mechanism this codebase launches VBA
    // entry points through - can only call a Sub or Function in a standard module, and can never
    // supply arguments. Class module members have no instance to call them on; a required
    // parameter has no value to be supplied; a Property Get/Let/Set is not addressable by
    // Application.Run at all. Listing any of them as a runnable target produces a deterministic
    // failure from Application.Run (DISP_E_PARAMNOTOPTIONAL - "Parameter not optional" - for the
    // parameter case, confirmed live against a real workbook) - this filter is what keeps the
    // picker showing only entries that can actually be launched.
    public static ModuleListResult Build(string workbookPath, IReadOnlyList<VbaModule> modules)
    {
        var listings = modules
            .Where(m => m.Kind == ModuleKind.Standard)
            .Select(m => new ModuleListing(
                m.Name,
                VbaParser.ParseModule(m.Code, m.Name).Procedures
                    .Where(p => p.Kind == ProcedureKind.Sub || p.Kind == ProcedureKind.Function)
                    .Where(p => p.Parameters.All(param => param.IsOptional))
                    .Select(p => p.Name)
                    .ToList()))
            .ToList();

        return new ModuleListResult(workbookPath, listings);
    }
}
