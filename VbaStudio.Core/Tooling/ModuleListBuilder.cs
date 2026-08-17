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
    // entry points through - can only call a Sub or Function in a standard module, can never
    // supply arguments, and per Microsoft's own documentation must target a Public procedure.
    // Class module members have no instance to call them on; a required parameter has no value
    // to be supplied; a Property Get/Let/Set is not addressable by Application.Run at all; a
    // Private (or Friend - not valid on a standard-module procedure, but handled the same way if
    // it ever appears) procedure is not visible to Application.Run from outside the module.
    // Listing any of them as a runnable target produces a deterministic failure from
    // Application.Run (DISP_E_PARAMNOTOPTIONAL - "Parameter not optional" - for the parameter
    // case, confirmed live against a real workbook) - this filter is what keeps the picker
    // showing only entries that can actually be launched.
    public static ModuleListResult Build(string workbookPath, IReadOnlyList<VbaModule> modules)
    {
        var listings = modules
            .Where(m => m.Kind == ModuleKind.Standard)
            .Select(m => new ModuleListing(
                m.Name,
                VbaParser.ParseModule(m.Code, m.Name).Procedures
                    .Where(p => p.Kind == ProcedureKind.Sub || p.Kind == ProcedureKind.Function)
                    .Where(p => p.Visibility == ProcedureVisibility.Public)
                    .Where(p => p.Parameters.All(param => param.IsOptional))
                    .Select(p => p.Name)
                    .ToList()))
            .ToList();

        return new ModuleListResult(workbookPath, listings);
    }
}
