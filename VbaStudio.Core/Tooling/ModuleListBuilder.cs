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
    public static ModuleListResult Build(string workbookPath, IReadOnlyList<VbaModule> modules)
    {
        var listings = modules
            .Select(m => new ModuleListing(
                m.Name,
                VbaParser.ParseModule(m.Code, m.Name).Procedures.Select(p => p.Name).ToList()))
            .ToList();

        return new ModuleListResult(workbookPath, listings);
    }
}
