using System.Collections.Generic;
using System.Linq;
using VbaStudio.Core.Excel;
using VbaStudio.Core.Model;

namespace VbaStudio.Tests.Fakes;

public sealed class FakeVbaProjectAccess : IVbaProjectAccess
{
    private readonly Dictionary<string, VbaModule> _modules = new();

    public bool IsMacroRunning { get; set; }
    public int WriteCallCount { get; private set; }
    public int DeleteCallCount { get; private set; }

    public void Add(VbaModule module) => _modules[module.Name] = module;

    public IReadOnlyList<VbaModule> ReadAll() => _modules.Values.ToList();

    public void Write(VbaModule module)
    {
        WriteCallCount++;
        _modules[module.Name] = module;
    }

    public void Delete(string name)
    {
        DeleteCallCount++;
        _modules.Remove(name);
    }
}
