using System.Collections.Generic;
using VbaStudio.Core.Model;

namespace VbaStudio.Core.Excel;

public interface IVbaProjectAccess
{
    IReadOnlyList<VbaModule> ReadAll();
    void Write(VbaModule module);
    void Delete(string name);
    bool IsMacroRunning { get; }
}
