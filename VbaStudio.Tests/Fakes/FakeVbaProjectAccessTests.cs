using System.Linq;
using VbaStudio.Core.Model;
using Xunit;

namespace VbaStudio.Tests.Fakes;

public class FakeVbaProjectAccessTests
{
    [Fact]
    public void ReadAll_ReturnsAddedModules()
    {
        var fake = new FakeVbaProjectAccess();
        fake.Add(new VbaModule("modCalc", ModuleKind.Standard, "Sub Foo()\r\nEnd Sub", ".bas"));

        var all = fake.ReadAll();

        Assert.Single(all);
        Assert.Equal("modCalc", all[0].Name);
    }

    [Fact]
    public void Write_ReplacesExistingModuleByName()
    {
        var fake = new FakeVbaProjectAccess();
        fake.Add(new VbaModule("modCalc", ModuleKind.Standard, "old", ".bas"));

        fake.Write(new VbaModule("modCalc", ModuleKind.Standard, "new", ".bas"));

        Assert.Equal("new", fake.ReadAll().Single().Code);
    }

    [Fact]
    public void Write_IncrementsWriteCallCount()
    {
        var fake = new FakeVbaProjectAccess();
        fake.Add(new VbaModule("modCalc", ModuleKind.Standard, "old", ".bas"));

        fake.Write(new VbaModule("modCalc", ModuleKind.Standard, "new", ".bas"));

        Assert.Equal(1, fake.WriteCallCount);
    }

    [Fact]
    public void IsMacroRunning_DefaultsFalse()
    {
        var fake = new FakeVbaProjectAccess();
        Assert.False(fake.IsMacroRunning);
    }
}
