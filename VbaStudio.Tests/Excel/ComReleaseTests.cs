using VbaStudio.Core.Excel;
using Xunit;

namespace VbaStudio.Tests.Excel;

public class ComReleaseTests
{
    [Fact]
    public void Release_IgnoresNullAndNonComObjects()
    {
        var plainObject = new object();
        var exception = Record.Exception(() => ComRelease.Release(null, plainObject, "a string"));
        Assert.Null(exception);
    }
}
