// VbaStudio.Tests/Tooling/StaleCheckerTests.cs
using VbaStudio.Core.Tooling;
using Xunit;

namespace VbaStudio.Tests.Tooling;

public class StaleCheckerTests
{
    [Fact]
    public void IsStale_DiskMatchesExcel_ReturnsFalse()
    {
        Assert.False(StaleChecker.IsStale("Sub A()\r\nEnd Sub\r\n", "Sub A()\r\nEnd Sub\r\n"));
    }

    [Fact]
    public void IsStale_DiskDiffersFromExcel_ReturnsTrue()
    {
        Assert.True(StaleChecker.IsStale("Sub A()\r\nEnd Sub\r\n", "Sub B()\r\nEnd Sub\r\n"));
    }

    [Fact]
    public void IsStale_DiskFileMissing_ReturnsTrue()
    {
        Assert.True(StaleChecker.IsStale(null, "Sub A()\r\nEnd Sub\r\n"));
    }
}
