using System.IO.Abstractions.TestingHelpers;
using VbaStudio.Core.Model;
using VbaStudio.Core.Sync;
using Xunit;

namespace VbaStudio.Tests.Sync;

public class SyncManifestTests
{
    [Fact]
    public void Load_MissingFile_ReturnsEmptyManifest()
    {
        var fs = new MockFileSystem();

        var manifest = SyncManifest.Load(fs, "src/.vbastudio-sync.json");

        Assert.Empty(manifest.Modules);
    }

    [Fact]
    public void SaveThenLoad_RoundTrips_Modules()
    {
        var fs = new MockFileSystem();
        var manifest = new SyncManifest();
        manifest.Modules["modCalc"] = new ManifestEntry(ModuleKind.Standard, "abc123");

        manifest.Save(fs, "src/.vbastudio-sync.json");
        var loaded = SyncManifest.Load(fs, "src/.vbastudio-sync.json");

        var entry = Assert.Single(loaded.Modules);
        Assert.Equal("modCalc", entry.Key);
        Assert.Equal(ModuleKind.Standard, entry.Value.Kind);
        Assert.Equal("abc123", entry.Value.Hash);
    }

    [Fact]
    public void ComputeHash_DifferingOnlyByLineEndings_ProducesSameHash()
    {
        var crlf = SyncManifest.ComputeHash("Sub Foo()\r\nEnd Sub");
        var lf = SyncManifest.ComputeHash("Sub Foo()\nEnd Sub");

        Assert.Equal(crlf, lf);
    }

    [Fact]
    public void ComputeHash_DifferentContent_ProducesDifferentHash()
    {
        var a = SyncManifest.ComputeHash("Sub Foo()\r\nEnd Sub");
        var b = SyncManifest.ComputeHash("Sub Bar()\r\nEnd Sub");

        Assert.NotEqual(a, b);
    }
}
