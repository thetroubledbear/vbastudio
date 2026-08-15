using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using VbaStudio.Core.Model;
using VbaStudio.Core.Tooling;

namespace VbaStudio.Core.Sync;

public sealed record ManifestEntry(ModuleKind Kind, string Hash);

// The last-synced state both sides (disk and Excel) are known to have agreed on. SyncEngine
// diffs each side's current hash against this baseline to tell "changed since last sync" apart
// from "changed relative to the other side" - without it, Push/Pull can only compare disk to
// Excel directly, which can't distinguish a real conflict from an ordinary one-sided edit.
public sealed class SyncManifest
{
    public Dictionary<string, ManifestEntry> Modules { get; init; } = new();

    public static SyncManifest Load(IFileSystem fileSystem, string path)
    {
        if (!fileSystem.File.Exists(path))
        {
            return new SyncManifest();
        }

        var json = fileSystem.File.ReadAllText(path);
        return JsonSerializer.Deserialize<SyncManifest>(json) ?? new SyncManifest();
    }

    public void Save(IFileSystem fileSystem, string path)
    {
        var directory = fileSystem.Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(directory) && !fileSystem.Directory.Exists(directory))
        {
            fileSystem.Directory.CreateDirectory(directory);
        }

        var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
        fileSystem.File.WriteAllText(path, json);
    }

    public static string ComputeHash(string code)
    {
        var normalized = StaleChecker.Normalize(code);
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(normalized));
        return Convert.ToHexString(bytes);
    }
}
