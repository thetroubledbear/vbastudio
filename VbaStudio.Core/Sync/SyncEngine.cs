using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using VbaStudio.Core.Excel;
using VbaStudio.Core.Model;

namespace VbaStudio.Core.Sync;

// Pull and Push both diff against a persisted baseline (SyncManifest: the hash each module had
// as of the last successful sync) rather than against each other directly. That 3-way comparison
// is what lets a one-sided edit ("only Excel changed") be told apart from a real conflict ("both
// changed, and differently") - a direct disk-vs-Excel compare alone can't make that distinction,
// which is what let Push silently overwrite live Excel edits before this existed.
public sealed class SyncEngine
{
    private static readonly ModuleKind[] AllKinds =
    {
        ModuleKind.Standard, ModuleKind.Class, ModuleKind.UserForm, ModuleKind.Document
    };

    private readonly IVbaProjectAccess _access;
    private readonly IFileSystem _fileSystem;
    private readonly string _srcDir;

    public SyncEngine(IVbaProjectAccess access, IFileSystem fileSystem, string srcDir)
    {
        _access = access;
        _fileSystem = fileSystem;
        _srcDir = srcDir;
    }

    private string ManifestPath => _fileSystem.Path.Combine(_srcDir, ".vbastudio-sync.json");

    public SyncResult Pull()
    {
        if (!_fileSystem.Directory.Exists(_srcDir))
        {
            _fileSystem.Directory.CreateDirectory(_srcDir);
        }

        var manifest = SyncManifest.Load(_fileSystem, ManifestPath);
        var written = new List<string>();
        var deleted = new List<string>();
        var conflicts = new List<string>();

        var excelModules = _access.ReadAll();
        var excelNames = new HashSet<string>(excelModules.Select(m => m.Name));

        foreach (var module in excelModules)
        {
            var folder = _fileSystem.Path.Combine(_srcDir, module.Kind.SourceFolder());
            if (!_fileSystem.Directory.Exists(folder))
            {
                _fileSystem.Directory.CreateDirectory(folder);
            }

            var path = _fileSystem.Path.Combine(folder, module.FileName);
            var encoding = module.Kind.SourceEncoding();
            var excelHash = SyncManifest.ComputeHash(module.Code);

            manifest.Modules.TryGetValue(module.Name, out var baseEntry);

            if (baseEntry == null)
            {
                // Never synced before (first pull, or a module Excel gained since): nothing to
                // compare against, so just take Excel's copy.
                _fileSystem.File.WriteAllText(path, module.Code, encoding);
                written.Add(module.Name);
                manifest.Modules[module.Name] = new ManifestEntry(module.Kind, excelHash);
                continue;
            }

            var diskExists = _fileSystem.File.Exists(path);
            var excelChanged = excelHash != baseEntry.Hash;

            if (!diskExists)
            {
                if (excelChanged)
                {
                    // Deleted locally AND changed in Excel since last sync - resurrecting the
                    // file would discard the local deletion, so this needs a human, not a guess.
                    conflicts.Add(module.Name);
                }
                // else: deleted locally, Excel unchanged - respect the local deletion, don't
                // recreate it. (Propagating that deletion into Excel is Push's job.)
                continue;
            }

            var diskContent = _fileSystem.File.ReadAllText(path, encoding);
            var diskHash = SyncManifest.ComputeHash(diskContent);
            var diskChanged = diskHash != baseEntry.Hash;

            if (excelChanged && !diskChanged)
            {
                _fileSystem.File.WriteAllText(path, module.Code, encoding);
                written.Add(module.Name);
            }
            else if (diskChanged && excelChanged && diskHash != excelHash)
            {
                conflicts.Add(module.Name);
                continue;
            }
            // else: no-op (nothing changed, or disk-only change - leave disk alone either way)

            manifest.Modules[module.Name] = new ManifestEntry(module.Kind, excelHash);
        }

        foreach (var name in manifest.Modules.Keys.ToList())
        {
            if (excelNames.Contains(name))
            {
                continue;
            }

            var entry = manifest.Modules[name];
            var folder = _fileSystem.Path.Combine(_srcDir, entry.Kind.SourceFolder());
            var path = _fileSystem.Path.Combine(folder, name + entry.Kind.FileExtension());

            if (!_fileSystem.File.Exists(path))
            {
                manifest.Modules.Remove(name);
                continue;
            }

            var diskHash = SyncManifest.ComputeHash(_fileSystem.File.ReadAllText(path, entry.Kind.SourceEncoding()));
            if (diskHash == entry.Hash)
            {
                _fileSystem.File.Delete(path);
                deleted.Add(name);
                manifest.Modules.Remove(name);
            }
            else
            {
                // Removed from Excel, but the local copy was also edited since last sync - keep
                // the file and let a human reconcile it rather than deleting an unsynced edit.
                conflicts.Add(name);
            }
        }

        manifest.Save(_fileSystem, ManifestPath);
        return new SyncResult(written, deleted, conflicts);
    }

    public SyncResult Push()
    {
        if (_access.IsMacroRunning)
        {
            throw new InvalidOperationException("Refusing to push: a macro is currently running.");
        }

        if (!_fileSystem.Directory.Exists(_srcDir))
        {
            _fileSystem.Directory.CreateDirectory(_srcDir);
            return new SyncResult(Array.Empty<string>(), Array.Empty<string>(), Array.Empty<string>());
        }

        var manifest = SyncManifest.Load(_fileSystem, ManifestPath);
        var written = new List<string>();
        var deleted = new List<string>();
        var conflicts = new List<string>();

        var excelByName = _access.ReadAll().ToDictionary(m => m.Name);
        var diskNames = new HashSet<string>();

        foreach (var kind in AllKinds)
        {
            var folder = _fileSystem.Path.Combine(_srcDir, kind.SourceFolder());
            if (!_fileSystem.Directory.Exists(folder))
            {
                continue;
            }

            var extension = kind.FileExtension();
            var encoding = kind.SourceEncoding();

            foreach (var path in _fileSystem.Directory.EnumerateFiles(folder))
            {
                if (!string.Equals(_fileSystem.Path.GetExtension(path), extension, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var name = _fileSystem.Path.GetFileNameWithoutExtension(path);
                diskNames.Add(name);

                var code = _fileSystem.File.ReadAllText(path, encoding);
                var diskHash = SyncManifest.ComputeHash(code);

                manifest.Modules.TryGetValue(name, out var baseEntry);
                excelByName.TryGetValue(name, out var excelModule);

                if (baseEntry == null)
                {
                    // Never synced before: fall back to a direct compare against Excel's current
                    // copy (matches the old no-baseline behavior) rather than assuming a conflict.
                    if (excelModule == null || excelModule.Code != code)
                    {
                        _access.Write(new VbaModule(name, kind, code, extension));
                        written.Add(name);
                    }
                    manifest.Modules[name] = new ManifestEntry(kind, diskHash);
                    continue;
                }

                if (excelModule == null)
                {
                    // Excel-side deletion Pull hasn't propagated yet, and the disk copy is still
                    // here - Push's only job is disk -> Excel, so it just recreates it.
                    _access.Write(new VbaModule(name, kind, code, extension));
                    written.Add(name);
                    manifest.Modules[name] = new ManifestEntry(kind, diskHash);
                    continue;
                }

                var excelHash = SyncManifest.ComputeHash(excelModule.Code);
                var diskChanged = diskHash != baseEntry.Hash;
                var excelChanged = excelHash != baseEntry.Hash;

                if (diskChanged && !excelChanged)
                {
                    _access.Write(new VbaModule(name, kind, code, extension));
                    written.Add(name);
                }
                else if (diskChanged && excelChanged && diskHash != excelHash)
                {
                    conflicts.Add(name);
                    continue;
                }
                // else: no-op (nothing changed, converged to the same content, or Excel changed
                // and disk didn't - pushing stale disk content would clobber the live Excel edit)

                manifest.Modules[name] = new ManifestEntry(kind, diskHash);
            }
        }

        foreach (var name in manifest.Modules.Keys.ToList())
        {
            if (diskNames.Contains(name))
            {
                continue;
            }

            var entry = manifest.Modules[name];
            if (!excelByName.TryGetValue(name, out var excelModule))
            {
                manifest.Modules.Remove(name);
                continue;
            }

            if (entry.Kind == ModuleKind.Document)
            {
                // Document modules (sheet code-behind) can't be removed via the VBE API - only
                // cleared - so a deleted disk file for one is left alone rather than acted on.
                continue;
            }

            var excelHash = SyncManifest.ComputeHash(excelModule.Code);
            if (excelHash == entry.Hash)
            {
                _access.Delete(name);
                deleted.Add(name);
                manifest.Modules.Remove(name);
            }
            else
            {
                // Deleted locally, but Excel's copy also changed since last sync - keep it and
                // let a human reconcile rather than deleting an edit that was never synced.
                conflicts.Add(name);
            }
        }

        manifest.Save(_fileSystem, ManifestPath);
        return new SyncResult(written, deleted, conflicts);
    }
}
