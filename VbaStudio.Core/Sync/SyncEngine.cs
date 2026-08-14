using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using VbaStudio.Core.Excel;
using VbaStudio.Core.Model;

namespace VbaStudio.Core.Sync;

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

    public void Pull()
    {
        if (!_fileSystem.Directory.Exists(_srcDir))
        {
            _fileSystem.Directory.CreateDirectory(_srcDir);
        }

        foreach (var module in _access.ReadAll())
        {
            var folder = _fileSystem.Path.Combine(_srcDir, module.Kind.SourceFolder());
            if (!_fileSystem.Directory.Exists(folder))
            {
                _fileSystem.Directory.CreateDirectory(folder);
            }

            var path = _fileSystem.Path.Combine(folder, module.FileName);
            var encoding = module.Kind.SourceEncoding();
            _fileSystem.File.WriteAllText(path, module.Code, encoding);
        }
    }

    public void Push()
    {
        if (_access.IsMacroRunning)
        {
            throw new InvalidOperationException("Refusing to push: a macro is currently running.");
        }

        if (!_fileSystem.Directory.Exists(_srcDir))
        {
            _fileSystem.Directory.CreateDirectory(_srcDir);
            return;
        }

        var existingByName = _access.ReadAll().ToDictionary(m => m.Name);

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
                var code = _fileSystem.File.ReadAllText(path, encoding);

                if (existingByName.TryGetValue(name, out var unchanged) && unchanged.Code == code)
                {
                    continue;
                }

                _access.Write(new VbaModule(name, kind, code, extension));
            }
        }
    }
}
