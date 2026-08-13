using System;
using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using System.Text;
using VbaStudio.Core.Excel;
using VbaStudio.Core.Model;

namespace VbaStudio.Core.Sync;

public sealed class SyncEngine
{
    private static readonly HashSet<string> SyncedExtensions = new(StringComparer.OrdinalIgnoreCase) { ".bas", ".cls", ".frm" };

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
            var path = _fileSystem.Path.Combine(_srcDir, module.FileName);
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

        foreach (var path in _fileSystem.Directory.EnumerateFiles(_srcDir))
        {
            var extension = _fileSystem.Path.GetExtension(path);
            if (!SyncedExtensions.Contains(extension))
            {
                continue;
            }

            var name = _fileSystem.Path.GetFileNameWithoutExtension(path);
            var kind = existingByName.TryGetValue(name, out var current)
                ? current.Kind
                : InferKindFromExtension(extension);
            var encoding = kind.SourceEncoding();
            var code = _fileSystem.File.ReadAllText(path, encoding);

            if (existingByName.TryGetValue(name, out var unchanged) && unchanged.Code == code)
            {
                continue;
            }

            _access.Write(new VbaModule(name, kind, code, extension));
        }
    }

    private static ModuleKind InferKindFromExtension(string extension) => extension switch
    {
        ".bas" => ModuleKind.Standard,
        ".frm" => ModuleKind.UserForm,
        ".cls" => ModuleKind.Class,
        _ => throw new ArgumentOutOfRangeException(nameof(extension), extension, "Unrecognized VBA source extension.")
    };
}
