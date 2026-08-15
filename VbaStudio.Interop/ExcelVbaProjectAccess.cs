// VbaStudio.Interop/ExcelVbaProjectAccess.cs
using System;
using System.Collections.Generic;
using System.IO;
using VbaStudio.Core.Excel;
using VbaStudio.Core.Model;

namespace VbaStudio.Interop;

public sealed class ExcelVbaProjectAccess : IVbaProjectAccess
{
    private readonly VBIDE.VBProject _project;

    public ExcelVbaProjectAccess(VBIDE.VBProject project)
    {
        _project = project;
    }

    // Real run-state detection needs a runner (M2). Until then, refuse nothing here;
    // SyncEngine.Push already gates on this - keep it truthful by not pretending to know.
    public bool IsMacroRunning => false;

    public IReadOnlyList<VbaModule> ReadAll()
    {
        var result = new List<VbaModule>();
        var components = _project.VBComponents;
        try
        {
            foreach (VBIDE.VBComponent component in components)
            {
                try
                {
                    if (Enum.IsDefined(typeof(ModuleKind), (int)component.Type))
                    {
                        result.Add(ReadComponent(component));
                    }
                    // else: component type not modeled by ModuleKind (e.g. ActiveXDesigner) - skip it,
                    // don't let one unusual component abort the whole project's sync.
                }
                finally
                {
                    ComRelease.Release(component);
                }
            }
        }
        finally
        {
            ComRelease.Release(components);
        }
        return result;
    }

    private static VbaModule ReadComponent(VBIDE.VBComponent component)
    {
        var kind = (ModuleKind)component.Type;
        var name = component.Name;
        var extension = kind.FileExtension();

        string code = kind == ModuleKind.Document
            ? ReadDocumentModuleCode(component)
            : ReadExportedModuleCode(component, kind);

        return new VbaModule(name, kind, code, extension);
    }

    private static string ReadDocumentModuleCode(VBIDE.VBComponent component)
    {
        var codeModule = component.CodeModule;
        try
        {
            var lineCount = codeModule.CountOfLines;
            return lineCount > 0 ? codeModule.get_Lines(1, lineCount) : string.Empty;
        }
        finally
        {
            ComRelease.Release(codeModule);
        }
    }

    private static string ReadExportedModuleCode(VBIDE.VBComponent component, ModuleKind kind)
    {
        var tempPath = Path.Combine(Path.GetTempPath(), component.Name + kind.FileExtension());
        component.Export(tempPath);
        var rawLines = File.ReadAllLines(tempPath, kind.SourceEncoding());
        File.Delete(tempPath);

        if (kind == ModuleKind.UserForm)
        {
            var originalLines = ReadCodeModuleLines(component);
            rawLines = VbaSourceText.TrimExtraLeadingBlankLine(originalLines, rawLines);
        }

        return string.Join(Environment.NewLine, rawLines);
    }

    private static string[] ReadCodeModuleLines(VBIDE.VBComponent component)
    {
        var codeModule = component.CodeModule;
        try
        {
            var lineCount = codeModule.CountOfLines;
            var content = lineCount > 0 ? codeModule.get_Lines(1, lineCount) : string.Empty;
            return content.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
        }
        finally
        {
            ComRelease.Release(codeModule);
        }
    }

    public void Write(VbaModule module)
    {
        switch (module.Kind)
        {
            case ModuleKind.Document:
                WriteDocumentModule(module);
                break;
            case ModuleKind.UserForm:
                WriteUserFormModule(module);
                break;
            default:
                WriteByRemoveImport(module);
                break;
        }
    }

    private void WriteDocumentModule(VbaModule module)
    {
        var component = FindComponent(module.Name)
            ?? throw new InvalidOperationException(
                $"Document module '{module.Name}' does not exist in the live project. " +
                "Document modules (sheets, ThisWorkbook) cannot be created by sync - add the sheet in Excel first.");
        try
        {
            ReplaceCodeModuleContent(component, module.Code);
        }
        finally
        {
            ComRelease.Release(component);
        }
    }

    private void WriteUserFormModule(VbaModule module)
    {
        var lines = module.Code.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
        var codeOnly = VbaSourceText.RemoveHeaderAttributeLines(lines);
        var codeString = string.Join(Environment.NewLine, codeOnly);

        var existing = FindComponent(module.Name);
        if (existing == null)
        {
            ImportFromTempFile(module);
            return;
        }

        try
        {
            ReplaceCodeModuleContent(existing, codeString);
        }
        finally
        {
            ComRelease.Release(existing);
        }
    }

    public void Delete(string name)
    {
        var existing = FindComponent(name)
            ?? throw new InvalidOperationException($"Cannot delete module '{name}': it does not exist in the live project.");
        try
        {
            var components = _project.VBComponents;
            try
            {
                components.Remove(existing);
            }
            finally
            {
                ComRelease.Release(components);
            }
        }
        finally
        {
            ComRelease.Release(existing);
        }
    }

    private void WriteByRemoveImport(VbaModule module)
    {
        var existing = FindComponent(module.Name);
        if (existing != null)
        {
            var components = _project.VBComponents;
            try
            {
                components.Remove(existing);
            }
            finally
            {
                ComRelease.Release(components);
            }
            ComRelease.Release(existing);
        }

        ImportFromTempFile(module);
    }

    private void ImportFromTempFile(VbaModule module)
    {
        var tempPath = Path.Combine(Path.GetTempPath(), module.Name + module.FileExtension);
        File.WriteAllText(tempPath, module.Code, module.Kind.SourceEncoding());
        var components = _project.VBComponents;
        try
        {
            var imported = components.Import(tempPath);
            ComRelease.Release(imported);
        }
        finally
        {
            ComRelease.Release(components);
        }
        File.Delete(tempPath);
    }

    private static void ReplaceCodeModuleContent(VBIDE.VBComponent component, string code)
    {
        var codeModule = component.CodeModule;
        try
        {
            if (codeModule.CountOfLines > 0)
            {
                codeModule.DeleteLines(1, codeModule.CountOfLines);
            }
            if (!string.IsNullOrEmpty(code))
            {
                codeModule.AddFromString(code);
            }
        }
        finally
        {
            ComRelease.Release(codeModule);
        }
    }

    private VBIDE.VBComponent? FindComponent(string name)
    {
        var components = _project.VBComponents;
        try
        {
            foreach (VBIDE.VBComponent component in components)
            {
                if (component.Name == name)
                {
                    return component;
                }
                ComRelease.Release(component);
            }
            return null;
        }
        finally
        {
            ComRelease.Release(components);
        }
    }
}
