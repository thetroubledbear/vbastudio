using System;

namespace VbaStudio.Core.Model;

public enum ModuleKind
{
    Standard = 1,
    Class = 2,
    UserForm = 3,
    Document = 100
}

public static class ModuleKindExtensions
{
    public static string FileExtension(this ModuleKind kind) => kind switch
    {
        ModuleKind.Standard => ".bas",
        ModuleKind.Class => ".cls",
        ModuleKind.UserForm => ".frm",
        ModuleKind.Document => ".cls",
        _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, null)
    };
}

public sealed record VbaModule(
    string Name,
    ModuleKind Kind,
    string Code,
    string FileExtension)
{
    public bool IsDocumentModule => Kind == ModuleKind.Document;

    public string FileName => $"{Name}{FileExtension}";
}
