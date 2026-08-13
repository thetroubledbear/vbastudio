using System;
using System.Text;

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

    public static Encoding SourceEncoding(this ModuleKind kind) => kind switch
    {
        ModuleKind.Document => new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
        _ => Encoding.GetEncoding(1252, EncoderFallback.ExceptionFallback, DecoderFallback.ReplacementFallback)
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
