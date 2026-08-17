namespace VbaStudio.Core.Model;

public sealed record Diagnostic(string? Module, int? Line, string Message);
