// VbaStudio.Core/Parsing/Symbols.cs
using System.Collections.Generic;

namespace VbaStudio.Core.Parsing;

public enum SymbolKind { Parameter, Local, Const, ModuleVariable }

public enum ProcedureKind { Sub, Function, PropertyGet, PropertyLet, PropertySet }

public enum ProcedureVisibility { Public, Private, Friend }

public sealed record Symbol(
    string Name,
    string DeclaredType,
    SymbolKind Kind,
    bool IsArray,
    bool IsOptional,
    string? PassingMode);

public sealed record ProcedureSymbols(
    string Name,
    ProcedureKind Kind,
    ProcedureVisibility Visibility,
    int StartLine,
    int EndLine,
    IReadOnlyList<Symbol> Parameters,
    IReadOnlyList<Symbol> Locals);

public sealed record ModuleSymbols(
    string ModuleName,
    IReadOnlyList<Symbol> ModuleVariables,
    IReadOnlyList<ProcedureSymbols> Procedures);
