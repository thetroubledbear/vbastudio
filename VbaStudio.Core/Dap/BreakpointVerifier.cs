using System;
using System.Collections.Generic;
using System.Linq;
using VbaStudio.Core.Instrumentation;

namespace VbaStudio.Core.Dap;

public static class BreakpointVerifier
{
    public static IReadOnlyList<DapBreakpoint> ComputeVerifiedBreakpoints(
        IReadOnlyList<ProbeSite> probeSites, string moduleName, IReadOnlyList<int> requestedLines)
    {
        var validLines = probeSites
            .Where(s => string.Equals(s.ModuleName, moduleName, StringComparison.OrdinalIgnoreCase))
            .Select(s => s.OriginalLine)
            .ToHashSet();

        return requestedLines
            .Select(line => new DapBreakpoint(Verified: validLines.Contains(line), Line: line))
            .ToList();
    }
}
