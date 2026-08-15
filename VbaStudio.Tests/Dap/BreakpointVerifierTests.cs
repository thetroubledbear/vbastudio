using System.Collections.Generic;
using VbaStudio.Core.Dap;
using VbaStudio.Core.Instrumentation;

namespace VbaStudio.Tests.Dap;

public class BreakpointVerifierTests
{
    [Fact]
    public void ComputeVerifiedBreakpoints_LineMatchingAProbeSite_IsVerified()
    {
        var probeSites = new[] { new ProbeSite(1, "MODCALC", 42) };

        var result = BreakpointVerifier.ComputeVerifiedBreakpoints(probeSites, "MODCALC", new[] { 42 });

        Assert.Single(result);
        Assert.True(result[0].Verified);
        Assert.Equal(42, result[0].Line);
    }

    [Fact]
    public void ComputeVerifiedBreakpoints_LineNotAProbeSite_IsUnverified()
    {
        var probeSites = new[] { new ProbeSite(1, "MODCALC", 42) };

        var result = BreakpointVerifier.ComputeVerifiedBreakpoints(probeSites, "MODCALC", new[] { 43 });

        Assert.Single(result);
        Assert.False(result[0].Verified);
        Assert.Equal(43, result[0].Line);
    }

    [Fact]
    public void ComputeVerifiedBreakpoints_MatchingLineInDifferentModule_IsUnverified()
    {
        var probeSites = new[] { new ProbeSite(1, "MODCALC", 42) };

        var result = BreakpointVerifier.ComputeVerifiedBreakpoints(probeSites, "MODOTHER", new[] { 42 });

        Assert.False(result[0].Verified);
    }

    [Fact]
    public void ComputeVerifiedBreakpoints_ModuleNameComparisonIsCaseInsensitive()
    {
        var probeSites = new[] { new ProbeSite(1, "ModCalc", 42) };

        var result = BreakpointVerifier.ComputeVerifiedBreakpoints(probeSites, "modcalc", new[] { 42 });

        Assert.True(result[0].Verified);
    }

    [Fact]
    public void ComputeVerifiedBreakpoints_NoProbeSitesYet_AllLinesUnverified()
    {
        var result = BreakpointVerifier.ComputeVerifiedBreakpoints(
            System.Array.Empty<ProbeSite>(), "MODCALC", new[] { 1, 2, 3 });

        Assert.All(result, bp => Assert.False(bp.Verified));
    }
}
