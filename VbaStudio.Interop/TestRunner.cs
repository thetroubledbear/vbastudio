using System;
using System.Collections.Generic;
using System.Diagnostics;
using VbaStudio.Core.Testing;

namespace VbaStudio.Interop;

public sealed record TestResult(TestCase Test, bool Passed, string? FailureMessage, TimeSpan Duration);

public sealed class TestRunner
{
    private readonly Runner _runner;

    public TestRunner(Runner runner)
    {
        _runner = runner;
    }

    public IReadOnlyList<TestResult> RunAll(IEnumerable<TestCase> tests)
    {
        var results = new List<TestResult>();

        foreach (var test in tests)
        {
            var stopwatch = Stopwatch.StartNew();
            try
            {
                var runResult = _runner.Run(test.QualifiedName);
                stopwatch.Stop();
                results.Add(new TestResult(test, runResult.Success, runResult.Diagnostic?.Message, stopwatch.Elapsed));
            }
            catch (InvalidOperationException ex)
            {
                stopwatch.Stop();
                // Runner.Run's VBProject.Mode precondition guard (M2) refused to start - the
                // project is still stuck in break/run mode from a prior test in this same batch.
                // Record it as a failure and continue: one stuck test must not silently swallow
                // every test that would have run after it.
                results.Add(new TestResult(test, false, ex.Message, stopwatch.Elapsed));
            }
        }

        return results;
    }
}
