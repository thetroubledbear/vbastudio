using System;
using System.Collections.Generic;
using System.Diagnostics;
using VbaStudio.Core.Testing;

namespace VbaStudio.Interop;

/// <summary>
/// Outcome of one test. <paramref name="Skipped"/> distinguishes "this test never ran" (Runner's
/// preconditions refused to start it - typically the project is stuck out of design mode after an
/// earlier test left a runtime-error dialog on Debug) from "this test ran and its assertion failed".
/// Both carry Passed=false; only the former is not a real red result.
/// </summary>
public sealed record TestResult(TestCase Test, bool Passed, string? FailureMessage, TimeSpan Duration, bool Skipped = false);

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
                // Runner.Run's precondition guards (M2) refused to start - typically the project is
                // still stuck in break/run mode from a prior test in this same batch. The test never
                // ran, so it is Skipped, not a red assertion: reporting it as a plain failure would
                // be indistinguishable from a genuine assertion failure. Continue regardless - one
                // stuck test must not silently swallow every test that would have run after it.
                results.Add(new TestResult(test, false, ex.Message, stopwatch.Elapsed, Skipped: true));
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                // Anything else escaping Runner.Run - notably a COMException from the unguarded live
                // COM reads its own preconditions and LogIfNotBackInDesignMode perform (_project.Name,
                // _vbe.ActiveVBProject, _project.Mode). Letting it propagate would discard every result
                // already collected in this batch, since the caller prints nothing on an exception.
                // This is a real failure, not a skip: we cannot tell whether the test body ran.
                results.Add(new TestResult(test, false, ex.Message, stopwatch.Elapsed));
            }
        }

        return results;
    }
}
