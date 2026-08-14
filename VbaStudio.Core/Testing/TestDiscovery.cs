using System.Collections.Generic;
using System.IO.Abstractions;
using System.Text.RegularExpressions;

namespace VbaStudio.Core.Testing;

public sealed record TestCase(string ModuleName, string ProcedureName)
{
    public string QualifiedName => $"{ModuleName}.{ProcedureName}";
}

public static class TestDiscovery
{
    private static readonly string[] SourceSubfolders = { "Classes", "Modules" };

    private static readonly Regex TestFileNamePattern =
        new(@"Tests\.(cls|bas)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex TestProcedurePattern =
        new(@"^\s*Public\s+Sub\s+(Test_\w+)\s*\(\s*\)\s*$", RegexOptions.Compiled);

    public static IReadOnlyList<TestCase> DiscoverTests(IFileSystem fileSystem, string srcDir)
    {
        var results = new List<TestCase>();

        foreach (var subfolder in SourceSubfolders)
        {
            var folder = fileSystem.Path.Combine(srcDir, subfolder);
            if (!fileSystem.Directory.Exists(folder))
            {
                continue;
            }

            foreach (var path in fileSystem.Directory.EnumerateFiles(folder))
            {
                var fileName = fileSystem.Path.GetFileName(path);
                if (!TestFileNamePattern.IsMatch(fileName))
                {
                    continue;
                }

                var moduleName = fileSystem.Path.GetFileNameWithoutExtension(path);
                foreach (var line in fileSystem.File.ReadAllLines(path))
                {
                    var match = TestProcedurePattern.Match(line);
                    if (match.Success)
                    {
                        results.Add(new TestCase(moduleName, match.Groups[1].Value));
                    }
                }
            }
        }

        return results;
    }
}
