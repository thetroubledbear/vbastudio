using System.IO.Abstractions;

namespace VbaStudio.Core.Shadow;

public static class ShadowWorkbook
{
    public static void CreateFromClosed(IFileSystem fileSystem, string sourcePath, string shadowPath)
    {
        var shadowDir = fileSystem.Path.GetDirectoryName(shadowPath);
        if (!string.IsNullOrEmpty(shadowDir) && !fileSystem.Directory.Exists(shadowDir))
        {
            fileSystem.Directory.CreateDirectory(shadowDir);
        }
        fileSystem.File.Copy(sourcePath, shadowPath, overwrite: true);
    }
}
