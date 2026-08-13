using System.Runtime.InteropServices;

namespace VbaStudio.Core.Excel;

public static class ComRelease
{
    public static void Release(params object?[] comObjects)
    {
        foreach (var o in comObjects)
        {
            if (o != null && Marshal.IsComObject(o))
            {
                Marshal.ReleaseComObject(o);
            }
        }
    }
}
