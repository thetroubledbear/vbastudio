using System;
using System.Runtime.InteropServices;

namespace VbaStudio.Core.Excel;

public static class ComHelpers
{
    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(
        ref Guid rclsid,
        IntPtr pvReserved,
        [MarshalAs(UnmanagedType.Interface)] out object ppunk);

    [DllImport("ole32.dll", PreserveSig = false)]
    private static extern void CLSIDFromProgID(
        [MarshalAs(UnmanagedType.LPWStr)] string lpszProgID,
        out Guid lpclsid);

    /// <summary>VBA's GetObject(, progId). Throws if the app isn't running.</summary>
    public static object GetRunningInstance(string progId)
    {
        CLSIDFromProgID(progId, out Guid clsid);
        GetActiveObject(ref clsid, IntPtr.Zero, out object instance);
        return instance;
    }
}
