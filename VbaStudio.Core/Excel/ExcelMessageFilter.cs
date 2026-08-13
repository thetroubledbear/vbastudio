using System.Runtime.InteropServices;

namespace VbaStudio.Core.Excel;

[ComImport, Guid("00000016-0000-0000-C000-000000000046"),
 InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IOleMessageFilter
{
    [PreserveSig] int HandleInComingCall(
        int dwCallType, System.IntPtr hTaskCaller, int dwTickCount, System.IntPtr lpInterfaceInfo);
    [PreserveSig] int RetryRejectedCall(
        System.IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
    [PreserveSig] int MessagePending(
        System.IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
}

public class ExcelMessageFilter : IOleMessageFilter
{
    private const int SERVERCALL_ISHANDLED = 0;
    private const int SERVERCALL_RETRYLATER = 2;
    private const int PENDINGMSG_WAITDEFPROCESS = 2;

    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(
        IOleMessageFilter? newFilter, out IOleMessageFilter? oldFilter);

    public static void Register()
        => CoRegisterMessageFilter(new ExcelMessageFilter(), out _);

    public static void Revoke()
        => CoRegisterMessageFilter(null, out _);

    int IOleMessageFilter.HandleInComingCall(
        int dwCallType, System.IntPtr hTaskCaller, int dwTickCount, System.IntPtr lpInterfaceInfo)
        => SERVERCALL_ISHANDLED;

    int IOleMessageFilter.RetryRejectedCall(
        System.IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
    {
        // Excel is busy. Wait 99ms and try again instead of failing.
        if (dwRejectType == SERVERCALL_RETRYLATER)
            return 99;
        return -1;   // cancel
    }

    int IOleMessageFilter.MessagePending(
        System.IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
        => PENDINGMSG_WAITDEFPROCESS;
}
