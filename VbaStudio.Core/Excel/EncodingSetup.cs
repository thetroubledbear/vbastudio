using System.Runtime.CompilerServices;
using System.Text;

namespace VbaStudio.Core.Excel;

internal static class EncodingSetup
{
    [ModuleInitializer]
    public static void Register()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }
}
