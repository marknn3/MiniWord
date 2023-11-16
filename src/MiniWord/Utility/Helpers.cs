using System.IO;

namespace MiniSoftware.Utility;

internal static partial class Helpers
{
    public static FileStream OpenSharedRead(string path) => File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
}