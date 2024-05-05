using System.Text.RegularExpressions;

namespace MiniSoftware.Extensions;

internal static class StringExtensions
{
    internal static int Count(this string str, string find)
    {
        int count = 0;
        int pos = 0;
        while ((pos = str.IndexOf(find, pos)) != -1)
        {
            count++;
            pos += find.Length;
        }
        return count;
    }

    internal static string RegexReplace(this string str, string pattern, string replacement)
    {
        return Regex.Replace(str, pattern, replacement);
    }
}
