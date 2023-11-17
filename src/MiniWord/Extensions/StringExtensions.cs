using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
}
