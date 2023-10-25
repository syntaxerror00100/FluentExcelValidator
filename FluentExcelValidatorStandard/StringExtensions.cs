using System;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.Formula.Functions;

namespace FluentExcelValidator
{
    internal static class StringExtensions
    {
        public static string CleanWhiteSpaces(this string val)
        {
            if (string.IsNullOrEmpty(val))
                return val;

            return val.Replace(" ", "").Replace(Environment.NewLine,"").Replace("\n","").Trim();
        }
    }
}
