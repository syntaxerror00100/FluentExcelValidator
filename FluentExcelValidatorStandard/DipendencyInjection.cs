using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml;

namespace FluentExcelValidatorStandard
{
    internal static class DipendencyInjection
    {
        public static void ResolveDependencies()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }
    }
}
