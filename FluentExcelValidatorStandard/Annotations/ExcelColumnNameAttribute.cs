using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace FluentExcelValidatorStandard.Annotations
{
    /// <summary>
    /// Data attribute for column name
    /// </summary>
    /// <seealso cref="System.ComponentModel.DisplayNameAttribute" />
    public sealed class ExcelColumnNameAttribute : DisplayNameAttribute
    {
        public string ExcelColumnName { get; }

        public ExcelColumnNameAttribute(string excelColumnName) : base(excelColumnName)
        {
            ExcelColumnName = excelColumnName;
        }

    }
}
