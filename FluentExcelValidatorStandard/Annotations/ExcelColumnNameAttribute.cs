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
        private readonly string _excelColumnName;
        public ExcelColumnNameAttribute(string excelColumnName) : base(excelColumnName)
        {
            _excelColumnName = excelColumnName;
        }

        public string ExcelColumnName => _excelColumnName;
        
    }
}
