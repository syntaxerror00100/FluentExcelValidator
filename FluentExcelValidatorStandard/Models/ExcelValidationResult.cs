using System;
using System.Collections.Generic;
using System.Text;
using FluentValidation.Results;

namespace FluentExcelValidatorStandard.Models
{
    public class ExcelValidationResult<T>
    {
        public bool IsValid { get; set; }
        public List<T> MappedData { get; set; } = new List<T>();
        public List<ExcelValidationError> Errors { get; set; } = new List<ExcelValidationError>();
        public Exception Exception { get; set; }
    }

    public class ExcelValidationError
    {
        public int RowNumber { get; set; }
        public ValidationResult ValidationResult { get; set; }
    }
}
