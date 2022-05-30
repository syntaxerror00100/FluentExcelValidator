using System;
using System.Collections.Generic;
using System.Text;

namespace FluentExcelValidatorStandard.Models
{
    internal class MappedDataModel<T> 
    {
        public int Row { get; set; }
        public T   Data   { get; set; }
    }
}
