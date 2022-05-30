using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FluentExcelValidatorStandard.Annotations;

namespace TestConsole
{
    public class TestModel
    {
        [ExcelColumnName("First Name")]
        public string FirstName { get; set; }

        [ExcelColumnName("Last Name")]
        public string      LastName         { get; set; }

    }
}
