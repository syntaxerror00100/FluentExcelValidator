
using FluentExcelValidatorStandard;
using TestConsole;

var settings = new ValidatorSettings()
{
    ColumnHeaderRowNumber = 1,
    DataRowNumber = 3,
    ExcelFIleBytes = File.ReadAllBytes("File\\TestFile.xlsx")
};


var smartValidator = new FluentExcelValidator();

var validator = new TestModelValidator();

var result = await smartValidator.ValidateAsync<TestModel>(validator, settings);



Console.WriteLine($"Is Success:{result.IsValid}\nErrors:{result.Errors.Count}");

foreach (var error in result.Errors)
{
   Console.WriteLine($"Error at row [{error.RowNumber}] Errors:{string.Join("\n", error.ValidationResult.Errors.Select(x => x.ErrorMessage))}"); 
}

Console.ReadLine();
