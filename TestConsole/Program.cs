
using FluentExcelValidatorStandard;
using TestConsole;

var settings = new ValidatorSettings()
{
    ColumnHeaderRowNumber = 1,
    DataRowNumber = 3,
    ExcelFIleBytes = File.ReadAllBytes("D:\\TEMP\\SmartTest.xlsx")
};


var smartValidator = new FluentExcelValidator();

var validator = new TestModelValidator();

var result = await smartValidator.ValidateAsync<TestModel>(validator, settings);

// See https://aka.ms/new-console-template for more information
Console.WriteLine("Hello, World!");

Console.ReadLine();
