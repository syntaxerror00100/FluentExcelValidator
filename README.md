# FluentExcelValidator
Clean and easy way to validate excel data using fluentvalidation

# Install
Install-Package FluentExcelValidator

# Sample usage
```
var settings = new ValidatorSettings()
{
    ColumnHeaderRowNumber = 1,
    DataRowNumber = 3,
    ExcelFIleBytes = File.ReadAllBytes("File\\TestFile.xlsx")
};

var smartValidator = new FluentExcelValidator();
var validator = new TestModelValidator();
var result = await smartValidator.ValidateAsync<TestModel>(validator, settings);
```
# Rule
All the properties of your excel model must be string
