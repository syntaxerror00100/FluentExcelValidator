using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using FastMember;
using FluentExcelValidatorStandard;
using FluentExcelValidatorStandard.Annotations;
using FluentExcelValidatorStandard.Models;
using FluentValidation;
using FluentValidation.Results;
using NPOI.SS.UserModel;
using Org.BouncyCastle.Asn1.X509;

namespace FluentExcelValidator
{
    public class FluentExcelValidator
    {
        private ValidatorSettings _validatorSettings;
        public async Task<ExcelValidationResult<T>> ValidateAsync<T>(IValidator<T> validator, ValidatorSettings settings) where T : class
        {
            _validatorSettings = settings;
            var result = new ExcelValidationResult<T>();

            var headers = GetColumnNames<T>();
            IWorkbook workbook = WorkbookFactory.Create(new MemoryStream(_validatorSettings.ExcelFIleBytes));
            ISheet workSheet = workbook.GetSheetAt(0);


            try
            {
                var headerValidationResult = ValidateHeaders(workSheet, headers);

                if (headerValidationResult != null && headerValidationResult.ValidationResult.Errors.Any())
                {
                    result.Errors.Add(headerValidationResult);
                    result.IsValid = false;
                    return result;
                }

                var mappedData = ReadAndMappData<T>(workSheet, headers);

                foreach (var data in mappedData)
                {
                    var validationResult = await validator.ValidateAsync(data.Data);
                    result.MappedData.Add(data.Data);
                    if (!validationResult.IsValid)
                    {
                        result.Errors.Add(new ExcelValidationError
                        {
                            RowNumber = data.Row,
                            ValidationResult = validationResult
                        });

                        result.IsValid = false;
                    }
                }
            }
            catch (Exception e)
            {
                result.Exception = e;
            }
            return result;
        }

        ExcelValidationError? ValidateHeaders(ISheet workSheet, List<ObjectPropertyAndColumnName> headers)
        {
            var headerRowCells = workSheet.GetRow(_validatorSettings.ColumnHeaderRowNumber - 1).Cells;

            foreach (var header in headers)
            {
                bool headerFound = false;
                
                if(_validatorSettings.IgnoreWhiteSpaces)
                    headerFound =  headerRowCells.Any(x => x.StringCellValue.ToLower().CleanWhiteSpaces() == header.CustomColumnName.ToLower().CleanWhiteSpaces());
                else
                    headerFound = headerRowCells.Any(x => x.StringCellValue == header.CustomColumnName);

                if (!headerFound)
                {
                    return new ExcelValidationError
                    {
                        RowNumber = _validatorSettings.ColumnHeaderRowNumber,
                        ValidationResult = new ValidationResult()
                        {
                            Errors = new List<ValidationFailure>()
                            {
                                new()
                                {
                                    ErrorMessage = $"Column {header.CustomColumnName} must exist"
                                }
                            }
                        }
                    };
                }
            }

            return null;
        }

        private List<MappedDataModel<T>> ReadAndMappData<T>(ISheet workSheet, List<ObjectPropertyAndColumnName> headers) where T : class
        {
            var mappedResultList = new List<MappedDataModel<T>>();
            IWorkbook workbook = WorkbookFactory.Create(new MemoryStream(_validatorSettings.ExcelFIleBytes));
            var lastRowIndex = workSheet.LastRowNum;

            var headerRow = workSheet.GetRow(_validatorSettings.ColumnHeaderRowNumber - 1);

            if (headerRow == null)
                throw new Exception("Header row is empty");

            var headerRowCells = headerRow.Cells;

            for (int rowIndex = _validatorSettings.DataRowNumber - 1; rowIndex <= lastRowIndex; rowIndex++)
            {
                var excelRow = workSheet.GetRow(rowIndex);

                if (excelRow == null)
                    continue;

                var mappedResult = new MappedDataModel<T>()
                {
                    Row = rowIndex + 1
                };

                var objOfT = Activator.CreateInstance<T>();
                var wrappedOfObjT = ObjectAccessor.Create(objOfT);

                for (int colIndex = 0; colIndex < headerRow.Cells.Count; colIndex++)
                {
                    var excelHeaderValue = headerRowCells[colIndex]?.StringCellValue ?? "";
                    var objectColumn = headers.FirstOrDefault(x =>
                        x.CustomColumnName.Equals(excelHeaderValue, StringComparison.CurrentCultureIgnoreCase));

                    if (objectColumn != null)
                    {
                        var excelRowCell = excelRow.GetCell(colIndex);
                        var cellValue = excelRowCell?.ToString();
                        wrappedOfObjT[objectColumn.PropertyName] = cellValue;
                    }
                }
                mappedResult.Data = objOfT;
                mappedResultList.Add(mappedResult);
            }


            return mappedResultList;
        }


        private List<ObjectPropertyAndColumnName> GetColumnNames<T>() where T : class
        {
            var result = new List<ObjectPropertyAndColumnName>();
            var props = GetOrderedProperties(typeof(T)).ToList();

            for (int i = 0; i < props.Count(); i++)
            {
                var prop = props[i];


                var objectPropertyAndColumnName = new ObjectPropertyAndColumnName
                {
                    PropertyIndex = i,
                    PropertyName = prop.Name,
                    CustomColumnName = prop.Name
                };

                object[] attrs = prop.GetCustomAttributes(true);


                if (attrs != null && attrs.Length != 0)
                {
                    var excelColNameAttr = attrs.FirstOrDefault(x => x is ExcelColumnNameAttribute);
                    if (excelColNameAttr != null)
                        objectPropertyAndColumnName.CustomColumnName = (excelColNameAttr as ExcelColumnNameAttribute)?.ExcelColumnName;

                }

                result.Add(objectPropertyAndColumnName);

            }

            return result;
        }

        private IEnumerable<PropertyInfo> GetOrderedProperties(Type type)
        {
            Dictionary<Type, int> lookup = new Dictionary<Type, int>();

            int count = 0;
            lookup[type] = count++;
            Type parent = type.BaseType;
            while (parent != null)
            {
                lookup[parent] = count;
                count++;
                parent = parent.BaseType;
            }

            return type.GetProperties().OrderByDescending(prop => lookup[prop.DeclaringType]);
        }
    }
}
