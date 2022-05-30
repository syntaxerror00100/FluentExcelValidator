using FastMember;
using FluentExcelValidatorStandard.Annotations;
using FluentExcelValidatorStandard.Models;
using FluentValidation;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace FluentExcelValidatorStandard
{
    public class FluentExcelValidator
    {
        public FluentExcelValidator()
        {
            DipendencyInjection.ResolveDependencies();
        }

        private ValidatorSettings ValidatorSettings { get; set; }


        public async Task<ExcelValidationResult<T>> ValidateAsync<T>(IValidator<T> validator, ValidatorSettings settings) where T : class
        {
            ValidatorSettings = settings;
            var result = new ExcelValidationResult<T>
            {
                IsValid = true,
            };

            try
            {
                var mappedData = ReadAndMappData<T>();

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


        private List<MappedDataModel<T>> ReadAndMappData<T>() where T : class
        {

            var mappedResultList = new List<MappedDataModel<T>>();
            var objectColumns = GetColumnNames<T>();

            var package = new ExcelPackage(new MemoryStream(ValidatorSettings.ExcelFIleBytes));

            var workSheet = package.Workbook.Worksheets[0];
            var start = workSheet.Dimension.Start;
            var end = workSheet.Dimension.End;
            for (int row = ValidatorSettings.DataRowNumber; row <= end.Row; row++)
            {
                var mappedResult = new MappedDataModel<T>()
                {
                    Row = row
                };

                var objOfT = Activator.CreateInstance<T>();
                var wrappedOfObjT = ObjectAccessor.Create(objOfT);

                for (int col = start.Column; col <= end.Column; col++)
                {
                    var excelHeaderValue = workSheet.Cells[ValidatorSettings.ColumnHeaderRowNumber, col].Text?.Trim();
                    var objectColumn = objectColumns.FirstOrDefault(x =>
                        x.CustomColumnName.Equals(excelHeaderValue, StringComparison.CurrentCultureIgnoreCase));

                    if (objectColumn != null)
                    {
                        var cellValue = workSheet.Cells[row, col].Text?.Trim();
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
