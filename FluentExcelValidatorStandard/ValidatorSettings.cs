namespace FluentExcelValidatorStandard
{
    public class ValidatorSettings
    {
        public int ColumnHeaderRowNumber { get; set; }
        public int DataRowNumber { get; set; }
        public byte[] ExcelFIleBytes { get; set; }
        public bool IgnoreWhiteSpaces { get; set; } = false;

    }
}
