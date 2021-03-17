namespace CSICorp.Web.Client.Helpers
{
    using System.Drawing;
    using OfficeOpenXml;
    using OfficeOpenXml.Style;

    public static class ExcelTableExtension
    {
        public static void HeaderTitle(this ExcelRange excelRange)
        {
            excelRange.Merge = true;
            excelRange.Style.Font.Bold = true;
            excelRange.Style.Font.Size = 14;
        }

        public static void TableTitle(this ExcelRange excelRange)
        {
            excelRange.Merge = true;
            excelRange.Style.Font.Bold = true;
            excelRange.Style.Font.Size = 12;
            excelRange.Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }

        public static void Centered(this ExcelRange excelRange)
        {
            excelRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }
        
        public static void DefaultDoubleFormat(this ExcelRange excelRange)
        {
            excelRange.Style.Numberformat.Format = "0.000";
        }

        public static void SetCellValue<T>(this ExcelWorksheet worksheet, int row, int col, T value)
        {
            var cell = worksheet.Cells[row, col];
            cell.Centered();
            if(value.GetType() is double)
                cell.DefaultDoubleFormat();
            cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            cell.Value = value;
        }
        public static void Colored(this ExcelWorksheet worksheet, int row, int col, Color color)
        {
            var cell = worksheet.Cells[row, col];
            cell.Style.Font.Color.SetColor(color);
        }
    }
}