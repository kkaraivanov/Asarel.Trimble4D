namespace CSICorp.Web.Client.Helpers
{
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
    }
}