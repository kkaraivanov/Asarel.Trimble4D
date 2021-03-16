namespace CSICorp.Web.Client.Helpers
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Models;
    using OfficeOpenXml;
    using OfficeOpenXml.Style;

    public static class WeeklyTable
    {
        private const string TABLE_ONE_TITLE = "4. Повърхностен отток на реките [l/s]";
        private const string TABLE_TWO_TITLE = "5. Водни нива на сондажните кладенци [m]";
        private const string TABLE_THREE_TITLE = "8. Промяна на водните нива в наблюдателните сондажи";

        private const int START_COL = 1;
        private const int END_COL = 10;

        private static string _fileName;

        public static async Task CreateTable(
            ExcelWorksheet worksheet,
            string fileName,
            SensorTable currentPeriodeWaterLevel, 
            SensorTable beforePeriodeWaterLevel,
            List<WaterLevelSensors> waterSensorDataList)
        {
            var row = 1;
            _fileName = fileName;
            var tableOne = CreateTableOne(worksheet, TABLE_ONE_TITLE, row);
            var tableTwo = CreateTableTwo(
                worksheet,
                TABLE_TWO_TITLE,
                tableOne,
                currentPeriodeWaterLevel,
                beforePeriodeWaterLevel,
                waterSensorDataList.Where(x => x.Name.Contains("SK")).ToList());

            worksheet.DefaultColWidth = 12;
            worksheet.Column(START_COL).Width = 16;
        }

        private static int CreateTableOne(ExcelWorksheet worksheet, string title, int row)
        {
            var tableTitle = worksheet.Cells[row, START_COL, row, END_COL];
            tableTitle.TableTitle();
            tableTitle.Value = title;

            return row + 2;
        }

        private static int CreateTableTwo(
            ExcelWorksheet worksheet,
            string title,
            int row,
            SensorTable currentPeriodeWaterLevel,
            SensorTable beforePeriodeWaterLevel,
            List<WaterLevelSensors> waterSensorDataList)
        {
            var tableTitle = worksheet.Cells[row, START_COL, row, END_COL];
            tableTitle.TableTitle();
            tableTitle.Value = title;
            row++;
            row = CreateTableHeader(worksheet, row);
            row++;

            var col4 = worksheet.Cells[row, START_COL + 3];
            col4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            col4.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            col4.Value = "Кота [m]";
            var col5 = worksheet.Cells[row, START_COL + 4];
            col5.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            col5.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            col5.Value = "WL [m]";
            var col6 = worksheet.Cells[row, START_COL + 5];
            col6.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            col6.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            col6.Value = "Кота [m]";
            var col7 = worksheet.Cells[row, START_COL + 6];
            col7.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            col7.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            col7.Value = "WL [m]";
            var col8 = worksheet.Cells[row, START_COL + 7];
            col8.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            col8.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            col8.Value = "Разлика";
            row++;



            return row + 2;
        }

        private static int CreateTableHeader(ExcelWorksheet worksheet, int row)
        {
            var col1 = worksheet.Cells[row, START_COL, row + 1, START_COL];
            col1.Merge = true;
            col1.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            col1.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            col1.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            col1.Value = "Номер сондаж";

            var col2 = worksheet.Cells[row, START_COL + 1];
            col2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            col2.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            col2.Value = "Кота";
            var bottomCol2 = worksheet.Cells[row + 1, START_COL + 1];
            bottomCol2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            bottomCol2.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            bottomCol2.Value = "Устие";

            var col3 = worksheet.Cells[row, START_COL + 2];
            col3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            col3.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            col3.Value = "Дълбочина";
            var bottomCol3 = worksheet.Cells[row + 1, START_COL + 2];
            bottomCol3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            bottomCol3.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            bottomCol3.Value = "[m]";

            var col4 = worksheet.Cells[row, START_COL + 3, row, START_COL + 4];
            col4.Merge = true;
            col4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            col4.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            int.TryParse(_fileName[0].ToString(), out int week);
            col4.Value = $"{week - 1} седмица";

            var col5 = worksheet.Cells[row, START_COL + 5, row, START_COL + 7];
            col5.Merge = true;
            col5.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            col5.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            col5.Value = $"{week} седмица";

            var col6 = worksheet.Cells[row, START_COL + 8, row + 1, START_COL + 9];
            col6.Merge = true;
            col6.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            col6.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            col6.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            col6.Value = "Забележка*";

            return row;
        }
    }
}