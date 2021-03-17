namespace CSICorp.Web.Client.Helpers
{
    using System.Collections.Generic;
    using System.Drawing;
    using System.Linq;
    using System.Threading.Tasks;
    using Models;
    using OfficeOpenXml;
    using OfficeOpenXml.Style;

    public static class WeeklyTable
    {
        private const string TABLE_ONE_TITLE = "4. Повърхностен отток на реките [l/s]";
        private const string TABLE_TWO_TITLE = "5. Водни нива на сондажните кладенци [m]";
        private const string TABLE_TWO_NOTE = "* СВН - статично водно ниво, водно ниво при неработеща помпа; ДВН - динамично водно ниво, водно ниво при работеща помпа.";
        private const string TABLE_THREE_TITLE = "8. Промяна на водните нива в наблюдателните сондажи";

        private const int START_COL = 1;
        private const int END_COL = 9;

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
            var tableOne = await CreateTableOne(worksheet, TABLE_ONE_TITLE, row);
            var tableTwo = await CreateTableTwo(
                worksheet,
                TABLE_TWO_TITLE,
                tableOne,
                currentPeriodeWaterLevel,
                beforePeriodeWaterLevel,
                waterSensorDataList.Where(x => x.Name.Contains("SK")).ToList());
            var tableThree = await CreateTableThree(
                worksheet,
                TABLE_THREE_TITLE,
                tableTwo,
                currentPeriodeWaterLevel,
                beforePeriodeWaterLevel,
                waterSensorDataList.Where(x => !x.Name.Contains("SK")).ToList());

            worksheet.DefaultColWidth = 12;
            worksheet.Column(START_COL).Width = 16;
        }

        private static async Task<int> CreateTableOne(ExcelWorksheet worksheet, string title, int row)
        {
            var tableTitle = worksheet.Cells[row, START_COL, row, END_COL];
            tableTitle.TableTitle();
            tableTitle.Value = title;

            return row + 2;
        }

        private static async Task<int> CreateTableTwo(
            ExcelWorksheet worksheet,
            string title,
            int row,
            SensorTable currentPeriodeWaterLevel,
            SensorTable beforePeriodeWaterLevel,
            List<WaterLevelSensors> waterSensorDataList)
        {
            #region Create table header

            var tableTitle = worksheet.Cells[row, START_COL, row, END_COL];
            tableTitle.TableTitle();
            tableTitle.Value = title;

            #endregion

            row++;

            #region Create row one

            row = await CreateTableHeader(worksheet, row);

            #endregion

            row++;

            #region Create row two

            var col4 = worksheet.Cells[row, START_COL + 3];
            col4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            col4.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            col4.Value = "Кота [m]";

            worksheet.SetCellValue(row, START_COL + 3, "Кота [m]");
            worksheet.SetCellValue(row, START_COL + 4, "WL [m]");
            worksheet.SetCellValue(row, START_COL + 5, "Кота [m]");
            worksheet.SetCellValue(row, START_COL + 6, "WL [m]");
            worksheet.SetCellValue(row, START_COL + 7, "Разлика");

            #endregion

            row++;

            #region Create table body

            var body = await AddTableTwoBody(worksheet, row, currentPeriodeWaterLevel, beforePeriodeWaterLevel, waterSensorDataList);

            #endregion

            var note = worksheet.Cells[body, START_COL, body++, END_COL];
            note.Merge = true;
            note.Style.Font.Italic = true;
            note.Value = TABLE_TWO_NOTE;

            return body + 2;
        }

        private static async Task<int> CreateTableThree(
            ExcelWorksheet worksheet,
            string title,
            int row,
            SensorTable currentPeriodeWaterLevel,
            SensorTable beforePeriodeWaterLevel,
            List<WaterLevelSensors> waterSensorDataList)
        {
            #region Create table header

            var tableTitle = worksheet.Cells[row, START_COL, row, END_COL];
            tableTitle.TableTitle();
            tableTitle.Value = title;

            #endregion

            row++;

            #region Create row one

            row = await CreateTableHeader(worksheet, row);

            #endregion

            row++;

            #region Create row two

            worksheet.SetCellValue(row, START_COL + 3, "Кота [m]");
            worksheet.SetCellValue(row, START_COL + 4, "WL [m]");
            worksheet.SetCellValue(row, START_COL + 5, "Кота [m]");
            worksheet.SetCellValue(row, START_COL + 6, "WL [m]");
            worksheet.SetCellValue(row, START_COL + 7, "Разлика");

            #endregion

            row++;

            #region Create table body

            var body = await AddTableTwoBody(worksheet, row, currentPeriodeWaterLevel, beforePeriodeWaterLevel, waterSensorDataList);

            #endregion

            var note = worksheet.Cells[body, START_COL, body, END_COL];
            note.Merge = true;
            note.Style.Font.Italic = true;
            note.Value = TABLE_TWO_NOTE;

            return row + 2;
        }

        private static async Task<int> AddTableTwoBody(
            ExcelWorksheet worksheet,
            int row,
            SensorTable currentPeriodeWaterLevel,
            SensorTable beforePeriodeWaterLevel,
            List<WaterLevelSensors> waterSensorDataList)
        {
            var currentRow = row;
            foreach (var sensorData in waterSensorDataList)
            {
                worksheet.SetCellValue(currentRow, START_COL, sensorData.Name);
                worksheet.SetCellValue(currentRow, START_COL + 1, sensorData.Elevation);
                worksheet.SetCellValue(currentRow, START_COL + 2, sensorData.Depth);

                double beforeMax = GetMaxData(sensorData.Name, beforePeriodeWaterLevel.Body);
                double currentMax = GetMaxData(sensorData.Name, currentPeriodeWaterLevel.Body);
                double beforeElevation = sensorData.Elevation - beforeMax;
                double currentElevation = sensorData.Elevation - currentMax;
                double diff = beforeMax - currentMax;

                worksheet.SetCellValue(currentRow, START_COL + 3, beforeMax != 0 ? beforeElevation : beforeMax);
                worksheet.SetCellValue(currentRow, START_COL + 4, beforeMax * -1);
                worksheet.Colored(currentRow, START_COL + 4, Color.Blue);
                worksheet.SetCellValue(currentRow, START_COL + 5, currentMax != 0 ? currentElevation : currentMax);
                worksheet.SetCellValue(currentRow, START_COL + 6, currentMax * -1);
                worksheet.Colored(currentRow, START_COL + 6, Color.Blue);
                
                if (diff > 0)
                {
                    worksheet.Colored(currentRow, START_COL + 7, Color.Red);
                    worksheet.SetCellValue(currentRow, START_COL + 7, diff);
                }
                else if (diff == 0)
                {
                    worksheet.Colored(currentRow, START_COL + 7, Color.Red);
                    worksheet.SetCellValue(currentRow, START_COL + 7, "-");
                }
                else
                {
                    worksheet.Colored(currentRow, START_COL + 7, Color.Green);
                    worksheet.SetCellValue(currentRow, START_COL + 7, diff);
                }

                if (beforeMax != 0 && currentMax != 0)
                {
                    if (diff == 0)
                    {
                        worksheet.Colored(currentRow, START_COL + 7, Color.Red);
                        worksheet.SetCellValue(currentRow, START_COL + 7, "няма ВН");
                    }
                }

                if (diff == 0)
                {
                    worksheet.SetCellValue(currentRow, START_COL + 8, "-");
                }
                else
                {
                    worksheet.SetCellValue(currentRow, START_COL + 8, "ДВН");
                }

                currentRow++;
            }

            return currentRow++;
        }

        private static double GetMaxData(string sensorDataName, Dictionary<string, List<string>> periode)
        {
            foreach (var (key, value) in periode)
            {
                if (key != sensorDataName)
                    continue;

                var tempList = new List<double>();
                value.ForEach(x =>
                {
                    var temp = double.Parse(x);
                    tempList.Add(temp);
                });

                return tempList.Max();
            }

            return 0;
        }

        private static async Task<int> CreateTableHeader(ExcelWorksheet worksheet, int row)
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

            var col6 = worksheet.Cells[row, START_COL + 8, row + 1, START_COL + 8];
            col6.Merge = true;
            col6.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            col6.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            col6.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            col6.Value = "Забележка*";

            return row;
        }
    }
}