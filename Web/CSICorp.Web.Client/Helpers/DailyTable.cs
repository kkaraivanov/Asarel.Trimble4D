namespace CSICorp.Web.Client.Helpers
{
    using System.Collections.Generic;
    using System.Drawing;
    using System.Linq;
    using System.Threading.Tasks;
    using Models;
    using OfficeOpenXml;
    using OfficeOpenXml.Style;

    public static class DailyTable
    {
        private const string APPENDIX_TITLE = "Приложение 2";
        private const string TITLE = "Данни за валежите, дебитите и водните нива на сондажите и водопроявленията";
        private const string TABLE_ONE_TITLE = "1. Данни за сондажните кладенци";
        private const string TABLE_TWO_TITLE = "2. Данни за заустените хоризонтални сондажи";
        private const string TABLE_Last_TITLE = "Таблица 1 Общо дренирано и изпомпано количество за седмицата";
        private const string TABLE_NOTE = "Забележка: Хоризонталните дренажни сондажи отбелязани със звездичка * се измерват ръчно веднъж седмично, поради демонтирани дебитометри.  Хоризонталните дренажни сондажи отбелязани със две звездички ** са с деинсталирано оборудване.";
        private const string LEFT_HEADER = "Дебит сондажни кладенци по дни и средно за седмицата [l/s]";
        private const string RIGHT_HEADER = "Общо изпомпано количество вода за седмицата [m3]";
        private const string QUALITY = "Общо количество";
        private const string QUALITY_FOR_WEEK = "Общ седм, дебит";
        private const string LITERS_TO_SECONDS = "[l/s]";
        private const string DECIMETERS = "[m3/d]";
        private const string METERS = "[m3]";
        private const int START_COL = 1;
        private const int END_COL = 11;
        private const double CONVERT_CONST = 86.4;

        private static string _currentPeriodStartDate;
        private static string _currentPeriodEndDate;
        private static string _beforePeriodStartDate;
        private static string _beforePeriodEndDate;
        private static int _lastTableRow;
        public static async Task CreateTable(
            ExcelWorksheet worksheet, 
            SensorTable currentPeriodeDebit, 
            SensorTable beforePeriodeDebit)
        {
            await CreateTable(worksheet, currentPeriodeDebit);

            var beforePeriodeHeader = beforePeriodeDebit.Header.ToArray();
            var currentPeriodBody = currentPeriodeDebit.Body;
            var beforePeriodeBody = beforePeriodeDebit.Body;

            _beforePeriodStartDate = beforePeriodeHeader[0];
            _beforePeriodEndDate = beforePeriodeHeader[^1];

            var currentPeriodeTableOneSum = GetSum(currentPeriodBody, "SK");
            var currentPeriodeTableTwoSum = GetSum(currentPeriodBody, "NC");
            var beforePeriodeTableOneSum = GetSum(beforePeriodeBody, "SK");
            var beforePeriodeTableTwoSum = GetSum(beforePeriodeBody, "NC");

            var tableLast = CreateLastTable(
                worksheet,
                currentPeriodeTableOneSum,
                currentPeriodeTableTwoSum,
                beforePeriodeTableOneSum,
                beforePeriodeTableTwoSum);
        }

        public static async Task CreateTable(ExcelWorksheet worksheet, SensorTable currentPeriodeDebit)
        {
            var currentPeriodHeader = currentPeriodeDebit.Header.ToArray();
            var currentPeriodBody = currentPeriodeDebit.Body;

            _currentPeriodStartDate = currentPeriodHeader[0];
            _currentPeriodEndDate = currentPeriodHeader[^1];

            CreateWorkSheetHeader(worksheet);
            var tableOne = CreateWorksheetTable(worksheet, currentPeriodHeader, currentPeriodBody, TABLE_ONE_TITLE, "SK", 4);
            var tableTwo = CreateWorksheetTable(worksheet, currentPeriodHeader, currentPeriodBody, TABLE_TWO_TITLE, "NC", tableOne);
            _lastTableRow = AddTableNote(worksheet, TABLE_NOTE, tableTwo);

            FormatWorksheet(worksheet);
        }

        private static int CreateWorksheetTable(
            ExcelWorksheet worksheet,
            string[] header,
            Dictionary<string, List<string>> body,
            string title,
            string criteriaForSensor,
            int startRow)
        {
            var row = startRow;
            AddTableTitle(worksheet, title, row++, START_COL, END_COL);
            AddTableHeader(worksheet, header, row++);

            var getBody = AddTableBody(worksheet, body, row, criteriaForSensor);
            var rowCount = getBody[0];
            var colCount = getBody[1];

            AddBodyLeftCol(worksheet, row, rowCount);
            AddBodyRightCol(worksheet, row, rowCount);

            AddTableFooter(worksheet, row, rowCount, colCount);
            worksheet.Cells[row, 3, rowCount - 1, 3].Style.Font.Bold = true;

            return rowCount + 3;
        }

        private static void CreateWorkSheetHeader(ExcelWorksheet worksheet)
        {
            var appendixTitle = worksheet.Cells[1, END_COL - 1, 1, END_COL];
            appendixTitle.HeaderTitle();
            appendixTitle.Value = APPENDIX_TITLE;

            var title = worksheet.Cells[2, START_COL, 2, END_COL];
            title.HeaderTitle();
            title.Centered();
            title.Value = TITLE;
        }

        private static void AddTableTitle(ExcelWorksheet worksheet, string title, int row, int colStart, int colEnd)
        {
            var tableHeaderTitle = worksheet.Cells[row, colStart, row, colEnd];
            tableHeaderTitle.TableTitle();
            tableHeaderTitle.Value = title;
        }

        private static void AddTableHeader(ExcelWorksheet worksheet, string[] header, int row)
        {
            worksheet.Cells[row, 1, row, header.Length + 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells[row, 1, row, header.Length + 3].Style.Font.Bold = true; ;
            worksheet.Cells[row, 2].Value = "Сензор/Дата";
            worksheet.Cells[row, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells[row, 3].Value = "Сред. сед.";
            worksheet.Cells[row, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            for (int i = 0; i < header.Length; i++)
            {
                worksheet.Cells[row, (4 + i)].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells[row, (4 + i)].Value = header[i];
            }
        }

        private static int[] AddTableBody(ExcelWorksheet worksheet, Dictionary<string, List<string>> body, int row, string criteriaForSensor)
        {
            var rowCount = row;
            var colCount = 0;

            foreach (var (key, values) in body.Where(x => x.Key.Contains(criteriaForSensor)))
            {
                var counterValue = 3;
                colCount = values.Count;
                foreach (var value in values)
                {
                    var sensorValue = worksheet.Cells[rowCount, counterValue++];

                    if (value == "0")
                    {
                        sensorValue.Value = "-";
                        sensorValue.Style.Font.Color.SetColor(Color.Red);
                    }
                    else
                    {
                        sensorValue.Value = decimal.Parse(value);
                        //sensorValue.Style.Font.Color.SetColor(Color.Green);
                        sensorValue.Style.Numberformat.Format = "0.000";
                    }

                    sensorValue.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                }

                var sensorData = worksheet.Cells[rowCount++, 2];
                sensorData.Value = key;
                sensorData.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            }

            return new int[] { rowCount, colCount };
        }
        private static void AddBodyRightCol(ExcelWorksheet worksheet, int row, in int rowCount)
        {
            var rightHeader = worksheet.Cells[row - 1, END_COL, rowCount, END_COL];
            rightHeader.Merge = true;
            rightHeader.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            rightHeader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            rightHeader.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            rightHeader.Style.WrapText = true;
            rightHeader.Style.Font.Size = 9;
            rightHeader.Value = RIGHT_HEADER;
        }

        private static void AddBodyLeftCol(ExcelWorksheet worksheet, int row, int rowCount)
        {
            var leftHeader = worksheet.Cells[row, START_COL, rowCount - 1, START_COL];
            leftHeader.Merge = true;
            leftHeader.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            leftHeader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            leftHeader.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            leftHeader.Style.WrapText = true;
            leftHeader.Style.Font.Size = 9;
            leftHeader.Value = LEFT_HEADER;
        }

        private static void AddTableFooter(ExcelWorksheet worksheet, int row, int rowCount, int colCount)
        {
            var totalDebitTitle = worksheet.Cells[rowCount, START_COL];
            totalDebitTitle.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            totalDebitTitle.Value = QUALITY_FOR_WEEK;

            var totalDebitParce = worksheet.Cells[rowCount, START_COL + 1];
            totalDebitParce.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            totalDebitParce.Value = LITERS_TO_SECONDS;

            var totalDebitAverage = worksheet.Cells[rowCount, START_COL + 2];
            totalDebitAverage.Formula = $"AVERAGE(D{rowCount}:J{rowCount})";
            totalDebitAverage.Style.Numberformat.Format = "0.000";
            totalDebitAverage.Style.Font.Bold = true;
            totalDebitAverage.Style.Border.BorderAround(ExcelBorderStyle.Thin);

            var totalDebitSum = worksheet.Cells[rowCount, 4, rowCount, colCount + 2];
            totalDebitSum.Formula = $"SUM(D{row}:D{rowCount - 1})";
            totalDebitSum.Style.Font.Bold = true;
            totalDebitSum.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            totalDebitSum.Style.Border.Left.Style = ExcelBorderStyle.Thin;

            var totalQuantity = worksheet.Cells[rowCount + 1, START_COL];
            totalQuantity.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            totalQuantity.Style.Font.Bold = true;
            totalQuantity.Value = QUALITY;

            var totalQuantityParce = worksheet.Cells[rowCount + 1, START_COL + 1];
            totalQuantityParce.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            totalQuantityParce.Style.Font.Bold = true;
            totalQuantityParce.Value = DECIMETERS;

            var totalQuantityAverage = worksheet.Cells[rowCount + 1, START_COL + 2];
            totalQuantityAverage.Formula = $"AVERAGE(D{rowCount + 1}:J{rowCount + 1})";
            totalQuantityAverage.Style.Numberformat.Format = "0.000";
            totalQuantityAverage.Style.Font.Bold = true;
            totalQuantityAverage.Style.Border.BorderAround(ExcelBorderStyle.Thin);

            var totalQuantityTotalSum = worksheet.Cells[rowCount + 1, colCount + 3];
            totalQuantityTotalSum.Formula = $"SUM(D{rowCount + 1}:J{rowCount + 1})";
            totalQuantityTotalSum.Style.Numberformat.Format = "0.000";
            totalQuantityTotalSum.Style.Font.Bold = true;
            totalQuantityTotalSum.Style.Border.BorderAround(ExcelBorderStyle.Thin);

            var totalQuantitySum = worksheet.Cells[rowCount + 1, 4, rowCount + 1, colCount + 2];
            totalQuantitySum.Formula = $"D{rowCount}*86.4";
            totalQuantitySum.Style.Numberformat.Format = "0.000";
            totalQuantitySum.Style.Font.Bold = true;
            totalQuantitySum.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            totalQuantitySum.Style.Border.Left.Style = ExcelBorderStyle.Thin;
        }

        private static int AddTableNote(ExcelWorksheet worksheet, string tableNote, int row)
        {
            var tableHeaderTitle = worksheet.Cells[row - 1, START_COL, row, END_COL];
            tableHeaderTitle.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            tableHeaderTitle.Style.Font.Size = 11;
            tableHeaderTitle.Merge = true;
            tableHeaderTitle.Style.Font.Italic = true;
            tableHeaderTitle.Style.WrapText = true;
            tableHeaderTitle.Value = tableNote;

            return row + 3;
        }

        private static int CreateLastTable(
            ExcelWorksheet worksheet,
            double currentPeriodeTableOneValue,
            double currentPeriodeTableTwoValue,
            double beforePeriodeTableOneValue,
            double beforePeriodeTableTwoValue)
        {
            var col1 = 6;
            var col2 = col1 + 1;
            var col3 = col1 + 2;
            var col4 = col1 + 3;
            var col5 = col1 + 4;
            var col6 = col1 + 5;
            var row = _lastTableRow;

            var title = worksheet.Cells[row, START_COL, row, END_COL];
            title.Merge = true;
            title.Style.Font.Italic = true;
            title.Value = TABLE_Last_TITLE;
            row++;

            #region Header first row

            var headerBeforePeriod = worksheet.Cells[row, col1, row, col1 + 1];
            headerBeforePeriod.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            headerBeforePeriod.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            headerBeforePeriod.Merge = true;
            var firstDate = !string.IsNullOrEmpty(_beforePeriodStartDate)
                ? _beforePeriodStartDate.Substring(0, _beforePeriodStartDate.Length - 5)
                : " ";
            headerBeforePeriod.Value = $"{firstDate}-{_beforePeriodEndDate}";

            var headerCurrentPeriod = worksheet.Cells[row, col1 + 2, row, col1 + 3];
            headerCurrentPeriod.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            headerCurrentPeriod.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            headerCurrentPeriod.Merge = true;
            firstDate = _currentPeriodStartDate.Substring(0, _currentPeriodStartDate.Length - 5);
            headerCurrentPeriod.Value = $"{firstDate}-{_currentPeriodEndDate}";

            var headerDiffPeriod = worksheet.Cells[row, col1 + 4, row, col1 + 5];
            headerDiffPeriod.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            headerDiffPeriod.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            headerDiffPeriod.Merge = true;
            headerDiffPeriod.Value = "Разлика";

            #endregion

            row++;

            #region Header secont row

            var headerCol1 = worksheet.Cells[row, col1];
            headerCol1.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            headerCol1.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            headerCol1.Value = LITERS_TO_SECONDS;

            var headerCol2 = worksheet.Cells[row, col2];
            headerCol2.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            headerCol2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            headerCol2.Value = METERS;

            var headerCol3 = worksheet.Cells[row, col3];
            headerCol3.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            headerCol3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            headerCol3.Value = LITERS_TO_SECONDS;

            var headerCol4 = worksheet.Cells[row, col4];
            headerCol4.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            headerCol4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            headerCol4.Value = METERS;

            var headerCol5 = worksheet.Cells[row, col5];
            headerCol5.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            headerCol5.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            headerCol5.Value = LITERS_TO_SECONDS;

            var headerCol6 = worksheet.Cells[row, col6];
            headerCol6.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            headerCol6.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            headerCol6.Value = METERS;

            #endregion

            row++;
            AddDataToLastTable(worksheet, row++, TABLE_ONE_TITLE, currentPeriodeTableOneValue, beforePeriodeTableOneValue);
            AddDataToLastTable(worksheet, row++, TABLE_TWO_TITLE, currentPeriodeTableTwoValue, beforePeriodeTableTwoValue);

            return row;
        }

        private static void AddDataToLastTable(
            ExcelWorksheet worksheet,
            int row,
            string tableTitle,
            double currentPeriodeValue,
            double beforePeriodeValue)
        {
            var col1 = 6;
            var col2 = col1 + 1;
            var col3 = col1 + 2;
            var col4 = col1 + 3;
            var col5 = col1 + 4;
            var col6 = col1 + 5;
            var currentPeriodeConvertedValue = (currentPeriodeValue * 7) * CONVERT_CONST;
            var beforePeriodeConvertedValue = (beforePeriodeValue * 7) * CONVERT_CONST;
            var firstTableName = worksheet.Cells[row, START_COL, row, col1 - 1];
            firstTableName.Merge = true;
            firstTableName.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            firstTableName.Value = tableTitle;

            worksheet.Cells[row, col1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            //worksheet.Cells[row, col1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[row, col1].Value = beforePeriodeValue;
            worksheet.Cells[row, col2].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            //worksheet.Cells[row, col2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[row, col2].Value = beforePeriodeConvertedValue;
            worksheet.Cells[row, col3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            //worksheet.Cells[row, col3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[row, col3].Style.Numberformat.Format = "0.000";
            worksheet.Cells[row, col3].Value = currentPeriodeValue;
            worksheet.Cells[row, col4].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            //worksheet.Cells[row, col4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[row, col4].Style.Numberformat.Format = "0.000";
            worksheet.Cells[row, col4].Value = currentPeriodeConvertedValue;

            var col5Value = worksheet.Cells[row, col5];
            col5Value.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            col5Value.Style.Numberformat.Format = "0.000";
            //col5Value.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            var value = currentPeriodeValue - beforePeriodeValue;
            col5Value.Value = value;
            SetCcllColor(col5Value, value);

            var col6Value = worksheet.Cells[row, col6];
            col6Value.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            col6Value.Style.Numberformat.Format = "0.000";
            //col6Value.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            value = currentPeriodeConvertedValue - beforePeriodeConvertedValue;
            col6Value.Value = value;
            SetCcllColor(col6Value, value);
        }

        private static double GetSum(Dictionary<string, List<string>> periodBody, string criteria)
        {
            double result = 0;
            foreach (var (key, values) in periodBody.Where(x => x.Key.Contains(criteria)))
            {
                result += double.Parse(values[0]);
            }

            return result;
        }

        private static void SetCcllColor(ExcelRange cellExcelRange, in double value)
        {
            if (value < 0)
            {
                cellExcelRange.Style.Font.Color.SetColor(Color.Red);
            }
            else
            {
                cellExcelRange.Style.Font.Color.SetColor(Color.Green);
            }
        }

        private static void FormatWorksheet(ExcelWorksheet worksheet)
        {
            worksheet.DefaultColWidth = 11;
            worksheet.Column(START_COL).Width = 20;
            worksheet.Column(START_COL + 1).Width = 12;
            worksheet.Column(END_COL).Width = 12;
        }
    }
}