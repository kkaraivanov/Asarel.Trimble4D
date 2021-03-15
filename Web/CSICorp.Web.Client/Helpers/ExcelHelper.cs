namespace CSICorp.Web.Client.Helpers
{
    using System.Collections.Generic;
    using System.Drawing;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;
    using Models;
    using OfficeOpenXml;
    using OfficeOpenXml.Style;

    public static class ExcelHelper
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

        private static string _currentPeriodStartDate;
        private static string _currentPeriodEndDate;
        private static string _beforePeriodStartDate = "";
        private static string _beforePeriodEndDate = "";

        public static async Task<byte[]> CreateReportAppendixTwo(SensorTable table, SensorTable tableBeforePeriod = null)
        {
            var header = table.Header.ToArray();
            var body = table.Body;
            _currentPeriodStartDate = header[0];
            _currentPeriodEndDate = header[^1];

            var stream = new MemoryStream();
            using var package = new ExcelPackage(stream);
            {
                var worksheet = package.Workbook.Worksheets.Add(APPENDIX_TITLE);
                CreateWorkSheetHeader(worksheet);
                var tableOne = CreateTable(worksheet, header, body, TABLE_ONE_TITLE, "SK", 4);
                var tableTwo = CreateTable(worksheet, header, body, TABLE_TWO_TITLE, "NC", tableOne);
                var tableNotes = AddTableNote(worksheet, TABLE_NOTE, tableTwo);

                if (tableBeforePeriod != null)
                {
                    var beforeHeader = table.Header.ToArray();
                    var beforeBody = table.Body;
                    _beforePeriodStartDate = beforeHeader[0];
                    _beforePeriodEndDate = beforeHeader[^1];
                }

                var tableLast = CreateLastTable(worksheet, tableOne, tableTwo, tableNotes);

                worksheet.DefaultColWidth = 11;
                worksheet.Column(START_COL).Width = 20;
                worksheet.Column(START_COL + 1).Width = 12;
                worksheet.Column(END_COL).Width = 12;

                return package.GetAsByteArray();
            }
        }

        private static int CreateLastTable(
            ExcelWorksheet worksheet,
            int tableOneRow,
            int tableTwoRow,
            int tableNoteRow)
        {
            var col1 = 6;
            var col2 = col1 + 1;
            var col3 = col1 + 2;
            var col4 = col1 + 3;
            var col5 = col1 + 4;
            var col6 = col1 + 5;
            var row = tableNoteRow;
            var tableOneAfterLs = worksheet.Cells[tableOneRow - 3, START_COL + 2];
            var tableOneAfterMt = worksheet.Cells[tableOneRow - 2, END_COL];
            var tableTwoAfterLs = worksheet.Cells[tableTwoRow - 3, START_COL + 2];
            var tableTwoAfterMt = worksheet.Cells[tableTwoRow - 2, END_COL];

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
            var firstDate = _beforePeriodStartDate.Substring(0, _currentPeriodStartDate.Length - 5);
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
            AddDataToLastTable(worksheet, row++, TABLE_ONE_TITLE, 0.00, 0.00, tableOneAfterLs, tableOneAfterMt);
            AddDataToLastTable(worksheet, row++, TABLE_TWO_TITLE, 0.00, 0.00, tableTwoAfterLs, tableTwoAfterMt);

            return 1;
        }

        private static void AddDataToLastTable(
            ExcelWorksheet worksheet,
            int row,
            string tableTitle,
            double beforeLs,
            double beforeMt,
            ExcelRange afterLs,
            ExcelRange afterMt)
        {
            var col1 = 6;
            var col2 = col1 + 1;
            var col3 = col1 + 2;
            var col4 = col1 + 3;
            var col5 = col1 + 4;
            var col6 = col1 + 5;
            var firstTableName = worksheet.Cells[row, START_COL, row, col1 - 1];
            firstTableName.Merge = true;
            firstTableName.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            firstTableName.Value = tableTitle;

            worksheet.Cells[row, col1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            //worksheet.Cells[row, col1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[row, col1].Value = beforeLs;
            worksheet.Cells[row, col2].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            //worksheet.Cells[row, col2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[row, col2].Value = beforeMt;
            worksheet.Cells[row, col3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            //worksheet.Cells[row, col3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[row, col3].Style.Numberformat.Format = "0.000";
            worksheet.Cells[row, col3].Formula = $"={afterLs}";
            worksheet.Cells[row, col4].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            //worksheet.Cells[row, col4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[row, col4].Style.Numberformat.Format = "0.000";
            worksheet.Cells[row, col4].Formula = $"={afterMt}";

            var col5Value = worksheet.Cells[row, col5];
            col5Value.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            col5Value.Style.Numberformat.Format = "0.000";
            //col5Value.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            col5Value.Formula = $"={worksheet.Cells[row, col3]}-{worksheet.Cells[row, col1]}";
            //var ruleCol5 = col5Value.ConditionalFormatting.AddThreeColorScale();
            //ruleCol5.LowValue.Color = Color.Red;
            //ruleCol5.HighValue.Color = Color.Green;

            var col6Value = worksheet.Cells[row, col6];
            col6Value.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            col6Value.Style.Numberformat.Format = "0.000";
            //col6Value.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            col6Value.Formula = $"={worksheet.Cells[row, col4]}-{worksheet.Cells[row, col2]}";
        }

        private static int CreateTable(
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
            appendixTitle.Merge = true;
            appendixTitle.Style.Font.Size = 14;
            appendixTitle.Style.Font.Bold = true;
            appendixTitle.Value = APPENDIX_TITLE;

            var title = worksheet.Cells[2, START_COL, 2, END_COL];
            title.Merge = true;
            title.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            title.Style.Font.Size = 14;
            title.Style.Font.Bold = true;
            title.Value = TITLE;
        }

        private static void AddTableTitle(ExcelWorksheet worksheet, string title, int row, int colStart, int colEnd)
        {
            var tableHeaderTitle = worksheet.Cells[row, colStart, row, colEnd];
            tableHeaderTitle.Merge = true;
            tableHeaderTitle.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            tableHeaderTitle.Style.Font.Size = 11;
            tableHeaderTitle.Style.Font.Bold = true;
            tableHeaderTitle.Style.Border.BorderAround(ExcelBorderStyle.Thin);
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
    }
}