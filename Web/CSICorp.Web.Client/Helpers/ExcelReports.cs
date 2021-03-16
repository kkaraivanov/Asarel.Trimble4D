namespace CSICorp.Web.Client.Helpers
{
    using System.Collections.Generic;
    using System.IO;
    using System.Threading.Tasks;
    using Models;
    using OfficeOpenXml;

    public static class ExcelReports
    {
        private const string SHEET_DAILY = "daily";
        private const string SHEET_WEEKLY = "weekly";

        public static async Task<byte[]> CreateReport(
            string fileName,
            SensorTable currentPeriodeDebit)
        {
            var stream = new MemoryStream();
            using var package = new ExcelPackage(stream);
            {
                var sheetName = $"{fileName[0]} {SHEET_DAILY}";
                var worksheetDaily = package.Workbook.Worksheets.Add(sheetName);

                await DailyTable.CreateTable(worksheetDaily, currentPeriodeDebit);

                return package.GetAsByteArray();
            }
        }

        public static async Task<byte[]> CreateReport(
            string fileName,
            SensorTable currentPeriodeDebit, 
            SensorTable currentPeriodeWaterLevel,
            SensorTable beforePeriodeDebit,
            SensorTable beforePeriodeWaterLevel,
            List<WaterLevelSensors> waterSensorDataList)
        {
            var stream = new MemoryStream();
            using var package = new ExcelPackage(stream);
            {
                var sheetName = $"{fileName[0]} {SHEET_DAILY}";
                var worksheetDaily = package.Workbook.Worksheets.Add(sheetName);
                sheetName = $"{fileName[0]} {SHEET_WEEKLY}";
                var worksheetWeekly = package.Workbook.Worksheets.Add(sheetName);

                await DailyTable.CreateTable(worksheetDaily, currentPeriodeDebit, beforePeriodeDebit);
                await WeeklyTable.CreateTable(worksheetWeekly, fileName, currentPeriodeWaterLevel, beforePeriodeWaterLevel, waterSensorDataList);

                return package.GetAsByteArray();
            }
        }
    }
}