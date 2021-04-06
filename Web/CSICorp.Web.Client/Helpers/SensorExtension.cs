namespace CSICorp.Web.Client.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Models;

    public static class SensorExtension
    {
        private const double MINIMUM_VALUE = 150.9;
        private const double SECONDS = 1800;
        private const string CRITERIA = "Kolichestvo";
        private const string DOUBLE_FORMAT = "0.000";

        public static async Task<SensorTable> GetDrillWheelsDataAsync(this List<ZipEntry> entries, bool isWater = false)
        {
            var result = ImputDataToTable(entries, isWater);

            if(!isWater)
            {
                result.IsNotWater();

            }
            else
            {
                result.IsWater();
            }
            
            return result;
        }

        private static SensorTable ImputDataToTable(List<ZipEntry> entries, bool isWater = false)
        {
            var result = new SensorTable();
            var counter = 0;

            foreach (var item in entries.Where(x => !x.Name.EndsWith("/")))
            {
                if (counter > 7)
                {
                    break;
                }

                var getDateFromItemName = item.Name.Split('_')[1].Split('.')[0];
                var dateSplit = getDateFromItemName.Split('-');
                var date = $"{dateSplit[2]}.{dateSplit[1]}.{dateSplit[0]}";
                result.Header.Add(date);
                var dataSensorsList = new List<Sensor>();
                var dataSensors = CollectionData(item.Content, isWater);
                dataSensorsList = DrillWheelDataList(dataSensors, dataSensorsList, isWater).ToList();

                foreach (var sensor in dataSensorsList)
                {
                    if (!result.Body.ContainsKey(sensor.SensorName))
                    {
                        result.Body[sensor.SensorName] = new List<string>();
                        if (counter > 0)
                            for (int i = 0; i < counter; i++)
                                result.Body[sensor.SensorName].Add("0");
                    }

                    var debit = sensor.Debit != 0 ? sensor.Debit.ToString(DOUBLE_FORMAT) : "0";
                    result.Body[sensor.SensorName].Add(debit);
                }

                counter++;
            }

            return result;
        }

        private static List<Sensor> DrillWheelDataList(IEnumerable<Sensor> dataSensors, List<Sensor> sensors, bool isWater = false)
        {
            var distinctSensorNames = dataSensors.Select(x => x.SensorName).Distinct().ToList();

            foreach (var sensorName in distinctSensorNames)
            {
                var currentSensor = sensors.FirstOrDefault(x => x.SensorName == sensorName);
                var dataSensorDebit = 0.00;
                var counter = 0;

                foreach (var dataSensor in dataSensors.Where(x => x.SensorName == sensorName))
                {
                    if (!isWater)
                    {
                        if (currentSensor != null)
                        {
                            dataSensorDebit += dataSensor.Debit;
                        }
                        else
                        {
                            dataSensorDebit += dataSensor.Debit;
                            dataSensor.Debit = 0;
                            currentSensor = dataSensor;
                            sensors.Add(currentSensor);
                        }

                        counter++;
                    }
                    else
                    {
                        if (currentSensor != null)
                        {
                            dataSensorDebit = dataSensor.Debit;
                        }
                        else
                        {
                            if (dataSensorDebit < dataSensor.Debit)
                            {
                                dataSensorDebit = dataSensor.Debit;
                            }

                            dataSensor.Debit = 0;
                            currentSensor = dataSensor;
                            sensors.Add(currentSensor);
                        }
                    }
                }

                var debitToSeconds = counter != 0 ? dataSensorDebit / SECONDS : dataSensorDebit;
                var diffForDay = counter != 0 ? Math.Abs(debitToSeconds / counter) : debitToSeconds;
                var currentSensorDebit = currentSensor.Debit;
                var debit = counter != 0 ? currentSensorDebit + diffForDay : dataSensorDebit;
                var dateStamp = currentSensor.DateStamp.Split(' ')[0].Replace('-', '.');

                currentSensor.DateStamp = dateStamp;
                currentSensor.Debit = debit;
            }

            return sensors.OrderBy(x => x.SensorName).ToList();
        }

        private static List<Sensor> CollectionData(string content, bool isWater = false)
        {
            var sensors = new List<Sensor>();
            var temp = content.Split('\n').ToList();
            temp.RemoveAt(0);
            temp = new List<string>(temp.Where(x => x.EndsWith('\r')));

            if (!string.IsNullOrEmpty(content))
            {
                var lines = content.Split('\n');
                var sensor = new Sensor();
                if (!isWater)
                {
                    foreach (var item in temp.Where(x => x.Contains(CRITERIA)))
                    {
                        if (item == null)
                            break;

                        var currentLine = item.Replace("\r", "").Split(',', 5);

                        if (currentLine.Length == 0)
                            break;

                        var tempDebit = !double.TryParse(currentLine[2], out double debit) ? 0 : debit;
                        var debitValue = 0.00;
                        var sensorName = GetName(currentLine[1].TrimStart(), isWater);
                        debitValue = tempDebit >= MINIMUM_VALUE ? tempDebit : 0;

                        if (!sensorName.Contains('-'))
                        {
                            sensorName = new string(sensorName.Insert(2, "-"));
                        }

                        sensor = new Sensor
                        {
                            DateStamp = currentLine[0],
                            SensorName = sensorName,
                            Debit = debitValue,
                            TotalLitersForDay = currentLine[3].TrimStart(),
                            TotalLitersFromStart = currentLine[4].TrimStart()
                        };

                        if (!string.IsNullOrEmpty(sensor.SensorName))
                            sensors.Add(sensor);
                    }
                }
                else
                {
                    foreach (var item in temp.Where(x => !x.Contains(CRITERIA)))
                    {
                        if (item == null)
                            break;

                        var currentLine = item.Replace("\r", "").Split(',', 5);

                        if (currentLine.Length == 0)
                            break;

                        var debitValue = !double.TryParse(currentLine[2], out double debit) ? 0 : debit;
                        var sensorName = GetName(currentLine[1].TrimStart(), isWater);

                        if (!sensorName.Contains('-'))
                        {
                            sensorName = new string(sensorName.Insert(2, "-"));
                        }

                        sensor = new Sensor
                        {
                            DateStamp = currentLine[0],
                            SensorName = sensorName,
                            Debit = debitValue,
                            TotalLitersForDay = currentLine[3].TrimStart(),
                            TotalLitersFromStart = currentLine[4].TrimStart()
                        };

                        if (!string.IsNullOrEmpty(sensor.SensorName))
                            sensors.Add(sensor);
                    }
                }
            }

            return sensors;
        }

        private static string GetName(string name, bool isWater)
        {
            if (!isWater)
            {
                var tempName = new StringBuilder();

                if (!name.Contains(' '))
                    return name;

                for (int c = 0; c < name.Length; c++)
                {
                    var currentChar = name[c];

                    if (currentChar != ' ')
                    {
                        tempName.Append(currentChar);
                    }
                    else
                    {
                        var index = tempName.ToString().Trim().Length - 1;
                        var before = tempName.ToString().Trim()[index];

                        if (char.IsDigit(before))
                        {
                            return tempName.ToString().Trim();
                        }

                        tempName.Append(currentChar);
                    }
                }

                return tempName.ToString().Trim();
            }
            else
            {
                var countWhiteSpace = name.Count(x => x == ' ');

                if (countWhiteSpace >= 2)
                {
                    var counter = 2;
                    var builder = new StringBuilder();
                    for (int i = 0; i < name.Length; i++)
                    {
                        if (name[i] == ' ')
                            counter--;

                        if (counter == 0)
                            break;

                        builder.Append(name[i]);
                    }

                    name = new string(builder.ToString().Trim().Replace(" ", ""));
                }
                else if (countWhiteSpace == 1)
                {
                    name = new string(name.Split(' ')[0]);
                }

                return name;
            }
        }

        private static void IsWater(this SensorTable table)
        {
            foreach (var (key, value) in table.Body)
            {
                var count = value.Count;
                if (count != table.Header.Count)
                {
                    for (int i = count; i < table.Header.Count; i++)
                        value.Add("0");
                }

                var dailyValues = new List<double>();
                value.ForEach(x => { dailyValues.Add(double.Parse(x)); });

                var max = dailyValues.Max();
                value.Insert(0, max.ToString(DOUBLE_FORMAT));
            }
        }

        private static void IsNotWater(this SensorTable table)
        {
            foreach (var (key, value) in table.Body)
            {
                var count = value.Count;
                if (count != table.Header.Count)
                {
                    for (int i = count; i < table.Header.Count; i++)
                        value.Add("0");
                }

                var dailyValues = new List<double>();
                value.ForEach(x => { dailyValues.Add(double.Parse(x)); });

                var average = 0.00;
                if (dailyValues.Any(x => x != 0))
                    average = dailyValues.Where(x => x != 0).Average();

                value.Insert(0, average.ToString(DOUBLE_FORMAT));
            }
        }
    }
}
