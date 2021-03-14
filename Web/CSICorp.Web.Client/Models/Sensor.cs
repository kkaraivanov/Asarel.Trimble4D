namespace CSICorp.Web.Client.Models
{
    using System.Text.Json.Serialization;

    public class Sensor
    {
        public string DateStamp { get; set; }
        public string SensorName { get; set; }
        public double Debit { get; set; }
        public string TotalLitersForDay { get; set; }
        public string TotalLitersFromStart { get; set; }
    }
}