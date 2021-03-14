namespace CSICorp.Web.Client.Models
{
    using System.Collections.Generic;

    public class SensorTable
    {
        public List<string> Header { get; set; }
        public Dictionary<string, List<string>> Body { get; set; }

        public SensorTable()
        {
            Header = new List<string>();
            Body = new Dictionary<string, List<string>>();
        }
    }
}