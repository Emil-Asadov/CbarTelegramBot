using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace ConsoleAppCbarTelegramBot.Models
{
    public class JsonFields
    {
        [JsonPropertyName("date")]
        public string Date { get; set; }

        [JsonPropertyName("valuteCode")]
        public string ValuteCode { get; set; }

        [JsonPropertyName("nominal")]
        public string Nominal { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("value")]
        public string Value { get; set; }
    }
}
