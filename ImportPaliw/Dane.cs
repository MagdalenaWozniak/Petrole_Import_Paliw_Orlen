//using Microsoft.Xrm.Sdk.Messages;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportPaliw
{
    public partial class Dane
    {
        public List<Paliwo> Paliwa { get; set; }
    }

    public partial class Paliwo
    {

        [JsonProperty("productName")]
        public string Nazwa { get; set; }

        [JsonProperty("effectiveDate")]
        public DateTime Data { get; set; }

        [JsonProperty("publishFrom")]
        public DateTime DataPublikacji { get; set; }

        [JsonProperty("value")]
        public Double Wartosc { get; set; }

        [JsonProperty("locationName")]
        public string Lokalizacjia { get; set; }
    }



    public partial class Dane
    {
        public static List<Paliwo> FromJson(string json) => JsonConvert.DeserializeObject<List<Paliwo>>(json); //JsonConvert.DeserializeObject<Dane>(json);
    }
}
