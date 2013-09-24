using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Newtonsoft.Json;


namespace Tek4.Highcharts.Exporting.Util
{
    public static class Json
    {
        public static T Parse<T>(string jsonString) {
          return  JsonConvert.DeserializeObject<T>(jsonString);
        }
    }
}
