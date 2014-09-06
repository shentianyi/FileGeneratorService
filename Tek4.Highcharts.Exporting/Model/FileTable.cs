using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;

namespace Tek4.Highcharts.Exporting.Model
{
  public  class FileTable
    {
      public List<string> series { get; set; }
      public List<Dictionary<string, string>> rows { get; set; }
      public Dictionary<string,string> xAxis { get; set; }
    }
}
