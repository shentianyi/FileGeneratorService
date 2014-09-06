using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.ServiceModel.Activation;
using System.ServiceModel.Web;
using System.Text;
using System.IO;
using System.Web;
using Tek4.Highcharts.Exporting;

namespace Brilliantech.FileGenerator.RestApp
{
    [ServiceContract]
    [AspNetCompatibilityRequirements(RequirementsMode = AspNetCompatibilityRequirementsMode.Allowed)]
    [ServiceBehavior(InstanceContextMode = InstanceContextMode.PerCall)]
    public class HighChartsFileService
    {
        [WebInvoke(Method = "POST", UriTemplate = "")]
        public void Generate()
        {
            ExportChart.ProcessExportRequest(HttpContext.Current);
        }
    }
}
