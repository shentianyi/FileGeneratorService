using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ServiceModel;
using System.ServiceModel.Activation;
using System.ServiceModel.Web;
using Brilliantech.Export.Report;

namespace Brilliantech.FileGenerator.RestApp
{

    [ServiceContract]
    [AspNetCompatibilityRequirements(RequirementsMode = AspNetCompatibilityRequirementsMode.Allowed)]
    [ServiceBehavior(InstanceContextMode = InstanceContextMode.PerCall)]
    public class BTReportService : ReportServiceBase
    {
        [WebInvoke(Method = "POST", UriTemplate = "ChartExcel")]
        public void GenerateChartExcel()
        {
            new ReportExporter(HttpContext.Current, ReportDefaultPath(), GetXmlString(HttpContext.Current));
        }
    }
}