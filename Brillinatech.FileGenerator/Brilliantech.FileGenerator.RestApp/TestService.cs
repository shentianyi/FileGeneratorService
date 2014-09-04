using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ServiceModel.Web;
using System.ServiceModel;
using System.ServiceModel.Activation;

namespace Brilliantech.FileGenerator.RestApp
{
    [ServiceContract]
    [AspNetCompatibilityRequirements(RequirementsMode = AspNetCompatibilityRequirementsMode.Allowed)]
    [ServiceBehavior(InstanceContextMode = InstanceContextMode.PerCall)]
    public class TestService
    {
        [WebInvoke(Method = "POST", UriTemplate = "")]
        public void Generate()
        {
          HttpRequest request=  HttpContext.Current.Request;
            HttpContext.Current.Response.Write("eeee");
           // return " ExportChart.ProcessExportRequest(HttpContext.Current);";
        }
    }
}