using System;
using System.ServiceModel.Activation;
using System.Web;
using System.Web.Routing;

namespace Brilliantech.FileGenerator.RestApp
{
    public class Global : HttpApplication
    {
        void Application_Start(object sender, EventArgs e)
        {
            RegisterRoutes();
        }

        private void RegisterRoutes()
        {
            RouteTable.Routes.Add(new ServiceRoute("HighChartsFileService", new WebServiceHostFactory(), typeof(HighChartsFileService)));
            RouteTable.Routes.Add(new ServiceRoute("TestService", new WebServiceHostFactory(), typeof(TestService)));
            RouteTable.Routes.Add(new ServiceRoute("DHXFileService", new WebServiceHostFactory(), typeof(DHXFileService)));
            RouteTable.Routes.Add(new ServiceRoute("BTReportService", new WebServiceHostFactory(), typeof(BTReportService)));
        }
    }
}
