using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using System.IO;

namespace Brillinatech.FileGeneratorService
{ 
    [ServiceContract]
    public interface IHighChartService
    {
        [OperationContract]
        [WebInvoke(Method = "POST", UriTemplate = "/generate_highcharts_file")]
        Stream GenerateHighChartsFile();
    }
}
