using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;
using System.IO;
using System.ServiceModel.Web;

namespace Brillinatech.FileGeneratorService
{
    // 注意: 使用“重构”菜单上的“重命名”命令，可以同时更改代码和配置文件中的类名“Service1”。
    public class HighChartService : IHighChartService
    {
        public Stream GenerateHighChartsFile()
        {
           WebOperationContext context= WebOperationContext.Current;
           return null;
        }
    }
}
