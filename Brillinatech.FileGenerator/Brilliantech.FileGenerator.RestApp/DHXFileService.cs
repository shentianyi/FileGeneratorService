﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ServiceModel;
using System.ServiceModel.Activation;
using System.ServiceModel.Web;
using DHX.Excel.Exporting;

namespace Brilliantech.FileGenerator.RestApp
{

    [ServiceContract]
    [AspNetCompatibilityRequirements(RequirementsMode = AspNetCompatibilityRequirementsMode.Allowed)]
    [ServiceBehavior(InstanceContextMode = InstanceContextMode.PerCall)]
    public class DHXFileService
    {
        [WebInvoke(Method = "POST", UriTemplate = "Excel")]
        public void GenerateExcel()
        {
            ExportExcel.ProcessExportRequest(HttpContext.Current);
        }
    }
}