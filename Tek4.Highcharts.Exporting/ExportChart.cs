// ExportChart.cs
// Tek4.Highcharts.Exporting.ExportChart class.
// Tek4.Highcharts.Exporting assembly.
// ==========================================================================
// <summary>
// Processes web requests to export Highcharts JS JavaScript charts.
// </summary>
// ==========================================================================
// Author: Kevin P. Rice, Tek4(TM) (http://Tek4.com/)
//
// Based upon ASP.NET Highcharts export module by Clément Agarini
//
// Copyright (C) 2011 by Tek4(TM) - Kevin P. Rice
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
// 
// REVISION HISTORY:
// 2011-07-17 KPR Created.
// 

namespace Tek4.Highcharts.Exporting
{
  using System;
  using System.Web;
    using System.Text.RegularExpressions;

  /// <summary>
  /// Processes web requests to export Highcharts JS JavaScript charts.
  /// </summary>
    public static class ExportChart
    {
        /// <summary>
        /// Processes HTTP Web requests to export SVG.
        /// </summary>
        /// <param name="context">An HttpContext object that provides references 
        /// to the intrinsic server objects (for example, Request, Response, 
        /// Session, and Server) used to service HTTP requests.</param>
        public static void ProcessExportRequest(HttpContext context)
        {
            if (context != null &&
              context.Request != null &&
              context.Response != null &&
              context.Request.HttpMethod == "POST")
            {
                HttpRequest request = context.Request;

                // Get HTTP POST form variables, ensuring they are not null.
                string filename = request.Form["filename"];
                string type = request.Form["type"];
                bool muti = (int.Parse( request.Form["muti"])==1);
                int width = 0;
                
                string[] svgs =null;
                if (muti)
                {
                    svgs = Regex.Split(request.Form["svg"], "</svg>,", RegexOptions.IgnoreCase);
                    for (int i = 0; i < svgs.Length - 1; i++)
                        svgs[i] = string.Concat(svgs[i], "</svg>");
                }
                else
                {
                    svgs = new string[1] { request.Form["svg"] };
                }
                if (filename != null &&
                  type != null &&
                  Int32.TryParse(request.Form["width"], out width) &&
                  request.Form["svg"] != null)
                {
                    // Create a new chart export object using form variables.
                    Exporter export = new Exporter(filename, type, width, svgs);

                    // Write the exported chart to the HTTP Response object.
                    export.WriteToHttpResponse(context.Response);

                    // Short-circuit this ASP.NET request and end. Short-circuiting
                    // prevents other modules from adding/interfering with the output.
                    HttpContext.Current.ApplicationInstance.CompleteRequest();
                    context.Response.End();
                }
            }
        }
    }
}