<%@ Page Language="C#" %>

<!-- 
==========================================================================
ASP.NET Highcharts Exporting Module Demo.
Uses Tek4.Highcharts.Exporting assembly.

(Uses Highcharts JS "Basic Line" example from Highcharts.com).
==========================================================================
Author: Kevin P. Rice, Tek4(TM) (http://Tek4.com/)

Based upon ASP.NET Highcharts export module by Clément Agarini

Copyright (C) 2011 by Tek4(TM) - Kevin P. Rice

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

rev. 2011-12-24 Latest Svg.dll requires .NET 3.5
rev. 2011-08-18 .NET 2.0
-->
<!DOCTYPE html>
<html>
<head runat="server">
    <title>ASP.NET Highcharts Exporting Module Demo (Tek4.Highcharts.Exporting assembly)
    </title>
    <!-- 1. Include jQuery and Highcharts scripts. -->
    <script src="js/jquery-1.8.3.min.js" type="text/javascript"></script>
    <script src="js/highcharts.src.js" type="text/javascript"></script>
    <!-- 2. Include the Highcharts exporting module script. -->
    <script src="js/exporting.src.js" type="text/javascript"></script>
    <script src="js/exporting.extend.src.js" type="text/javascript"></script>
    <script src="js/init.js" type="text/javascript"></script>
    <!-- 3. DON'T FORGET to add the exporting url to your chart (along with
            any other desired exporting options):
    
          exporting: { 
            url: "HighchartsExport.axd",
            filename: 'MyChart',
            width: 1200
          }
    -->
  
</head>
<body>
    <div>
        <h2>
            Tek4 ASP.NET Exporting Module for Highcharts JS Demo.</h2>
        <ul style="font-family: Verdana, Courier">
            <li>Exports Highcharts JS charts to PNG/JPG/PDF/SVG.</li>
            <li>Uses three precompiled .DLL files in /bin directory and configuration via web.config.</li>
            <li>Can be called as either an ASP.NET page (HighchartsExport.aspx) or as an HttpHandler
                (HighchartsExport.axd).</li>
            <li>Supports Highcharts exporting 'width' option to generate high quality images of
                any size (PDF images are exported at 150 dpi).</li>
            <li>Supports Highcharts exporting 'filename' option to specify downloaded file name.</li>
            <li>Works with .NET 3.5 Framework and above.</li>
        </ul>
    </div> 
    <p>       Highcharts JS "Basic Line" example from Highcharts.com:
     <input type="button" id="exportPDFs_btn" value="Export PDFs"/>
     <input type="button" id="exportDOCs_btn" value="Export DOCs"/>
     <input type="button" id="exportDOCXs_btn" value="Export DOCXs"/>
          <input type="button" id="exportXLSXs_btn" value="Export XLSXs"/>
     </p>
    <div id="container_1" style="width: 900px;" ></div>
    <div id="container_2" style="width: 900px;" ></div>
     <div id="container_3" style="width: 900px;" ></div>
</body>
</html>
