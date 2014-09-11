using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Brilliantech.Export.Report.Chart;

namespace Brilliantech.Export.Report.XmlParser
{
    public class ChartXmlParser : ReportXmlParser
    {
        private RChart[] charts;

        public ChartXmlParser(string xml)
            : base(xml)
        {

        }

        public RChart[] GetCharts() {
            var nodes = root.SelectNodes(chartPath());
            if (nodes != null && nodes.Count > 0) {
                charts = new RChart[nodes.Count];
                for (int i = 0; i < nodes.Count; i++) {
                    charts[i] = new RChart(nodes[i]);
                }
            }
            return charts;
        }

        private string chartPath() {
            return "/report/charts/chart";
        }
    }
}
