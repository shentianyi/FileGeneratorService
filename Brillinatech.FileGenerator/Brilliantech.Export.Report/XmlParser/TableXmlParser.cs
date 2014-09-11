using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using Brilliantech.Export.Report.Table;

namespace Brilliantech.Export.Report.XmlParser
{
    public class TableXmlParser : ReportXmlParser
    {
        protected RTColumn[][] columns;
        protected RTRow[] rows;
        private int[] widths;
        private Boolean header = false;
        private Boolean footer = false;
        private Boolean without_header = false;
        private string profile = "gray";

        public TableXmlParser(string xml):base(xml)
        { 
        }

        public RTColumn[][] GetColumnsInfo()
        {
            RTColumn[] colLine = null;
            try
            {
                var head_rows = root.SelectNodes(columnInfoPath());
                if (head_rows != null && head_rows.Count > 0)
                {
                    columns = new RTColumn[head_rows.Count][];
                    for (int i = 0; i < head_rows.Count; i++)
                    {

                        XmlElement cols = (XmlElement)head_rows[i];
                        var head_cols = cols.GetElementsByTagName("column");
                        if (head_cols != null && head_cols.Count > 0)
                        {
                            colLine = new RTColumn[head_cols.Count];
                            for (int j = 0; j < head_cols.Count; j++)
                            {
                                XmlElement col_xml = (XmlElement)head_cols[j];
                                RTColumn col = new RTColumn(col_xml);
                                colLine[j] = col;
                            }
                        }
                        columns[i] = colLine;
                    }
                }
            }
            catch (Exception e)
            {

            }
            createWidthsArray();
            optimizeColumns();
            return columns;
        }
        
        public RTRow[] GetRows()
        {
            var nodes = root.SelectNodes(columnPath());
            if (nodes != null && nodes.Count > 0) {
                rows = new RTRow[nodes.Count];
                for (int i = 0; i < nodes.Count; i++) {
                    rows[i] = new RTRow(nodes[i]);
                }
            }
            return rows;
        }
       
        private void createWidthsArray()
        {
            widths = new int[columns[0].Length];
            for (int i = 0; i < columns[0].Length; i++)
            {
                widths[i] = columns[0][i].Width;
            }
        }

        private string columnInfoPath( )
        {
            return "/report/table/head/columns";
        }

        private string columnPath()
        {
            return "/report/table/body/row";
        } 

        private void optimizeColumns()
        {
            for (int i = 1; i < columns.Length; i++)
            {
                for (int j = 0; j < columns[i].Length; j++)
                {
                    columns[i][j].Width= columns[0][j].Width;
                }
            }
            for (int i = 0; i < columns.Length; i++)
            {
                for (int j = 0; j < columns[i].Length; j++)
                {
                    if (columns[i][j].Colspan > 0)
                    {
                        for (int k = j + 1; k < j + columns[i][j].Colspan; k++)
                        {
                            columns[i][j].Width = columns[i][j].Width + columns[i][k].Width;
                            columns[i][k].Width=0;
                        }
                    }
                    if (columns[i][j].Rowspan > 0)
                    {
                        for (int k = i + 1; k < i + columns[i][j].Rowspan; k++)
                        {
                            columns[i][j].Height=columns[i][j].Height + 1;
                            columns[k][j].Height=0;
                        }
                    }
                }
            }
        }

        public int GetColCount() {
            return widths.Length;
        }

        public int[] Widths
        {
            get { return widths; }
        }
        public Boolean Header
        {
            get { return header; }
        }
        public Boolean Footer
        {
            get { return footer; }
        }
        public Boolean Without_header
        {
            get { return without_header; }
        }
        public string Profile
        {
            get { return profile; }
        }
    }
}
