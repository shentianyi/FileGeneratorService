namespace OpenExcel.Common
{
    using OpenExcel.Common.FormulaParser;
    using System;
    using System.Text;

    public static class ExcelFormula
    {
        private static Parser _parser = new Parser(new Scanner());

        private static void BuildTranslated(StringBuilder buf, ParseNode n, Func<ParseNode, string> translateFn)
        {
            foreach (ParseNode node in n.Nodes)
            {
                if (node.Token.Type == OpenExcel.Common.FormulaParser.TokenType.Range)
                {
                    buf.Append(translateFn(node));
                }
                else
                {
                    if (node.Token.Text != null)
                    {
                        buf.Append(node.Token.Text);
                    }
                    BuildTranslated(buf, node, translateFn);
                }
            }
        }

        public static ParseTree Parse(string formula)
        {
            return _parser.Parse(formula);
        }

        public static string Translate(string formula, int rowDelta, int colDelta)
        {
            if (formula == null)
            {
                return null;
            }
            ParseTree tree = Parse(formula);
            StringBuilder buf = new StringBuilder();
            if (tree.Errors.Count > 0)
            {
                throw new ArgumentException("Error in parsing formula");
            }
            BuildTranslated(buf, tree, n => TranslateRangeParseNodeWithOffset(n, rowDelta, colDelta));
            return buf.ToString();
        }

        public static string TranslateForSheetChange(string formula, SheetChange sheetChange, string currentSheetName)
        {
            if (formula == null)
            {
                return null;
            }
            ParseTree tree = Parse(formula);
            StringBuilder buf = new StringBuilder();
            if (tree.Errors.Count > 0)
            {
                throw new ArgumentException("Error in parsing formula");
            }
            BuildTranslated(buf, tree, n => TranslateRangeParseNodeForSheetChange(n, sheetChange, currentSheetName));
            return buf.ToString();
        }

        private static string TranslateRangeParseNodeForSheetChange(ParseNode rangeNode, SheetChange sheetChange, string currentSheetName)
        {
            string range = "";
            foreach (ParseNode node in rangeNode.Nodes)
            {
                range = range + node.Token.Text;
            }
            return ExcelRange.TranslateForSheetChange(range, sheetChange, currentSheetName);
        }

        private static string TranslateRangeParseNodeWithOffset(ParseNode rangeNode, int rowDelta, int colDelta)
        {
            string range = "";
            foreach (ParseNode node in rangeNode.Nodes)
            {
                range = range + node.Token.Text;
            }
            return ExcelRange.Translate(range, rowDelta, colDelta);
        }
    }
}

