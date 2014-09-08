namespace OpenExcel.Common.FormulaParser
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    public class ParseTree : ParseNode
    {
        public ParseErrors Errors;
        public List<Token> Skipped;

        public ParseTree() : base(new Token(), "ParseTree")
        {
            base.Token.Type = OpenExcel.Common.FormulaParser.TokenType.Start;
            base.Token.Text = "Root";
            this.Skipped = new List<Token>();
            this.Errors = new ParseErrors();
        }

        public object Eval(params object[] paramlist)
        {
            return base.Nodes[0].Eval(this, paramlist);
        }

        private void PrintNode(StringBuilder sb, ParseNode node, int indent)
        {
            string str = "".PadLeft(indent, ' ');
            sb.Append(str);
            sb.AppendLine(node.Text);
            foreach (ParseNode node2 in node.Nodes)
            {
                this.PrintNode(sb, node2, indent + 2);
            }
        }

        public string PrintTree()
        {
            StringBuilder sb = new StringBuilder();
            int indent = 0;
            this.PrintNode(sb, this, indent);
            return sb.ToString();
        }
    }
}

