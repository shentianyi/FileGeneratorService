namespace OpenExcel.Common.FormulaParser
{
    using System;
    using System.Collections.Generic;

    public class ParseNode
    {
        protected List<ParseNode> nodes;
        public ParseNode Parent;
        protected string text;
        public OpenExcel.Common.FormulaParser.Token Token;

        protected ParseNode(OpenExcel.Common.FormulaParser.Token token, string text)
        {
            this.Token = token;
            this.text = text;
            this.nodes = new List<ParseNode>();
        }

        public virtual ParseNode CreateNode(OpenExcel.Common.FormulaParser.Token token, string text)
        {
            return new ParseNode(token, text) { Parent = this };
        }

        internal object Eval(ParseTree tree, params object[] paramlist)
        {
            switch (this.Token.Type)
            {
                case OpenExcel.Common.FormulaParser.TokenType.Start:
                    return this.EvalStart(tree, paramlist);

                case OpenExcel.Common.FormulaParser.TokenType.ComplexExpr:
                    return this.EvalComplexExpr(tree, paramlist);

                case OpenExcel.Common.FormulaParser.TokenType.Params:
                    return this.EvalParams(tree, paramlist);

                case OpenExcel.Common.FormulaParser.TokenType.ArrayElems:
                    return this.EvalArrayElems(tree, paramlist);

                case OpenExcel.Common.FormulaParser.TokenType.FuncCall:
                    return this.EvalFuncCall(tree, paramlist);

                case OpenExcel.Common.FormulaParser.TokenType.Array:
                    return this.EvalArray(tree, paramlist);

                case OpenExcel.Common.FormulaParser.TokenType.Range:
                    return this.EvalRange(tree, paramlist);

                case OpenExcel.Common.FormulaParser.TokenType.Expr:
                    return this.EvalExpr(tree, paramlist);
            }
            return this.Token.Text;
        }

        protected virtual object EvalArray(ParseTree tree, params object[] paramlist)
        {
            throw new NotImplementedException();
        }

        protected virtual object EvalArrayElems(ParseTree tree, params object[] paramlist)
        {
            throw new NotImplementedException();
        }

        protected virtual object EvalComplexExpr(ParseTree tree, params object[] paramlist)
        {
            throw new NotImplementedException();
        }

        protected virtual object EvalExpr(ParseTree tree, params object[] paramlist)
        {
            throw new NotImplementedException();
        }

        protected virtual object EvalFuncCall(ParseTree tree, params object[] paramlist)
        {
            throw new NotImplementedException();
        }

        protected virtual object EvalParams(ParseTree tree, params object[] paramlist)
        {
            throw new NotImplementedException();
        }

        protected virtual object EvalRange(ParseTree tree, params object[] paramlist)
        {
            throw new NotImplementedException();
        }

        protected virtual object EvalStart(ParseTree tree, params object[] paramlist)
        {
            return "Could not interpret input; no semantics implemented.";
        }

        protected object GetValue(ParseTree tree, OpenExcel.Common.FormulaParser.TokenType type, int index)
        {
            return this.GetValue(tree, type, ref index);
        }

        protected object GetValue(ParseTree tree, OpenExcel.Common.FormulaParser.TokenType type, ref int index)
        {
            if (index >= 0)
            {
                foreach (ParseNode node in this.nodes)
                {
                    if (node.Token.Type == type)
                    {
                        index--;
                        if (index < 0)
                        {
                            return node.Eval(tree, new object[0]);
                        }
                    }
                }
            }
            return null;
        }

        public List<ParseNode> Nodes
        {
            get
            {
                return this.nodes;
            }
        }

        public string Text
        {
            get
            {
                return this.text;
            }
            set
            {
                this.text = value;
            }
        }
    }
}

