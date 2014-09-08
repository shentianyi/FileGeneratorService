namespace OpenExcel.Common.FormulaParser
{
    using System;

    public class ParseError
    {
        private int code;
        private int col;
        private int length;
        private int line;
        private string message;
        private int pos;

        public ParseError(string message, int code, ParseNode node) : this(message, code, 0, node.Token.StartPos, node.Token.StartPos, node.Token.Length)
        {
        }

        public ParseError(string message, int code, int line, int col, int pos, int length)
        {
            this.message = message;
            this.code = code;
            this.line = line;
            this.col = col;
            this.pos = pos;
            this.length = length;
        }

        public int Code
        {
            get
            {
                return this.code;
            }
        }

        public int Column
        {
            get
            {
                return this.col;
            }
        }

        public int Length
        {
            get
            {
                return this.length;
            }
        }

        public int Line
        {
            get
            {
                return this.line;
            }
        }

        public string Message
        {
            get
            {
                return this.message;
            }
        }

        public int Position
        {
            get
            {
                return this.pos;
            }
        }
    }
}

