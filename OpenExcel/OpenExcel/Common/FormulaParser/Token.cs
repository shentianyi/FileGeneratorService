namespace OpenExcel.Common.FormulaParser
{
    using System;

    public class Token
    {
        private int endpos;
        private int startpos;
        private string text;
        public OpenExcel.Common.FormulaParser.TokenType Type;
        private object value;

        public Token() : this(0, 0)
        {
        }

        public Token(int start, int end)
        {
            this.Type = OpenExcel.Common.FormulaParser.TokenType._UNDETERMINED_;
            this.startpos = start;
            this.endpos = end;
            this.Text = "";
            this.Value = null;
        }

        public override string ToString()
        {
            if (this.Text != null)
            {
                return (this.Type.ToString() + " '" + this.Text + "'");
            }
            return this.Type.ToString();
        }

        public void UpdateRange(Token token)
        {
            if (token.StartPos < this.startpos)
            {
                this.startpos = token.StartPos;
            }
            if (token.EndPos > this.endpos)
            {
                this.endpos = token.EndPos;
            }
        }

        public int EndPos
        {
            get
            {
                return this.endpos;
            }
            set
            {
                this.endpos = value;
            }
        }

        public int Length
        {
            get
            {
                return (this.endpos - this.startpos);
            }
        }

        public int StartPos
        {
            get
            {
                return this.startpos;
            }
            set
            {
                this.startpos = value;
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

        public object Value
        {
            get
            {
                return this.value;
            }
            set
            {
                this.value = value;
            }
        }
    }
}

