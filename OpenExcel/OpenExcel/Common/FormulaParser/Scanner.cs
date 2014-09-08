namespace OpenExcel.Common.FormulaParser
{
    using System;
    using System.Collections.Generic;
    using System.Text.RegularExpressions;

    public class Scanner
    {
        public int CurrentColumn;
        public int CurrentLine;
        public int CurrentPosition;
        public int EndPos;
        public string Input;
        private Token LookAheadToken = null;
        public Dictionary<OpenExcel.Common.FormulaParser.TokenType, Regex> Patterns = new Dictionary<OpenExcel.Common.FormulaParser.TokenType, Regex>();
        private List<OpenExcel.Common.FormulaParser.TokenType> SkipList = new List<OpenExcel.Common.FormulaParser.TokenType>();
        public List<Token> Skipped;
        public int StartPos;
        private List<OpenExcel.Common.FormulaParser.TokenType> Tokens = new List<OpenExcel.Common.FormulaParser.TokenType>();

        public Scanner()
        {
            this.SkipList.Add(OpenExcel.Common.FormulaParser.TokenType.WSPC);
            Regex regex = new Regex(@"\(", RegexOptions.Compiled);
            this.Patterns.Add(OpenExcel.Common.FormulaParser.TokenType.PARENOPEN, regex);
            this.Tokens.Add(OpenExcel.Common.FormulaParser.TokenType.PARENOPEN);
            regex = new Regex(@"\)", RegexOptions.Compiled);
            this.Patterns.Add(OpenExcel.Common.FormulaParser.TokenType.PARENCLOSE, regex);
            this.Tokens.Add(OpenExcel.Common.FormulaParser.TokenType.PARENCLOSE);
            regex = new Regex(@"\{", RegexOptions.Compiled);
            this.Patterns.Add(OpenExcel.Common.FormulaParser.TokenType.BRACEOPEN, regex);
            this.Tokens.Add(OpenExcel.Common.FormulaParser.TokenType.BRACEOPEN);
            regex = new Regex(@"\}", RegexOptions.Compiled);
            this.Patterns.Add(OpenExcel.Common.FormulaParser.TokenType.BRACECLOSE, regex);
            this.Tokens.Add(OpenExcel.Common.FormulaParser.TokenType.BRACECLOSE);
            regex = new Regex(",", RegexOptions.Compiled);
            this.Patterns.Add(OpenExcel.Common.FormulaParser.TokenType.COMMA, regex);
            this.Tokens.Add(OpenExcel.Common.FormulaParser.TokenType.COMMA);
            regex = new Regex(@"\:", RegexOptions.Compiled);
            this.Patterns.Add(OpenExcel.Common.FormulaParser.TokenType.COLON, regex);
            this.Tokens.Add(OpenExcel.Common.FormulaParser.TokenType.COLON);
            regex = new Regex(";", RegexOptions.Compiled);
            this.Patterns.Add(OpenExcel.Common.FormulaParser.TokenType.SEMICOLON, regex);
            this.Tokens.Add(OpenExcel.Common.FormulaParser.TokenType.SEMICOLON);
            regex = new Regex(@"@?[A-Za-z_][A-Za-z0-9_]*(?=\()", RegexOptions.Compiled);
            this.Patterns.Add(OpenExcel.Common.FormulaParser.TokenType.FUNC, regex);
            this.Tokens.Add(OpenExcel.Common.FormulaParser.TokenType.FUNC);
            regex = new Regex(@"\#((NULL|DIV\/0|VALUE|REF|NUM)\!|NAME\?|N\/A)", RegexOptions.Compiled);
            this.Patterns.Add(OpenExcel.Common.FormulaParser.TokenType.ERR, regex);
            this.Tokens.Add(OpenExcel.Common.FormulaParser.TokenType.ERR);
            regex = new Regex(@"(?x)(  '([^*?\[\]\/\\\'\\]+)'  |  ([^*?\[\]\/\\\(\)\!\+\-\&\,]+)  )\!", RegexOptions.Compiled);
            this.Patterns.Add(OpenExcel.Common.FormulaParser.TokenType.SHEETNAME, regex);
            this.Tokens.Add(OpenExcel.Common.FormulaParser.TokenType.SHEETNAME);
            regex = new Regex(@"(\$)?([A-Za-z]+)(\$)?[0-9]+", RegexOptions.Compiled);
            this.Patterns.Add(OpenExcel.Common.FormulaParser.TokenType.ADDRESS, regex);
            this.Tokens.Add(OpenExcel.Common.FormulaParser.TokenType.ADDRESS);
            regex = new Regex("(?i)(NULL)", RegexOptions.Compiled);
            this.Patterns.Add(OpenExcel.Common.FormulaParser.TokenType.NULL, regex);
            this.Tokens.Add(OpenExcel.Common.FormulaParser.TokenType.NULL);
            regex = new Regex("(?i)(TRUE|FALSE)", RegexOptions.Compiled);
            this.Patterns.Add(OpenExcel.Common.FormulaParser.TokenType.BOOL, regex);
            this.Tokens.Add(OpenExcel.Common.FormulaParser.TokenType.BOOL);
            regex = new Regex(@"(\+|-)?[0-9]+(\.[0-9]+)?", RegexOptions.Compiled);
            this.Patterns.Add(OpenExcel.Common.FormulaParser.TokenType.NUMBER, regex);
            this.Tokens.Add(OpenExcel.Common.FormulaParser.TokenType.NUMBER);
            regex = new Regex("\\\"(\\\"\\\"|[^\\\"])*\\\"", RegexOptions.Compiled);
            this.Patterns.Add(OpenExcel.Common.FormulaParser.TokenType.STRING, regex);
            this.Tokens.Add(OpenExcel.Common.FormulaParser.TokenType.STRING);
            regex = new Regex(@"\*|/|\+|-|&|==|!=|<>|<=|>=|>|<|=", RegexOptions.Compiled);
            this.Patterns.Add(OpenExcel.Common.FormulaParser.TokenType.OP, regex);
            this.Tokens.Add(OpenExcel.Common.FormulaParser.TokenType.OP);
            regex = new Regex("^$", RegexOptions.Compiled);
            this.Patterns.Add(OpenExcel.Common.FormulaParser.TokenType.EOF, regex);
            this.Tokens.Add(OpenExcel.Common.FormulaParser.TokenType.EOF);
            regex = new Regex(@"\s+", RegexOptions.Compiled);
            this.Patterns.Add(OpenExcel.Common.FormulaParser.TokenType.WSPC, regex);
            this.Tokens.Add(OpenExcel.Common.FormulaParser.TokenType.WSPC);
        }

        public Token GetToken(OpenExcel.Common.FormulaParser.TokenType type)
        {
            return new Token(this.StartPos, this.EndPos) { Type = type };
        }

        public void Init(string input)
        {
            this.Input = input;
            this.StartPos = 0;
            this.EndPos = 0;
            this.CurrentLine = 0;
            this.CurrentColumn = 0;
            this.CurrentPosition = 0;
            this.Skipped = new List<Token>();
            this.LookAheadToken = null;
        }

        public Token LookAhead(params OpenExcel.Common.FormulaParser.TokenType[] expectedtokens)
        {
            List<OpenExcel.Common.FormulaParser.TokenType> tokens;
            int startPos = this.StartPos;
            Token item = null;
            if (((this.LookAheadToken != null) && (this.LookAheadToken.Type != OpenExcel.Common.FormulaParser.TokenType._UNDETERMINED_)) && (this.LookAheadToken.Type != OpenExcel.Common.FormulaParser.TokenType._NONE_))
            {
                return this.LookAheadToken;
            }
            if (expectedtokens.Length == 0)
            {
                tokens = this.Tokens;
            }
            else
            {
                tokens = new List<OpenExcel.Common.FormulaParser.TokenType>(expectedtokens);
                tokens.AddRange(this.SkipList);
            }
            do
            {
                int length = -1;
                OpenExcel.Common.FormulaParser.TokenType type = (OpenExcel.Common.FormulaParser.TokenType) 0x7fffffff;
                string input = this.Input.Substring(startPos);
                item = new Token(startPos, this.EndPos);
                for (int i = 0; i < tokens.Count; i++)
                {
                    Match match = this.Patterns[tokens[i]].Match(input);
                    if ((match.Success && (match.Index == 0)) && ((match.Length > length) || ((((OpenExcel.Common.FormulaParser.TokenType) tokens[i]) < type) && (match.Length == length))))
                    {
                        length = match.Length;
                        type = tokens[i];
                    }
                }
                if ((type >= OpenExcel.Common.FormulaParser.TokenType._NONE_) && (length >= 0))
                {
                    item.EndPos = startPos + length;
                    item.Text = this.Input.Substring(item.StartPos, length);
                    item.Type = type;
                }
                else if (item.StartPos < (item.EndPos - 1))
                {
                    item.Text = this.Input.Substring(item.StartPos, 1);
                }
                if (this.SkipList.Contains(item.Type))
                {
                    startPos = item.EndPos;
                    this.Skipped.Add(item);
                }
            }
            while (this.SkipList.Contains(item.Type));
            this.LookAheadToken = item;
            return item;
        }

        public Token Scan(params OpenExcel.Common.FormulaParser.TokenType[] expectedtokens)
        {
            Token token = this.LookAhead(expectedtokens);
            this.LookAheadToken = null;
            this.StartPos = token.EndPos;
            this.EndPos = token.EndPos;
            return token;
        }
    }
}

