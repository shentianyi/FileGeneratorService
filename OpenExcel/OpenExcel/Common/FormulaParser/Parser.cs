namespace OpenExcel.Common.FormulaParser
{
    using System;

    public class Parser
    {
        private Scanner scanner;
        private ParseTree tree;

        public Parser(Scanner scanner)
        {
            this.scanner = scanner;
        }

        public ParseTree Parse(string input)
        {
            this.tree = new ParseTree();
            return this.Parse(input, this.tree);
        }

        public ParseTree Parse(string input, ParseTree tree)
        {
            this.scanner.Init(input);
            this.tree = tree;
            this.ParseStart(tree);
            tree.Skipped = this.scanner.Skipped;
            return tree;
        }

        private void ParseArray(ParseNode parent)
        {
            ParseNode item = parent.CreateNode(this.scanner.GetToken(OpenExcel.Common.FormulaParser.TokenType.Array), "Array");
            parent.Nodes.Add(item);
            Token token = this.scanner.Scan(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.BRACEOPEN });
            if (token.Type != OpenExcel.Common.FormulaParser.TokenType.BRACEOPEN)
            {
                this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found. Expected " + OpenExcel.Common.FormulaParser.TokenType.BRACEOPEN.ToString(), 0x1001, 0, token.StartPos, token.StartPos, token.Length));
            }
            ParseNode node = item.CreateNode(token, token.ToString());
            item.Token.UpdateRange(token);
            item.Nodes.Add(node);
            this.ParseArrayElems(item);
            token = this.scanner.Scan(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.BRACECLOSE });
            if (token.Type != OpenExcel.Common.FormulaParser.TokenType.BRACECLOSE)
            {
                this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found. Expected " + OpenExcel.Common.FormulaParser.TokenType.BRACECLOSE.ToString(), 0x1001, 0, token.StartPos, token.StartPos, token.Length));
            }
            node = item.CreateNode(token, token.ToString());
            item.Token.UpdateRange(token);
            item.Nodes.Add(node);
            parent.Token.UpdateRange(item.Token);
        }

        private void ParseArrayElems(ParseNode parent)
        {
            ParseNode item = parent.CreateNode(this.scanner.GetToken(OpenExcel.Common.FormulaParser.TokenType.ArrayElems), "ArrayElems");
            parent.Nodes.Add(item);
            Token token = this.scanner.LookAhead(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.NULL, OpenExcel.Common.FormulaParser.TokenType.BOOL, OpenExcel.Common.FormulaParser.TokenType.NUMBER, OpenExcel.Common.FormulaParser.TokenType.STRING, OpenExcel.Common.FormulaParser.TokenType.PARENOPEN, OpenExcel.Common.FormulaParser.TokenType.BRACEOPEN, OpenExcel.Common.FormulaParser.TokenType.SHEETNAME, OpenExcel.Common.FormulaParser.TokenType.ADDRESS, OpenExcel.Common.FormulaParser.TokenType.ERR, OpenExcel.Common.FormulaParser.TokenType.FUNC });
            if ((((token.Type == OpenExcel.Common.FormulaParser.TokenType.NULL) || (token.Type == OpenExcel.Common.FormulaParser.TokenType.BOOL)) || ((token.Type == OpenExcel.Common.FormulaParser.TokenType.NUMBER) || (token.Type == OpenExcel.Common.FormulaParser.TokenType.STRING))) || ((((token.Type == OpenExcel.Common.FormulaParser.TokenType.PARENOPEN) || (token.Type == OpenExcel.Common.FormulaParser.TokenType.BRACEOPEN)) || ((token.Type == OpenExcel.Common.FormulaParser.TokenType.SHEETNAME) || (token.Type == OpenExcel.Common.FormulaParser.TokenType.ADDRESS))) || ((token.Type == OpenExcel.Common.FormulaParser.TokenType.ERR) || (token.Type == OpenExcel.Common.FormulaParser.TokenType.FUNC))))
            {
                this.ParseComplexExpr(item);
            }
            for (token = this.scanner.LookAhead(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.COMMA, OpenExcel.Common.FormulaParser.TokenType.SEMICOLON }); (token.Type == OpenExcel.Common.FormulaParser.TokenType.COMMA) || (token.Type == OpenExcel.Common.FormulaParser.TokenType.SEMICOLON); token = this.scanner.LookAhead(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.COMMA, OpenExcel.Common.FormulaParser.TokenType.SEMICOLON }))
            {
                ParseNode node;
                token = this.scanner.LookAhead(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.COMMA, OpenExcel.Common.FormulaParser.TokenType.SEMICOLON });
                switch (token.Type)
                {
                    case OpenExcel.Common.FormulaParser.TokenType.COMMA:
                        token = this.scanner.Scan(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.COMMA });
                        if (token.Type != OpenExcel.Common.FormulaParser.TokenType.COMMA)
                        {
                            this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found. Expected " + OpenExcel.Common.FormulaParser.TokenType.COMMA.ToString(), 0x1001, 0, token.StartPos, token.StartPos, token.Length));
                        }
                        node = item.CreateNode(token, token.ToString());
                        item.Token.UpdateRange(token);
                        item.Nodes.Add(node);
                        break;

                    case OpenExcel.Common.FormulaParser.TokenType.SEMICOLON:
                        token = this.scanner.Scan(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.SEMICOLON });
                        if (token.Type != OpenExcel.Common.FormulaParser.TokenType.SEMICOLON)
                        {
                            this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found. Expected " + OpenExcel.Common.FormulaParser.TokenType.SEMICOLON.ToString(), 0x1001, 0, token.StartPos, token.StartPos, token.Length));
                        }
                        node = item.CreateNode(token, token.ToString());
                        item.Token.UpdateRange(token);
                        item.Nodes.Add(node);
                        break;

                    default:
                        this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found.", 2, 0, token.StartPos, token.StartPos, token.Length));
                        break;
                }
                this.ParseComplexExpr(item);
            }
            parent.Token.UpdateRange(item.Token);
        }

        private void ParseComplexExpr(ParseNode parent)
        {
            ParseNode item = parent.CreateNode(this.scanner.GetToken(OpenExcel.Common.FormulaParser.TokenType.ComplexExpr), "ComplexExpr");
            parent.Nodes.Add(item);
            this.ParseExpr(item);
            for (Token token = this.scanner.LookAhead(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.OP }); token.Type == OpenExcel.Common.FormulaParser.TokenType.OP; token = this.scanner.LookAhead(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.OP }))
            {
                token = this.scanner.Scan(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.OP });
                if (token.Type != OpenExcel.Common.FormulaParser.TokenType.OP)
                {
                    this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found. Expected " + OpenExcel.Common.FormulaParser.TokenType.OP.ToString(), 0x1001, 0, token.StartPos, token.StartPos, token.Length));
                }
                ParseNode node = item.CreateNode(token, token.ToString());
                item.Token.UpdateRange(token);
                item.Nodes.Add(node);
                this.ParseExpr(item);
            }
            parent.Token.UpdateRange(item.Token);
        }

        private void ParseExpr(ParseNode parent)
        {
            ParseNode node;
            ParseNode item = parent.CreateNode(this.scanner.GetToken(OpenExcel.Common.FormulaParser.TokenType.Expr), "Expr");
            parent.Nodes.Add(item);
            Token token = this.scanner.LookAhead(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.NULL, OpenExcel.Common.FormulaParser.TokenType.BOOL, OpenExcel.Common.FormulaParser.TokenType.NUMBER, OpenExcel.Common.FormulaParser.TokenType.STRING, OpenExcel.Common.FormulaParser.TokenType.PARENOPEN, OpenExcel.Common.FormulaParser.TokenType.BRACEOPEN, OpenExcel.Common.FormulaParser.TokenType.SHEETNAME, OpenExcel.Common.FormulaParser.TokenType.ADDRESS, OpenExcel.Common.FormulaParser.TokenType.ERR, OpenExcel.Common.FormulaParser.TokenType.FUNC });
            switch (token.Type)
            {
                case OpenExcel.Common.FormulaParser.TokenType.PARENOPEN:
                    token = this.scanner.Scan(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.PARENOPEN });
                    if (token.Type != OpenExcel.Common.FormulaParser.TokenType.PARENOPEN)
                    {
                        this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found. Expected " + OpenExcel.Common.FormulaParser.TokenType.PARENOPEN.ToString(), 0x1001, 0, token.StartPos, token.StartPos, token.Length));
                    }
                    node = item.CreateNode(token, token.ToString());
                    item.Token.UpdateRange(token);
                    item.Nodes.Add(node);
                    this.ParseExpr(item);
                    token = this.scanner.Scan(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.PARENCLOSE });
                    if (token.Type != OpenExcel.Common.FormulaParser.TokenType.PARENCLOSE)
                    {
                        this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found. Expected " + OpenExcel.Common.FormulaParser.TokenType.PARENCLOSE.ToString(), 0x1001, 0, token.StartPos, token.StartPos, token.Length));
                    }
                    node = item.CreateNode(token, token.ToString());
                    item.Token.UpdateRange(token);
                    item.Nodes.Add(node);
                    break;

                case OpenExcel.Common.FormulaParser.TokenType.BRACEOPEN:
                    this.ParseArray(item);
                    break;

                case OpenExcel.Common.FormulaParser.TokenType.FUNC:
                    this.ParseFuncCall(item);
                    break;

                case OpenExcel.Common.FormulaParser.TokenType.ERR:
                case OpenExcel.Common.FormulaParser.TokenType.SHEETNAME:
                case OpenExcel.Common.FormulaParser.TokenType.ADDRESS:
                    this.ParseRange(item);
                    break;

                case OpenExcel.Common.FormulaParser.TokenType.NULL:
                    token = this.scanner.Scan(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.NULL });
                    if (token.Type != OpenExcel.Common.FormulaParser.TokenType.NULL)
                    {
                        this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found. Expected " + OpenExcel.Common.FormulaParser.TokenType.NULL.ToString(), 0x1001, 0, token.StartPos, token.StartPos, token.Length));
                    }
                    node = item.CreateNode(token, token.ToString());
                    item.Token.UpdateRange(token);
                    item.Nodes.Add(node);
                    break;

                case OpenExcel.Common.FormulaParser.TokenType.BOOL:
                    token = this.scanner.Scan(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.BOOL });
                    if (token.Type != OpenExcel.Common.FormulaParser.TokenType.BOOL)
                    {
                        this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found. Expected " + OpenExcel.Common.FormulaParser.TokenType.BOOL.ToString(), 0x1001, 0, token.StartPos, token.StartPos, token.Length));
                    }
                    node = item.CreateNode(token, token.ToString());
                    item.Token.UpdateRange(token);
                    item.Nodes.Add(node);
                    break;

                case OpenExcel.Common.FormulaParser.TokenType.NUMBER:
                    token = this.scanner.Scan(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.NUMBER });
                    if (token.Type != OpenExcel.Common.FormulaParser.TokenType.NUMBER)
                    {
                        this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found. Expected " + OpenExcel.Common.FormulaParser.TokenType.NUMBER.ToString(), 0x1001, 0, token.StartPos, token.StartPos, token.Length));
                    }
                    node = item.CreateNode(token, token.ToString());
                    item.Token.UpdateRange(token);
                    item.Nodes.Add(node);
                    break;

                case OpenExcel.Common.FormulaParser.TokenType.STRING:
                    token = this.scanner.Scan(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.STRING });
                    if (token.Type != OpenExcel.Common.FormulaParser.TokenType.STRING)
                    {
                        this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found. Expected " + OpenExcel.Common.FormulaParser.TokenType.STRING.ToString(), 0x1001, 0, token.StartPos, token.StartPos, token.Length));
                    }
                    node = item.CreateNode(token, token.ToString());
                    item.Token.UpdateRange(token);
                    item.Nodes.Add(node);
                    break;

                default:
                    this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found.", 2, 0, token.StartPos, token.StartPos, token.Length));
                    break;
            }
            parent.Token.UpdateRange(item.Token);
        }

        private void ParseFuncCall(ParseNode parent)
        {
            ParseNode item = parent.CreateNode(this.scanner.GetToken(OpenExcel.Common.FormulaParser.TokenType.FuncCall), "FuncCall");
            parent.Nodes.Add(item);
            Token token = this.scanner.Scan(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.FUNC });
            if (token.Type != OpenExcel.Common.FormulaParser.TokenType.FUNC)
            {
                this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found. Expected " + OpenExcel.Common.FormulaParser.TokenType.FUNC.ToString(), 0x1001, 0, token.StartPos, token.StartPos, token.Length));
            }
            ParseNode node = item.CreateNode(token, token.ToString());
            item.Token.UpdateRange(token);
            item.Nodes.Add(node);
            token = this.scanner.Scan(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.PARENOPEN });
            if (token.Type != OpenExcel.Common.FormulaParser.TokenType.PARENOPEN)
            {
                this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found. Expected " + OpenExcel.Common.FormulaParser.TokenType.PARENOPEN.ToString(), 0x1001, 0, token.StartPos, token.StartPos, token.Length));
            }
            node = item.CreateNode(token, token.ToString());
            item.Token.UpdateRange(token);
            item.Nodes.Add(node);
            token = this.scanner.LookAhead(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.NULL, OpenExcel.Common.FormulaParser.TokenType.BOOL, OpenExcel.Common.FormulaParser.TokenType.NUMBER, OpenExcel.Common.FormulaParser.TokenType.STRING, OpenExcel.Common.FormulaParser.TokenType.PARENOPEN, OpenExcel.Common.FormulaParser.TokenType.BRACEOPEN, OpenExcel.Common.FormulaParser.TokenType.SHEETNAME, OpenExcel.Common.FormulaParser.TokenType.ADDRESS, OpenExcel.Common.FormulaParser.TokenType.ERR, OpenExcel.Common.FormulaParser.TokenType.FUNC, OpenExcel.Common.FormulaParser.TokenType.COMMA });
            if ((((token.Type == OpenExcel.Common.FormulaParser.TokenType.NULL) || (token.Type == OpenExcel.Common.FormulaParser.TokenType.BOOL)) || ((token.Type == OpenExcel.Common.FormulaParser.TokenType.NUMBER) || (token.Type == OpenExcel.Common.FormulaParser.TokenType.STRING))) || ((((token.Type == OpenExcel.Common.FormulaParser.TokenType.PARENOPEN) || (token.Type == OpenExcel.Common.FormulaParser.TokenType.BRACEOPEN)) || ((token.Type == OpenExcel.Common.FormulaParser.TokenType.SHEETNAME) || (token.Type == OpenExcel.Common.FormulaParser.TokenType.ADDRESS))) || (((token.Type == OpenExcel.Common.FormulaParser.TokenType.ERR) || (token.Type == OpenExcel.Common.FormulaParser.TokenType.FUNC)) || (token.Type == OpenExcel.Common.FormulaParser.TokenType.COMMA))))
            {
                this.ParseParams(item);
            }
            token = this.scanner.Scan(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.PARENCLOSE });
            if (token.Type != OpenExcel.Common.FormulaParser.TokenType.PARENCLOSE)
            {
                this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found. Expected " + OpenExcel.Common.FormulaParser.TokenType.PARENCLOSE.ToString(), 0x1001, 0, token.StartPos, token.StartPos, token.Length));
            }
            node = item.CreateNode(token, token.ToString());
            item.Token.UpdateRange(token);
            item.Nodes.Add(node);
            parent.Token.UpdateRange(item.Token);
        }

        private void ParseParams(ParseNode parent)
        {
            ParseNode item = parent.CreateNode(this.scanner.GetToken(OpenExcel.Common.FormulaParser.TokenType.Params), "Params");
            parent.Nodes.Add(item);
            Token token = this.scanner.LookAhead(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.NULL, OpenExcel.Common.FormulaParser.TokenType.BOOL, OpenExcel.Common.FormulaParser.TokenType.NUMBER, OpenExcel.Common.FormulaParser.TokenType.STRING, OpenExcel.Common.FormulaParser.TokenType.PARENOPEN, OpenExcel.Common.FormulaParser.TokenType.BRACEOPEN, OpenExcel.Common.FormulaParser.TokenType.SHEETNAME, OpenExcel.Common.FormulaParser.TokenType.ADDRESS, OpenExcel.Common.FormulaParser.TokenType.ERR, OpenExcel.Common.FormulaParser.TokenType.FUNC });
            if ((((token.Type == OpenExcel.Common.FormulaParser.TokenType.NULL) || (token.Type == OpenExcel.Common.FormulaParser.TokenType.BOOL)) || ((token.Type == OpenExcel.Common.FormulaParser.TokenType.NUMBER) || (token.Type == OpenExcel.Common.FormulaParser.TokenType.STRING))) || ((((token.Type == OpenExcel.Common.FormulaParser.TokenType.PARENOPEN) || (token.Type == OpenExcel.Common.FormulaParser.TokenType.BRACEOPEN)) || ((token.Type == OpenExcel.Common.FormulaParser.TokenType.SHEETNAME) || (token.Type == OpenExcel.Common.FormulaParser.TokenType.ADDRESS))) || ((token.Type == OpenExcel.Common.FormulaParser.TokenType.ERR) || (token.Type == OpenExcel.Common.FormulaParser.TokenType.FUNC))))
            {
                this.ParseComplexExpr(item);
            }
            for (token = this.scanner.LookAhead(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.COMMA }); token.Type == OpenExcel.Common.FormulaParser.TokenType.COMMA; token = this.scanner.LookAhead(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.COMMA }))
            {
                token = this.scanner.Scan(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.COMMA });
                if (token.Type != OpenExcel.Common.FormulaParser.TokenType.COMMA)
                {
                    this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found. Expected " + OpenExcel.Common.FormulaParser.TokenType.COMMA.ToString(), 0x1001, 0, token.StartPos, token.StartPos, token.Length));
                }
                ParseNode node = item.CreateNode(token, token.ToString());
                item.Token.UpdateRange(token);
                item.Nodes.Add(node);
                this.ParseComplexExpr(item);
            }
            parent.Token.UpdateRange(item.Token);
        }

        private void ParseRange(ParseNode parent)
        {
            Token token;
            ParseNode node;
            ParseNode item = parent.CreateNode(this.scanner.GetToken(OpenExcel.Common.FormulaParser.TokenType.Range), "Range");
            parent.Nodes.Add(item);
            if (this.scanner.LookAhead(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.SHEETNAME }).Type == OpenExcel.Common.FormulaParser.TokenType.SHEETNAME)
            {
                token = this.scanner.Scan(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.SHEETNAME });
                if (token.Type != OpenExcel.Common.FormulaParser.TokenType.SHEETNAME)
                {
                    this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found. Expected " + OpenExcel.Common.FormulaParser.TokenType.SHEETNAME.ToString(), 0x1001, 0, token.StartPos, token.StartPos, token.Length));
                }
                node = item.CreateNode(token, token.ToString());
                item.Token.UpdateRange(token);
                item.Nodes.Add(node);
            }
            token = this.scanner.LookAhead(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.ADDRESS, OpenExcel.Common.FormulaParser.TokenType.ERR });
            switch (token.Type)
            {
                case OpenExcel.Common.FormulaParser.TokenType.ERR:
                    token = this.scanner.Scan(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.ERR });
                    if (token.Type != OpenExcel.Common.FormulaParser.TokenType.ERR)
                    {
                        this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found. Expected " + OpenExcel.Common.FormulaParser.TokenType.ERR.ToString(), 0x1001, 0, token.StartPos, token.StartPos, token.Length));
                    }
                    node = item.CreateNode(token, token.ToString());
                    item.Token.UpdateRange(token);
                    item.Nodes.Add(node);
                    break;

                case OpenExcel.Common.FormulaParser.TokenType.ADDRESS:
                    token = this.scanner.Scan(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.ADDRESS });
                    if (token.Type != OpenExcel.Common.FormulaParser.TokenType.ADDRESS)
                    {
                        this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found. Expected " + OpenExcel.Common.FormulaParser.TokenType.ADDRESS.ToString(), 0x1001, 0, token.StartPos, token.StartPos, token.Length));
                    }
                    node = item.CreateNode(token, token.ToString());
                    item.Token.UpdateRange(token);
                    item.Nodes.Add(node);
                    break;

                default:
                    this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found.", 2, 0, token.StartPos, token.StartPos, token.Length));
                    break;
            }
            if (this.scanner.LookAhead(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.COLON }).Type == OpenExcel.Common.FormulaParser.TokenType.COLON)
            {
                token = this.scanner.Scan(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.COLON });
                if (token.Type != OpenExcel.Common.FormulaParser.TokenType.COLON)
                {
                    this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found. Expected " + OpenExcel.Common.FormulaParser.TokenType.COLON.ToString(), 0x1001, 0, token.StartPos, token.StartPos, token.Length));
                }
                node = item.CreateNode(token, token.ToString());
                item.Token.UpdateRange(token);
                item.Nodes.Add(node);
                token = this.scanner.LookAhead(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.ADDRESS, OpenExcel.Common.FormulaParser.TokenType.ERR });
                switch (token.Type)
                {
                    case OpenExcel.Common.FormulaParser.TokenType.ERR:
                        token = this.scanner.Scan(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.ERR });
                        if (token.Type != OpenExcel.Common.FormulaParser.TokenType.ERR)
                        {
                            this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found. Expected " + OpenExcel.Common.FormulaParser.TokenType.ERR.ToString(), 0x1001, 0, token.StartPos, token.StartPos, token.Length));
                        }
                        node = item.CreateNode(token, token.ToString());
                        item.Token.UpdateRange(token);
                        item.Nodes.Add(node);
                        goto Label_059B;

                    case OpenExcel.Common.FormulaParser.TokenType.ADDRESS:
                        token = this.scanner.Scan(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.ADDRESS });
                        if (token.Type != OpenExcel.Common.FormulaParser.TokenType.ADDRESS)
                        {
                            this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found. Expected " + OpenExcel.Common.FormulaParser.TokenType.ADDRESS.ToString(), 0x1001, 0, token.StartPos, token.StartPos, token.Length));
                        }
                        node = item.CreateNode(token, token.ToString());
                        item.Token.UpdateRange(token);
                        item.Nodes.Add(node);
                        goto Label_059B;
                }
                this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found.", 2, 0, token.StartPos, token.StartPos, token.Length));
            }
        Label_059B:
            parent.Token.UpdateRange(item.Token);
        }

        private void ParseStart(ParseNode parent)
        {
            ParseNode item = parent.CreateNode(this.scanner.GetToken(OpenExcel.Common.FormulaParser.TokenType.Start), "Start");
            parent.Nodes.Add(item);
            this.ParseComplexExpr(item);
            Token token = this.scanner.Scan(new OpenExcel.Common.FormulaParser.TokenType[] { OpenExcel.Common.FormulaParser.TokenType.EOF });
            if (token.Type != OpenExcel.Common.FormulaParser.TokenType.EOF)
            {
                this.tree.Errors.Add(new ParseError("Unexpected token '" + token.Text.Replace("\n", "") + "' found. Expected " + OpenExcel.Common.FormulaParser.TokenType.EOF.ToString(), 0x1001, 0, token.StartPos, token.StartPos, token.Length));
            }
            ParseNode node = item.CreateNode(token, token.ToString());
            item.Token.UpdateRange(token);
            item.Nodes.Add(node);
            parent.Token.UpdateRange(item.Token);
        }
    }
}

