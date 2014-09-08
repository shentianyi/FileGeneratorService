namespace OpenExcel.Common.FormulaParser
{
    using System;

    public enum TokenType
    {
        _NONE_,
        _UNDETERMINED_,
        Start,
        ComplexExpr,
        Params,
        ArrayElems,
        FuncCall,
        Array,
        Range,
        Expr,
        PARENOPEN,
        PARENCLOSE,
        BRACEOPEN,
        BRACECLOSE,
        COMMA,
        COLON,
        SEMICOLON,
        FUNC,
        ERR,
        SHEETNAME,
        ADDRESS,
        NULL,
        BOOL,
        NUMBER,
        STRING,
        OP,
        EOF,
        WSPC
    }
}

