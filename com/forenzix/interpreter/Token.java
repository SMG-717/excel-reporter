package com.forenzix.interpreter;

import java.util.Set;

/*
 * Token
 * 
 * A Token is a single coherent ordered sequence of characters. Additional information
 * can be stored in a Token about its type and precedence (if applicable). Tokens
 * that share the same ordered sequence of characters are NOT necessarily identical.
 * They must also share the same type and precedence. Tokens are also immutable.
 */
public class Token {

    final String value;
    final Set<TokenType> types;
    final int precedence;
    final boolean rightassoc;

    static final Token Empty = new Token("empty", TokenType.Keyword);
    static final Token If = new Token("if", TokenType.Keyword);
    static final Token Else = new Token("else", TokenType.Keyword);
    static final Token While = new Token("while", TokenType.Keyword);
    static final Token Let = new Token("let", TokenType.Keyword);
    static final Token Define = new Token("define", TokenType.Keyword);
    static final Token Return = new Token("return", TokenType.Keyword);
    static final Token True = new Token("true", TokenType.BooleanLiteral);
    static final Token False = new Token("false", TokenType.BooleanLiteral);

    static final Token Caret = new Token("^", TokenType.BinaryArithmetic, 8, true);
    static final Token Asterisk = new Token("*", TokenType.BinaryArithmetic, 7);
    static final Token ForwardSlash = new Token("/", TokenType.BinaryArithmetic, 7);
    static final Token Percent = new Token("%", TokenType.BinaryArithmetic, 7);
    static final Token Plus = new Token("+", TokenType.BinaryArithmetic, 6);
    static final Token ShiftLeft = new Token("<<", TokenType.BinaryArithmetic, 5);
    static final Token ShiftRight = new Token(">>", TokenType.BinaryArithmetic, 5);
    static final Token Greater = new Token(">", TokenType.BinaryArithmetic, 4);
    static final Token Less = new Token("<", TokenType.BinaryArithmetic, 4);
    static final Token GreaterEqual = new Token(">=", TokenType.BinaryArithmetic, 4);
    static final Token LessEqual = new Token("<=", TokenType.BinaryArithmetic, 4);
    static final Token Equals = new Token("==", TokenType.BinaryArithmetic, 3);
    static final Token NotEquals = new Token("!=", TokenType.BinaryArithmetic, 3);
    static final Token Ampersand = new Token("&", TokenType.BinaryArithmetic, 2);
    static final Token Pipe = new Token("|", TokenType.BinaryArithmetic, 2);
    static final Token Xor = new Token("xor", TokenType.BinaryArithmetic, 2);
    static final Token And = new Token("and", TokenType.BinaryArithmetic, 1);
    static final Token Or = new Token("or", TokenType.BinaryArithmetic, 1);
    static final Token Not = new Token("not", TokenType.UnaryArithmetic);
    static final Token Tilde = new Token("~", TokenType.UnaryArithmetic);

    static final Token At = new Token("@", TokenType.Punctuation);
    static final Token Underscore = new Token("_", TokenType.Punctuation);
    static final Token Hashtag = new Token("#", TokenType.Punctuation);
    static final Token Question = new Token("?", TokenType.Punctuation);
    static final Token Comma = new Token(",", TokenType.Punctuation);
    static final Token Colon = new Token(":", TokenType.Punctuation);
    static final Token Period = new Token(".", TokenType.Punctuation);
    static final Token BackSlash = new Token("\\", TokenType.Punctuation);
    static final Token OpenParen = new Token("(", TokenType.Punctuation);
    static final Token CloseParen = new Token(")", TokenType.Punctuation);
    static final Token OpenCurly = new Token("{", TokenType.Punctuation);
    static final Token OpenSquare = new Token("[", TokenType.Punctuation);
    static final Token CloseSquare = new Token("]", TokenType.Punctuation);
    static final Token DoubleQuote = new Token("\"", TokenType.Punctuation);
    static final Token SingleQuote = new Token("\'", TokenType.Punctuation);
    static final Token EqualSign = new Token("=", TokenType.Punctuation);

    static final Token CarriageReturn = new Token("\r", TokenType.WhiteSpace);
    static final Token Tab = new Token("\t", TokenType.WhiteSpace);
    static final Token BackSpace = new Token("\b", TokenType.WhiteSpace);

    static final Token Hyphen = new Token("-", Set.of(TokenType.BinaryArithmetic, TokenType.UnaryArithmetic), 6);
    static final Token Exclaim = new Token("!", Set.of(TokenType.Punctuation, TokenType.UnaryArithmetic), 0);
    static final Token SemiColon = new Token(";", Set.of(TokenType.Punctuation, TokenType.StatementTerminator), 0);
    static final Token CloseCurly = new Token("}", Set.of(TokenType.Punctuation, TokenType.ScopeTerminator), 0);
    static final Token Newline = new Token("\n", Set.of(TokenType.WhiteSpace, TokenType.StatementTerminator), 0);
    static final Token EOT = new Token("End", Set.of(TokenType.StatementTerminator, TokenType.ScopeTerminator), 0);

    static final char EOF = '\0';

    private Token(String val, TokenType type) {
        this(val, type, 0);
    }

    private Token(String val, TokenType type, int precedence, boolean rightassoc) {
        this(val, Set.of(type), precedence, rightassoc);
    }

    private Token(String val, TokenType type, int precedence) {
        this(val, Set.of(type), precedence, false);
    }

    private Token(String val, Set<TokenType> types, int precedence) {
        this(val, types, precedence, false);
    }

    private Token(String val, Set<TokenType> types, int precedence, boolean rightassoc) {
        this.value = val;
        this.types = Set.copyOf(types);
        this.precedence = precedence;
        this.rightassoc = rightassoc;
    }

    public boolean hasValue() {
        return !value.isBlank();
    }

    static Token makeToken(String name, TokenType type) {
        return new Token(name, type);
    }

    public boolean isAny(TokenType... types) {
        for (TokenType type : types) {
            if (this.types.contains(type)) {
                return true;
            }
        }
        return false;
    }

    public boolean isAll(TokenType... types) {
        for (TokenType type : types) {
            if (!this.types.contains(type)) {
                return false;
            }
        }
        return true;
    }

    @Override
    public String toString() {
        return this.value;
    }

    @Override
    public boolean equals(Object other) {
        if (!(other instanceof Token)) {
            return false;
        }

        Token token = (Token) other;
        return this.value.equals(token.value) && this.types.equals(token.types) && this.precedence == token.precedence;
    }

    @Override
    public int hashCode() {
        // Not sure how collision free this is. -SMG
        return (this.value.hashCode() + this.types.hashCode()) ^ this.precedence;
    }

}