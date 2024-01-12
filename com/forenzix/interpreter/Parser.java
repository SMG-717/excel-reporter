package com.forenzix.interpreter;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;
import java.util.function.Supplier;
import java.util.stream.Collectors;

public class Parser {
    
    private final List<Token> cache;
    private final Tokeniser tokeniser;
    private int lineNumber = 1;
    private NodeScope root = null;

    public Parser(Tokeniser tokeniser) {
        this.tokeniser = tokeniser;
        this.cache = new LinkedList<>();
    }
    
    public Parser(String input) {
        // The input expression should always end in an End of File token (\0 character)
        this(new Tokeniser(input + Token.EOF));
    }

    private RuntimeException error(String message) {
        return new RuntimeException(message + " (line: " + lineNumber + ")");
    }

    /***************************************************************************
     * Parser
     **************************************************************************/
    public NodeScope parse() {
        tokeniser.reset();
        lineNumber = 1;
        root = parseProgram();

        if (peek() != Token.EOT) {
            throw error("Unexpected token at End of Expression: " + peek().value);
        }
        return root;
    }

    private <T> T tryParse(Supplier<T> supplier, String error) {
        final T node =  supplier.get();
        if (node == null) {
            throw error(error);
        }
        return node;
    }

    private NodeScope parseScope() {
        tryConsume(Token.OpenCurly, "Expected '{'");
        final NodeScope scope = tryParse(() -> parseProgram(), "Unparsable Scope.");
        tryConsume(Token.CloseCurly, "Expected '}'");

        return scope;
    }

    private NodeScope parseProgram() {
        List<NodeStatement> statements = new ArrayList<>();
        while (peek() != Token.EOT) {
            
            if (peek().isAny(TokenType.StatementTerminator, TokenType.Comment)) {
                consume();
                continue;
            }
            
            if (peek().isAny(TokenType.ScopeTerminator))
                break;

                
            NodeStatement statement = tryParse(() -> parseStatement(), "Incomplete/Unparsable statement");
            if (!peek().isAny(TokenType.StatementTerminator, TokenType.StatementTerminator, TokenType.Comment)) {
                throw error("Statement must end in a termination character ('\\n', or ';' or EOT)");
            }
            
            statements.add(statement);
        }

        return new NodeScope(statements);
    }

    private NodeStatement parseStatement() {

        final NodeStatement statement;
        final int cacheLineNumber = lineNumber;

        // Declaration
        if (tryConsume(Token.Let)) {
            
            final NodeVariable var = tryParse(() -> parseVariable(), "Expected qualifier: " + peek().value);
            tryConsume(Token.EqualSign, "Expected '='");

            final NodeExpression expr = parseExpression();
            statement = expr == null ? null : new NodeStatement.Declare(var, expr);
        }

        // If statement
        else if (tryConsume(Token.If)) {
            final NodeExpression expr = tryParse(() -> parseExpression(), "Expected if statement expression.");
            final NodeScope scope = tryParse(() -> parseScope(), "Unparsable Scope.");

            int ahead = 0;
            while (peek(ahead) == Token.Newline || peek(ahead).isAny(TokenType.Comment)) {
                ahead += 1;
            }
            
            final NodeScope scopeElse;
            if (peek(ahead) == Token.Else) {
                while (peek() != Token.Else) consume();
                consume();
                
                if (peek() == Token.OpenCurly) {
                    scopeElse = tryParse(() -> parseScope(), "Unparsable Scope.");
                }
                else if (peek() == Token.If) {
                    scopeElse = new NodeScope(List.of(parseStatement()));
                }
                else {
                    throw error("Expected else statement definition");
                }
            }
            else {
                scopeElse = null;
            }
            statement = new NodeStatement.If(expr, scope, scopeElse);
        }

        // While loop
        else if (tryConsume(Token.While)) {
            final NodeExpression expr = tryParse(() -> parseExpression(), "Expected while statement expression.");
            final NodeScope scope = tryParse(() -> parseScope(), "Unparsable Scope.");
            statement = new NodeStatement.While(expr, scope);
        }

        // Standalone scope
        else if (peek() == Token.OpenCurly) {
            final NodeScope scope = tryParse(() -> parseScope(), "Expected scope definition");
            statement = new NodeStatement.Scope(scope);
        }

        // Assignment
        else if (peek().isAny(TokenType.Qualifier) && peek(1) == Token.EqualSign) {

            final NodeVariable var = tryParse(() -> parseVariable(), "Expected qualifier: " + peek().value);
            tryConsume(Token.EqualSign, "Expected '='");
            
            final NodeExpression expr = parseExpression();
            statement = expr == null ? null : new NodeStatement.Assign(var, expr);
        }

        // Member Assignment
        else if (peek().isAny(TokenType.Qualifier) && peek(1) == Token.Period && peek(2).isAny(TokenType.Qualifier) && peek(3) == Token.EqualSign) {

            final NodeVariable var = tryParse(() -> parseVariable(), "Expected qualifier: " + peek().value);
            tryConsume(Token.Period, "Expected '.'");
            final NodeVariable member = tryParse(() -> parseVariable(), "Expected member: " + peek().value);
            tryConsume(Token.EqualSign, "Expected '='");
            final NodeExpression expr = parseExpression();
            
            statement = expr == null ? null : new NodeStatement.MemberAssign(var, member, expr);
        }

        // Bare Expressions
        else {
            final NodeExpression expr = parseExpression();
            statement = expr == null ? null : new NodeStatement.Expression(expr);
        }

        if (statement != null)
            statement.lineNumber = cacheLineNumber;

        return statement;
    }

    private NodeExpression parseExpression() { 
        return parseExpression(parseAtom(), 0); 
    }

    private NodeExpression parseExpression(NodeExpression left, int prec) {

        while (peek().isAny(TokenType.BinaryArithmetic) && peek().precedence >= prec) {
            final Token op = consume();
            NodeExpression right = parseAtom();

            while (peek().isAny(TokenType.BinaryArithmetic) && (
                (peek().precedence > op.precedence) || 
                (peek().precedence >= op.precedence && peek().rightassoc))) {
                right = parseExpression(right, op.precedence + (peek().precedence > op.precedence ? 1 : 0));
            }
            left = arthmeticNode(op, left, right);
        }

        return left;
    }

    private NodeExpression arthmeticNode(Token op, NodeExpression lhs, NodeExpression rhs) {
        if (op == Token.Plus)                   return new NodeExpression.Binary(BinaryOperator.Add, lhs, rhs);
        else if (op == Token.Hyphen)            return new NodeExpression.Binary(BinaryOperator.Subtract, lhs, rhs);
        else if (op == Token.Asterisk)          return new NodeExpression.Binary(BinaryOperator.Multiply, lhs, rhs);
        else if (op == Token.ForwardSlash)      return new NodeExpression.Binary(BinaryOperator.Divide, lhs, rhs);
        else if (op == Token.Percent)           return new NodeExpression.Binary(BinaryOperator.Modulo, lhs, rhs);
        else if (op == Token.Caret)             return new NodeExpression.Binary(BinaryOperator.Exponent, lhs, rhs);
        else if (op == Token.Greater)           return new NodeExpression.Binary(BinaryOperator.Greater, lhs, rhs);
        else if (op == Token.GreaterEqual)      return new NodeExpression.Binary(BinaryOperator.GreaterEqual, lhs, rhs);
        else if (op == Token.Less)              return new NodeExpression.Binary(BinaryOperator.Less, lhs, rhs);
        else if (op == Token.LessEqual)         return new NodeExpression.Binary(BinaryOperator.LessEqual, lhs, rhs);
        else if (op == Token.Equals)            return new NodeExpression.Binary(BinaryOperator.Equal, lhs, rhs);
        else if (op == Token.NotEquals)         return new NodeExpression.Binary(BinaryOperator.NotEqual, lhs, rhs);
        else if (op == Token.And)               return new NodeExpression.Binary(BinaryOperator.And, lhs, rhs);
        else if (op == Token.Or)                return new NodeExpression.Binary(BinaryOperator.Or, lhs, rhs);
        else {
            throw error("Unsupported arithmetic operation: " + peek().value);
        }
    }

    private NodeExpression arthmeticNode(Token op, NodeExpression val) {
        if (op == Token.Hyphen) return new NodeExpression.Unary(UnaryOperator.Negate, val);
        else if (op == Token.Tilde) return new NodeExpression.Unary(UnaryOperator.Invert, val);
        else if (op == Token.Not) return new NodeExpression.Unary(UnaryOperator.Not, val);
        else if (op == Token.Exclaim) return new NodeExpression.Unary(UnaryOperator.Not, val);
        else {
            throw error("Unsupported unary arithmetic operation: " + peek().value);
        }
    }

    /*
     * Parse atom
     * 
     * Atomic expressions can be either Variables (such as X), Literals (such as 10), or smaller Terms surrounded by 
     * parentheses. This gives parentheses the highest precedence among operations.
     * 
     * The production for atom is as follows:
     * 
     * Atom -> ([Term]) | [Variable] | [Literal]
     */

    private NodeExpression parseAtom() {
        if (tryConsume(Token.OpenParen)) {
            final NodeExpression exp = tryParse(() -> parseExpression(), "Unparsable/Invalid expression");
            tryConsume(Token.CloseParen, "Expected ')'");
            return exp;
        } else if (peek().isAny(TokenType.UnaryArithmetic)) {
            final Token op = consume();
            final NodeExpression expr = tryParse(() -> parseAtom(), "Unsupported unary operation: " + op);
            return arthmeticNode(op, expr);
        }
        else {            
            final NodeTerm atom;
            if (peek().isAny(TokenType.Qualifier)) {
                final NodeVariable var = parseVariable();
                if (peek() != Token.Period) {
                    atom = new NodeTerm.Variable(var);
                }
                else {
                    List<String> memberBuilder = new LinkedList<>();
                    while (tryConsume(Token.Period)) {
                        final NodeVariable member = tryParse(() -> parseVariable(), "Expected member name");
                        memberBuilder.add(member.name);
                    }
                    atom = new NodeTerm.MemberAccess(var, new NodeVariable(String.join(".", memberBuilder)));
                }
            } 
            else {
                atom = parseLiteral();
            }

            if (atom == null)
                return null;

            return new NodeExpression.Term(atom);
        }
    }


    /*
     * Parse Variable
     * 
     * Variable names, here referred to as Qualifiers, can consist of any combination of alphanumeric characters or 
     * underscores but not start with a number. Variables refer to values that are provided later at run time under 
     * specific names. Variables can consist of two Qualifiers separated by a period.
     * 
     * Variable production is as follows:
     * 
     * Variable -> [Qualifier] | [Qualifier].[Qualifier] 
     */
    private NodeVariable parseVariable() {
        if (!peek().isAny(TokenType.Qualifier)) {
            return null;   
        }

        return new NodeVariable(consume().value);
    }

    /*
     * Parse Literal
     * 
     * Any free form values not bound to variables or the results of computation are called Literals.
     * Literals can be boolean such as true or false, numeric such as 10 or 17.07, strings such as "SMG" or 'Kyle' and 
     * dates such as 31/12/2023 (must be in the format dd/MM/yyy), and finally the empty keyword, which represents null
     * or 0. 
     */
    private static final SimpleDateFormat format = new SimpleDateFormat("dd/MM/yyyy"); 
    private NodeTerm.Literal<?> parseLiteral() {

        if (tryConsume(Token.Empty)) {
            return new NodeTerm.Literal<Void>(null);
        }

        if (peek().isAny(TokenType.BooleanLiteral)) 
            return new NodeTerm.Literal<Boolean>(consume().equals(Token.True));
        else if (peek().isAny(TokenType.StringLiteral))
            return new NodeTerm.Literal<String>(consume().value);
        else if (peek().isAny(TokenType.NumberLiteral)) {
            final String repr = consume().value;
            try {
                return new NodeTerm.Literal<Integer>(Integer.parseInt(repr));
            }
            catch (NumberFormatException e) {
                return new NodeTerm.Literal<Double>(Double.parseDouble(repr));
            }
        }
        else if (peek().isAny(TokenType.DateLiteral)) {
            final String dateToken = consume().value;
            try {
                return new NodeTerm.Literal<Date>(format.parse(dateToken));
            } catch (ParseException e) {
                throw error("Date format error: " + dateToken);
            }
        }
        else {
            return null;
        }
    }

    /***************************************************************************
     * Token Handling
     * 
     * The Parser periodically requests tokens from the Tokeniser as and when it 
     * needs them. This saves space and is at the minimum as efficient as generating 
     * all tokens first. It also helps in the case of an error, where all the tokens
     * after an incorrect syntax need not be generated and overall reduce compute time.
     * 
     * The Tokeniser keeps a cache of tokens for use and although it usually by 
     * default keeps one token in cache, it can generate more tokens on demand. 
     * These extra tokens would not need to be generated again on future 
     * invocations of peek() or consume(). 
     **************************************************************************/
    private Token peek() {
        return peek(0);
    }
    
    private Token consume() {
        final Token consumable = cache.size() > 0 ? cache.remove(0) : tokeniser.nextToken();
        lineNumber += consumable.value.length() - consumable.value.replace("\n", "").length();
        // lineNumber += consumable.value.chars().filter(i -> i == '\n').count(); // Possible improvement
        return consumable;
    }

    private Token peek(int offset) {
        while (cache.size() <= offset) {
            cache.add(tokeniser.nextToken());
        }
        return cache.get(offset);
    }

    private boolean tryConsume(Token token) {
        final boolean success;
        if ((success = peek() == token))
            consume();
        return success;
    }

    private void tryConsume(Token token, String message) {
        if (!tryConsume(token)) {
            throw error(message);
        }
    }

    public NodeScope getRoot() {
        return root;
    }
}

class NodeScope {
    public final List<NodeStatement> statements;

    NodeScope(List<NodeStatement> statements) {
        this.statements = List.copyOf(statements);
    }

    public String toString() {
        return "Scope: { " + String.join(", ", statements.stream().map(x -> x.toString()).collect(Collectors.toList())) + " }";
    }
}

abstract class NodeStatement {

    public int lineNumber = 0;
    static class If extends NodeStatement {
    
        public final NodeExpression expression;
        public final NodeScope success; 
        public final NodeScope fail;
    
        If(NodeExpression expression, NodeScope success, NodeScope fail) {
            this.expression = expression;
            this.success = success;
            this.fail = fail;
        }
        
        public Object host(Visitor visitor) { return visitor.visit(this); }

        public String toString() {
            return "StmtIf: if " + expression + " then " + success + " else " + fail + " }";
        } 
    }

    static class While extends NodeStatement {
    
        public final NodeExpression expression;
        public final NodeScope scope; 
    
        While(NodeExpression expression, NodeScope scope) {
            this.expression = expression;
            this.scope = scope;
        }
        
        public Object host(Visitor visitor) { return visitor.visit(this); }

        public String toString() {
            return "StmtWhile: while " + expression + " do " + scope + " }";
        } 
    }

    static class Scope extends NodeStatement {
    
        public final NodeScope scope; 
    
        Scope(NodeScope scope) {
            this.scope = scope;
        }
        
        public Object host(Visitor visitor) { return visitor.visit(this); }

        public String toString() {
            return "StmtScope: " + scope + " }";
        } 
    }

    static class Declare extends NodeStatement {
    
        public final NodeVariable qualifier;
        public final NodeExpression expression;
    
        Declare(NodeVariable qualifier, NodeExpression expression) {
            this.qualifier = qualifier;
            this.expression = expression;
        }
        
        public Object host(Visitor visitor) { return visitor.visit(this); }

        public String toString() {
            return "StmtDeclare: { " + qualifier + "=" + expression + " }";
        } 
    }
    
    static class Assign extends NodeStatement {
    
        public final NodeVariable qualifier;
        public final NodeExpression expression;
    
        Assign(NodeVariable qualifier, NodeExpression expression) {
            this.qualifier = qualifier;
            this.expression = expression;
        }
        
        public Object host(Visitor visitor) { return visitor.visit(this); }

        public String toString() {
            return "StmtAssign: { " + qualifier + "=" + expression + " }";
        } 
    }

    static class MemberAssign extends NodeStatement {
    
        public final NodeVariable qualifier;
        public final NodeVariable member;
        public final NodeExpression expression;
    
        MemberAssign(NodeVariable qualifier, NodeVariable member, NodeExpression expression) {
            this.qualifier = qualifier;
            this.expression = expression;
            this.member = member;
        }
        
        public Object host(Visitor visitor) { return visitor.visit(this); }

        public String toString() {
            return "MemberAssign: { " + qualifier + "=" + expression + " }";
        } 
    }


    static class Expression extends NodeStatement {
        public final NodeExpression expression;

        Expression(NodeExpression expression) {
            this.expression = expression;
        }
        
        public Object host(Visitor visitor) { return visitor.visit(this); }

        public String toString() {
            return "StmtExp: { " + expression + " }";
        } 
    }

    abstract Object host(Visitor visitor);
    interface Visitor {
        Object visit(Assign assignment);
        Object visit(MemberAssign assignment);
        Object visit(Declare declaration);
        Object visit(Expression expression);
        Object visit(If expression);
        Object visit(While expression);
        Object visit(Scope expression);
    }
}

enum BinaryOperator {                               // Precedence
    Exponent,                                       // 8
    Multiply, Divide, Modulo,                       // 7
    Add, Subtract,                                  // 6
    ShiftLeft, ShiftRight,                          // 5
    Less, LessEqual, Greater, GreaterEqual,         // 4
    Equal, NotEqual,                                // 3
    BitAnd, BitOr, BitXor,                          // 2
    And, Or                                         // 1
}

enum UnaryOperator {
    Increment, Decrement, Negate, Invert, Not
}

interface NodeExpression {
    class Binary implements NodeExpression {
        public final NodeExpression lhs, rhs;
        public final BinaryOperator op;
    
        Binary(BinaryOperator op, NodeExpression lhs, NodeExpression rhs) {
            this.op = op;
            this.lhs = lhs;
            this.rhs = rhs;
        }

        public <R> R host(Visitor<R> visitor) { return visitor.visit(this); }

        public String toString() {
            return "ExprBin: { " + lhs + " " + op + " " + rhs + " }";
        } 
    }
    
    class Unary implements NodeExpression {
        public final NodeExpression val;
        public final UnaryOperator op;
    
        Unary(UnaryOperator op, NodeExpression val) {
            this.op = op;
            this.val = val;
        }

        public <R> R host(Visitor<R> visitor) {
            return visitor.visit(this);
        }
        
        public String toString() {
            return "ExprUn: { " + op + " " + val + " }";
        } 
    }
    
    class Term implements NodeExpression {
        public final NodeTerm val;
    
        Term(NodeTerm val) {
            this.val = val;
        }

        public <R> R host(Visitor<R> visitor) {
            return visitor.visit(this);
        }
        
        public String toString() {
            return "ExprTerm: { " + val + " }";
        } 
    }

    <R> R host(Visitor<R> visitor);
    interface Visitor<R> {
        R visit(Binary node);
        R visit(Unary node);
        R visit(Term node);
    }
}

interface NodeTerm {
    class Literal<R> implements NodeTerm {
        public final R lit;
        public Literal(R lit) {
            this.lit = lit;
        }
        
        public Object host(Visitor visitor) {
            return visitor.visit(this);
        }

        public String toString() {
            return "TermLit: { " + lit + " }";
        } 
    }
    
    class Variable implements NodeTerm {
        public final NodeVariable var;
        public Variable(NodeVariable var) {
            this.var = var;
        }
        
        @Override
        public Object host(Visitor visitor) {
            return visitor.visit(this);
        }
        
        public String toString() {
            return "TermVar: { " + var + " }";
        } 
    }
    
    class MemberAccess implements NodeTerm {
        public final NodeVariable object;
        public final NodeVariable member;
        public MemberAccess(NodeVariable object, NodeVariable member) {
            this.object = object;
            this.member = member;
        }
        
        @Override
        public Object host(Visitor visitor) {
            return visitor.visit(this);
        }
        
        public String toString() {
            return "TermMAccess: { " + object + " }";
        } 
    }


    Object host(Visitor term);
    interface Visitor {
        Object visit(Literal<?> lit);
        Object visit(Variable var);
        Object visit(MemberAccess var);
    }
}

class NodeVariable {
    public final String name;
    NodeVariable(String name) {
        this.name = name;
    }

    public String toString() {
        return "Var: " + this.name;
    }
}
