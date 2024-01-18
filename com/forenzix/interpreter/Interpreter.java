package com.forenzix.interpreter;

import java.math.BigDecimal;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.Map;
import java.util.Optional;
import java.util.function.Supplier;

public class Interpreter {

    /**
     * Parser for a small custom language.
     * 
     * The process is split into three interleaving parts, namely:
     *   - Tokeniser
     *   - Parser
     *   - Interpreter
     * 
     * Although the Tokeniser is independent from the Parser, it works with the
     * latter in tandem to minimise space and resources used.
     */


    private final LinkedList<Map<String, Object>> scopes;
    private final Parser parser;
    private int lineNumber = 0;
    private Object lastResult;
    private MemberAccessor<Object, String, Object> memberAccessCallback;
    private MemberUpdater<Object, String, Object> memberUpdateCallback;

    public Interpreter(String input) {
        this(input, new HashMap<>());
    }
    
    public Interpreter(String input, Map<String, Object> variables) {
        this.parser = new Parser(input);
        this.scopes = new LinkedList<>();
        scopes.add(new HashMap<>(variables));
    }

    /**
     * Variable management.
     * 
     * Crucial for the execution of the interpreter when variables are involved.
     */
    public Interpreter addVariable(String key, Object value) {
        findVariable(key).orElse(scopes.getLast()).put(key, value);
        return this;
    }

    public Map<String, Object> getGlobalScopeVariables() {
        return scopes.getFirst();
    }

    public Object getVariable(String key) {
        return findVariable(key).orElseThrow(() -> error("Variable " + key + " is undefined")).get(key);
    }

    public boolean defined(String key) {
        return findVariable(key).isPresent();
    }

    public Optional<Map<String, Object>> findVariable(String key) {
        Iterator<Map<String, Object>> itr = scopes.descendingIterator();
        while (itr.hasNext()) {
            final Map<String, Object> scope = itr.next();
            if (scope.containsKey(key)) {
                return Optional.of(scope);
            }
        }

        return Optional.empty();
    }

    public Object getLastResult() {
        return lastResult;
    }

    private void enterScope() {
        scopes.add(new HashMap<>());
    }

    private void exitScope() {
        scopes.removeLast();
    }

    public Interpreter clearVariables() {
        scopes.clear();
        return this;
    }

    public String getTree() {
        return parser.getRoot() != null ? parser.getRoot().toString() : null;
    }

    public void setMemberAccessCallback(MemberAccessor<Object, String, Object> callback) {
        memberAccessCallback = callback;
    }

    public void setMemberUpdateCallback(MemberUpdater<Object, String, Object> callback) {
        memberUpdateCallback = callback;
    }

    private RuntimeException error(String message) {
        return new RuntimeException(message + " (line: " + lineNumber + ")");
    }
    
    /***************************************************************************
     * Interpreter
     * 
     * Traverses the Syntax Tree in in-order fashion. Calls the parser if no root
     * node can be found.
     **************************************************************************/
    public Object interpret() {
        if (parser.getRoot() == null)
            parser.parse();
        return lastResult = interpretGlobalScope(parser.getRoot());
    }
    
    public Object interpretGlobalScope(NodeScope scope) {
        return interpretScope(scope, true);
    }

    public Object interpretScope(NodeScope scope) {
        return interpretScope(scope, false);
    }

    public Object interpretScope(NodeScope scope, boolean global) {

        if (!global) enterScope();
        Object output = null;
        for (NodeStatement statement : scope.statements) {
            output = interpretStatement(statement);
        }
        if (!global) exitScope();

        return output;
    }

    private final NodeStatement.Visitor statementVisitor = new NodeStatement.Visitor() {
        public Object visit(NodeStatement.Assign assignment) {
            // Uncomment this if you do not want to allow users to use undeclared variables
            // However, it doesn't make much sense to have this on without typing.
            // if (!defined(assignment.qualifier.name)) {
            //     throw error("Assignment to an undefined variable: " + assignment.qualifier.name);
            // }

            final Object value = interpretExpression(assignment.expression);
            addVariable(assignment.qualifier.name, value);
            return value;
        }

        public Object visit(NodeStatement.MemberAssign assignment) {
            
            if (memberUpdateCallback == null) {
                throw error("Member updater callback not defined.");
            }

            final Object value = interpretExpression(assignment.expression); 
            memberUpdateCallback.consume((getVariable(assignment.qualifier.name)), assignment.member.name, value);
            return null;
        }

        public Object visit(NodeStatement.Declare declaration) {

            if (defined(declaration.qualifier.name)) {
                throw error("Redefining an existing variable: " + declaration.qualifier.name);
            }
            
            final Object value = interpretExpression(declaration.expression);
            addVariable(declaration.qualifier.name, value);
            return value;
        }

        public Object visit(NodeStatement.Expression expression) {
            return interpretExpression(expression.expression);
        }

        public Object visit(NodeStatement.If ifStmt) {
            if ((Boolean) interpretExpression(ifStmt.expression)) {
                return interpretScope(ifStmt.success);
            }
            else if (ifStmt.fail != null) {
                return interpretScope(ifStmt.fail);
            }
            else {
                return null;
            }
        }
        
        public Object visit(NodeStatement.While whileStmt) {
            while ((Boolean) interpretExpression(whileStmt.expression)) {
                interpretScope(whileStmt.scope);
            }
            return null;
        }
        
        public Object visit(NodeStatement.Scope scope) {
            return interpretScope(scope.scope);
        }
    };

    public Object interpretStatement(NodeStatement statement) {
        lineNumber = statement.lineNumber;
        return statement.host(statementVisitor);
    }

    /**
     * Interpret Boolean Expressrion node.
     * 
     * Since all nodes directly, or indirectly extend BExpr, they get redirected here to their respective interpreters.
     * This interpreter also benefits from Java's inherent evaluation system, where if an expression such as 'true or x'
     * which would normally produce an error if 'x' cannot be evaluated as a boolean would actually evaluate to true.
     * This can be beneficial but potentially hard to debug.
     */

    final NodeExpression.Visitor<Object> nodeExpressionVisitor = new NodeExpression.Visitor<Object>() {
        @Override
        public Object visit(NodeExpression.Binary node) {

            final Object left = interpretExpression(node.lhs);
            // final Object right = interpretExpression(node.rhs);
            final Object right;

            // Special case: String concatenation
            if (left instanceof String && node.op == BinaryOperator.Add) {
                return ((String) left).concat(stringValue(right = interpretExpression(node.rhs)));
            }
            
            // Special case: String formatting
            if (left instanceof String && node.op == BinaryOperator.Modulo) {
                return String.format((String) left, interpretExpression(node.rhs));
            }

            // Special case: Null check
            if (node.op == BinaryOperator.Equal) {
                right = interpretExpression(node.rhs);

                if (left == null || right == null) {
                    return left == right;
                }
            }
            
            else if (node.op == BinaryOperator.NotEqual) {
                right = interpretExpression(node.rhs);

                if (left == null || right == null) {
                    return left != right;
                }
            }

            else {
                right = null;
            }

            final Supplier<Object> rGetter = right != null ? () -> right : () -> interpretExpression(node.rhs);

            switch (node.op) {
                case Exponent:          return Math.pow(evaluate(left), evaluate(rGetter.get()));
                case Multiply:          return evaluate(left) * evaluate(rGetter.get());
                case Divide:            return evaluate(left) / evaluate(rGetter.get());
                case Modulo:            return evaluate(left) % evaluate(rGetter.get());
                case Add:               return evaluate(left) + evaluate(rGetter.get());
                case Subtract:          return evaluate(left) - evaluate(rGetter.get());
                case Greater:           return evaluate(left) > evaluate(rGetter.get());
                case GreaterEqual:      return evaluate(left) >= evaluate(rGetter.get());
                case Less:              return evaluate(left) < evaluate(rGetter.get());
                case LessEqual:         return evaluate(left) <= evaluate(rGetter.get());
                case NotEqual:          return evaluate(left) != evaluate(rGetter.get());
                case Equal:             return evaluate(left) == evaluate(rGetter.get());
                case BitAnd:            return (Integer) left & (Integer) rGetter.get();
                case BitOr:             return (Integer) left | (Integer) rGetter.get();
                case BitXor:            return (Integer) left ^ (Integer) rGetter.get();
                case ShiftLeft:         return (Integer) left << (Integer) rGetter.get();
                case ShiftRight:        return (Integer) left >> (Integer) rGetter.get();
                case And:               return (Boolean) left && (Boolean) rGetter.get();
                case Or:                return (Boolean) left || (Boolean) rGetter.get();
                default:
                    throw error("Unsupported operation: " + node.op);
            }
        }

        @Override
        public Object visit(NodeExpression.Unary node) {
            switch (node.op) {
                case Not:               return ! (Boolean) interpretExpression(node.val);
                case Invert:            return ~ (Integer) interpretExpression(node.val);
                case Negate:            return - evaluate(interpretExpression(node.val));
                case Decrement:
                case Increment:
                default:
                    throw error("Unsupported operation: " + node.op);

            }
        }
        @Override
        public Object visit(NodeExpression.Term node) {
            return interpretTerm(node.val);
        }
    };

    private static final SimpleDateFormat format = new SimpleDateFormat("dd/MM/yyyy"); 
    private String stringValue(Object o) {

        if (o instanceof Date) {
            return format.format(o);
        }

        else if (o instanceof Double) {
            return NumberFormat.getInstance().format((Double) o);
        }

        return String.valueOf(o);
    }

    private Object interpretExpression(NodeExpression node) {
        return node.host(nodeExpressionVisitor);
    }

    /**
     * Interpret Comparision node.
     * 
     * Comparision nodes compare values on both sides of an operator. For a value to be comparable it must first be 
     * represented as a double.
     */
    
    /**
     * Interpret Term
     * 
     * Term is an atomic node and is the terminal node in any syntax tree branch. Terms can be literal values or variables
     * that can be provided on Parser creation. Terms can be any value of any type, so long as they fit in higher level 
     * expressions. Qualifier-member nodes will be concatenated with a period "." when grabbed fromn the variable map
     */

    final NodeTerm.Visitor termVisitor = new NodeTerm.Visitor() {
        public Object visit(NodeTerm.Literal<?> literal) {
            return literal.lit;
        }

        public Object visit(NodeTerm.Variable variable) {
            return getVariable(variable.var.name);
        }
        
        public Object visit(NodeTerm.MemberAccess maccess) {
            if (memberAccessCallback == null) {
                throw error("Member access callback not defined.");
            }

            return memberAccessCallback.apply(getVariable(maccess.object.name), maccess.member.name);
        }
    };

    private Object interpretTerm(NodeTerm term) { return term.host(termVisitor); }
    
    /**
     * To Double.
     * 
     * For a node to viably exist in an equality, or an arithmetic expression, it must have a numerical representation.
     * For this reason, all values are converted and cast into double. Strings are hashed before being evaluated. While
     * it can lead to bizzare results with most inequality operations, it is "good enough" testing if two strings are the 
     * same. All strings that are equal must have the same hash but not all strings with the same has must be equivalent.
     * Read the Java Documentation on Strings for more info.
     */
    private double evaluate(Object value) {
        if (value == null) return 0;
        else if (value instanceof Date) return ((Date) value).getTime();
        else if (value instanceof Double) return (Double) value;
        else if (value instanceof BigDecimal) return ((BigDecimal) value).doubleValue();
        else if (value instanceof Integer) return (Integer) value;
        else if (value instanceof String) return ((String) value).hashCode();
        else if (value instanceof Long) return (Long) value;
        
        throw error("Atomic expression required to be integer, or integer similar, but is not: " + value);
    }

    @FunctionalInterface
    public interface MemberAccessor<S, M, R> {
        public R apply(S source, M member);
    }
    
    @FunctionalInterface
    public interface MemberUpdater<S, M, R> {
        public void consume(S source, M member, R rvalue);
    }
}

