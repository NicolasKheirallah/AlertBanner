"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.forEachChild = exports.StatementParser = exports.SyntaxKind = void 0;
const tokenizer_1 = require("./tokenizer");
var SyntaxKind;
(function (SyntaxKind) {
    SyntaxKind[SyntaxKind["ASSIGNMENT"] = 0] = "ASSIGNMENT";
    SyntaxKind[SyntaxKind["BINARY_OPERATOR"] = 1] = "BINARY_OPERATOR";
    SyntaxKind[SyntaxKind["CATCH_STATEMENT"] = 2] = "CATCH_STATEMENT";
    SyntaxKind[SyntaxKind["DO_STATEMENT"] = 3] = "DO_STATEMENT";
    SyntaxKind[SyntaxKind["FOR_STATEMENT"] = 4] = "FOR_STATEMENT";
    SyntaxKind[SyntaxKind["IDENTIFIER"] = 5] = "IDENTIFIER";
    SyntaxKind[SyntaxKind["IF_STATEMENT"] = 6] = "IF_STATEMENT";
    SyntaxKind[SyntaxKind["NUMERIC_LITERAL"] = 7] = "NUMERIC_LITERAL";
    SyntaxKind[SyntaxKind["POST_CONDITION"] = 8] = "POST_CONDITION";
    SyntaxKind[SyntaxKind["QUIT_STATEMENT"] = 9] = "QUIT_STATEMENT";
    SyntaxKind[SyntaxKind["RETURN_STATEMENT"] = 10] = "RETURN_STATEMENT";
    SyntaxKind[SyntaxKind["SET_STATEMENT"] = 11] = "SET_STATEMENT";
    SyntaxKind[SyntaxKind["MULTIPLE_VARIABLE_SET"] = 12] = "MULTIPLE_VARIABLE_SET";
    SyntaxKind[SyntaxKind["STRING_LITERAL"] = 13] = "STRING_LITERAL";
    SyntaxKind[SyntaxKind["WHILE_STATEMENT"] = 14] = "WHILE_STATEMENT";
    SyntaxKind[SyntaxKind["TYPE_STATEMENT"] = 15] = "TYPE_STATEMENT";
    SyntaxKind[SyntaxKind["VARIABLE_DECLARATION"] = 16] = "VARIABLE_DECLARATION";
    SyntaxKind[SyntaxKind["TYPE_IDENTIFIER"] = 17] = "TYPE_IDENTIFIER";
})(SyntaxKind = exports.SyntaxKind || (exports.SyntaxKind = {}));
var OPERATOR_VALUE;
(function (OPERATOR_VALUE) {
    OPERATOR_VALUE["AND_LITERAL"] = "and";
    OPERATOR_VALUE["APOSTROPHE"] = "'";
    OPERATOR_VALUE["AT"] = "@";
    OPERATOR_VALUE["BACK_SLASH"] = "\\";
    OPERATOR_VALUE["CARROT"] = "^";
    OPERATOR_VALUE["COLON"] = ":";
    OPERATOR_VALUE["DOLLAR"] = "$";
    OPERATOR_VALUE["DOT"] = ".";
    OPERATOR_VALUE["EQUAL"] = "=";
    OPERATOR_VALUE["EXCLAMATION"] = "!";
    OPERATOR_VALUE["GREATER_THAN"] = ">";
    OPERATOR_VALUE["HASH"] = "#";
    OPERATOR_VALUE["LEFT_BRACKET"] = "[";
    OPERATOR_VALUE["LESS_THAN"] = "<";
    OPERATOR_VALUE["MINUS"] = "-";
    OPERATOR_VALUE["NOT_LITERAL"] = "not";
    OPERATOR_VALUE["OR_LITERAL"] = "or";
    OPERATOR_VALUE["PLUS"] = "+";
    OPERATOR_VALUE["QUESTION_MARK"] = "?";
    OPERATOR_VALUE["RIGHT_BRACKET"] = "]";
    OPERATOR_VALUE["SLASH"] = "/";
    OPERATOR_VALUE["STAR"] = "*";
    OPERATOR_VALUE["UNDERSCORE"] = "_";
    OPERATOR_VALUE["RET"] = "ret";
})(OPERATOR_VALUE || (OPERATOR_VALUE = {}));
var STORAGE_MODIFIERS;
(function (STORAGE_MODIFIERS) {
    STORAGE_MODIFIERS["STATIC"] = "static";
    STORAGE_MODIFIERS["NEW"] = "new";
    STORAGE_MODIFIERS["LITERAL"] = "literal";
})(STORAGE_MODIFIERS || (STORAGE_MODIFIERS = {}));
var ACCESS_MODIFIERS;
(function (ACCESS_MODIFIERS) {
    ACCESS_MODIFIERS["PUBLIC"] = "public";
    ACCESS_MODIFIERS["PRIVATE"] = "private";
})(ACCESS_MODIFIERS || (ACCESS_MODIFIERS = {}));
var STATEMENT_KEYWORD;
(function (STATEMENT_KEYWORD) {
    STATEMENT_KEYWORD["DO"] = "do";
    STATEMENT_KEYWORD["SET"] = "set";
    STATEMENT_KEYWORD["IF"] = "if";
    STATEMENT_KEYWORD["CATCH"] = "catch";
    STATEMENT_KEYWORD["FOR"] = "for";
    STATEMENT_KEYWORD["QUIT"] = "quit";
    STATEMENT_KEYWORD["RETURN"] = "return";
    STATEMENT_KEYWORD["WHILE"] = "while";
    STATEMENT_KEYWORD["TYPE"] = "type";
})(STATEMENT_KEYWORD || (STATEMENT_KEYWORD = {}));
const UNARY_OPERATORS = [
    { value: OPERATOR_VALUE.APOSTROPHE },
    { value: OPERATOR_VALUE.AT },
    { value: OPERATOR_VALUE.CARROT },
    { value: OPERATOR_VALUE.DOLLAR, appendable: true },
    { value: OPERATOR_VALUE.DOT },
    { value: OPERATOR_VALUE.MINUS },
    { value: OPERATOR_VALUE.NOT_LITERAL },
    { value: OPERATOR_VALUE.PLUS },
    { value: OPERATOR_VALUE.RIGHT_BRACKET },
    { value: OPERATOR_VALUE.RET },
];
const BINARY_OPERATORS = [
    { value: OPERATOR_VALUE.AND_LITERAL },
    { value: OPERATOR_VALUE.APOSTROPHE, appendable: true },
    { value: OPERATOR_VALUE.AT },
    { value: OPERATOR_VALUE.BACK_SLASH },
    { value: OPERATOR_VALUE.CARROT },
    { value: OPERATOR_VALUE.DOT },
    { value: OPERATOR_VALUE.EQUAL },
    { value: OPERATOR_VALUE.EXCLAMATION },
    { value: OPERATOR_VALUE.GREATER_THAN, appendable: true },
    { value: OPERATOR_VALUE.HASH },
    { value: OPERATOR_VALUE.LEFT_BRACKET },
    { value: OPERATOR_VALUE.LESS_THAN, appendable: true },
    { value: OPERATOR_VALUE.MINUS },
    { value: OPERATOR_VALUE.NOT_LITERAL, appendable: true },
    { value: OPERATOR_VALUE.OR_LITERAL },
    { value: OPERATOR_VALUE.PLUS },
    { value: OPERATOR_VALUE.QUESTION_MARK },
    { value: OPERATOR_VALUE.SLASH },
    { value: OPERATOR_VALUE.STAR },
    { value: OPERATOR_VALUE.UNDERSCORE },
];
class StatementParser {
    constructor(arg) {
        this.tokens = [];
        if (typeof arg === 'string') {
            this.tokenizer = (0, tokenizer_1.getTokens)(arg);
        }
        else if (Array.isArray(arg)) {
            this.tokenizer = arg[Symbol.iterator]();
        }
        else {
            this.tokenizer = arg;
        }
        this.next(); // should I?
    }
    parseLine() {
        if (!this.activeToken)
            return [];
        const statements = [];
        do {
            if (this.activeToken.isNewLine())
                break;
            if (this.activeToken.isAlphanumeric()) {
                const statement = this.parseStatement();
                if (!statement) {
                    this.next();
                    continue;
                }
                statements.push(statement);
            }
            else if (this.activeToken.isWhiteSpace())
                this.next(true);
            else
                break;
        } while (this.activeToken);
        return statements;
    }
    parseStatement() {
        if (!this.activeToken)
            return;
        if (!this.activeToken.isAlphanumeric())
            return;
        const action = this.activeToken;
        let loadFunction;
        let kind;
        const loadSingleExpression = () => {
            if (!this.next(true))
                return { action, kind, expressions: [] };
            const expression = loadFunction();
            const expressions = expression ? [expression] : [];
            return { kind, action, expressions };
        };
        const loadCommaSeparatedExpressions = () => {
            if (!this.next(true))
                return { action, kind, expressions: [] };
            const expressions = this.loadCommaSeparated(loadFunction);
            return { kind, action, expressions };
        };
        switch (action.value) {
            case STATEMENT_KEYWORD.DO:
                loadFunction = () => this.parseExpression();
                kind = SyntaxKind.DO_STATEMENT;
                return loadCommaSeparatedExpressions();
            case STATEMENT_KEYWORD.SET:
                loadFunction = () => this.parseSetExpression();
                kind = SyntaxKind.SET_STATEMENT;
                return loadCommaSeparatedExpressions();
            case STATEMENT_KEYWORD.IF:
                loadFunction = () => this.parseExpression();
                kind = SyntaxKind.IF_STATEMENT;
                return loadCommaSeparatedExpressions();
            case STATEMENT_KEYWORD.CATCH:
                loadFunction = () => this.parseExpression();
                kind = SyntaxKind.CATCH_STATEMENT;
                return loadCommaSeparatedExpressions();
            case STATEMENT_KEYWORD.FOR:
                return this.parseForStatement();
            case STATEMENT_KEYWORD.QUIT:
                loadFunction = () => this.parseExpression();
                kind = SyntaxKind.QUIT_STATEMENT;
                return loadCommaSeparatedExpressions();
            case STATEMENT_KEYWORD.RETURN:
                loadFunction = () => this.parseExpression();
                kind = SyntaxKind.RETURN_STATEMENT;
                return loadSingleExpression();
            case STATEMENT_KEYWORD.WHILE:
                loadFunction = () => this.parseExpression();
                kind = SyntaxKind.WHILE_STATEMENT;
                return loadSingleExpression();
            case STATEMENT_KEYWORD.TYPE:
                return this.parseTypeStatement();
            default:
                return;
        }
    }
    parseTypeStatement() {
        const action = this.activeToken;
        this.next(true);
        let staticToken;
        let newToken;
        let publicToken;
        let literalToken;
        const getKeyWordToken = (token) => {
            switch (token.value) {
                case ACCESS_MODIFIERS.PUBLIC:
                    publicToken = token;
                    return true;
                case STORAGE_MODIFIERS.LITERAL:
                    literalToken = token;
                    return true;
                case STORAGE_MODIFIERS.NEW:
                    newToken = token;
                    return true;
                case STORAGE_MODIFIERS.STATIC:
                    staticToken = token;
                    return true;
                default:
                    return false;
            }
        };
        const parseTypeAssignments = () => {
            if (!this.activeToken || !this.activeToken.isAlphanumeric()) {
                if (literalToken || newToken || publicToken || staticToken) {
                    const emptyDeclaration = {
                        id: undefined,
                        kind: SyntaxKind.VARIABLE_DECLARATION,
                        literalToken: undefined,
                        newToken,
                        publicToken,
                        staticToken,
                        type: undefined,
                    };
                    return [emptyDeclaration];
                }
                return [];
            }
            const type = { id: this.activeToken, kind: SyntaxKind.TYPE_IDENTIFIER };
            if (!this.next(true) || staticToken) {
                const declaration = {
                    id: undefined,
                    kind: SyntaxKind.VARIABLE_DECLARATION,
                    literalToken,
                    newToken,
                    publicToken,
                    staticToken,
                    type,
                };
                return [declaration];
            }
            const assignments = this.loadCommaSeparated(() => {
                return this.parseAssignment(() => {
                    const variable = this.parseValue(); // why not parseIdentifier
                    return {
                        args: variable.args,
                        id: variable.id,
                        kind: SyntaxKind.VARIABLE_DECLARATION,
                        literalToken,
                        newToken,
                        publicToken,
                        staticToken,
                        type,
                    };
                });
            });
            assignments.forEach(expression => {
                forEachChild(expression, node => {
                    if (!node)
                        return;
                    if (node.kind === SyntaxKind.VARIABLE_DECLARATION) {
                        const declaration = node;
                        if (declaration.args) {
                            declaration.args = declaration.args.map((arg) => {
                                if (!arg)
                                    return;
                                arg.kind = SyntaxKind.TYPE_IDENTIFIER;
                                return arg;
                            });
                        }
                    }
                    return true;
                });
            });
            return assignments;
        };
        while (this.activeToken && getKeyWordToken(this.activeToken)) {
            if (!this.next(true))
                break;
        }
        const expressions = parseTypeAssignments();
        return {
            action,
            expressions,
            kind: SyntaxKind.TYPE_STATEMENT,
        };
    }
    parseAssignment(getLeft) {
        const left = getLeft();
        let rootNode = left;
        if (this.activeToken && this.activeToken.isEqualSign()) {
            const equalSign = this.activeToken;
            this.next(true);
            const expression = this.parseExpression();
            rootNode = { operator: [equalSign], kind: SyntaxKind.ASSIGNMENT };
            rootNode.left = left;
            rootNode.right = expression;
        }
        return rootNode;
    }
    parseForStatement() {
        if (!this.activeToken)
            return;
        const action = this.activeToken;
        const forStatement = { action, kind: SyntaxKind.FOR_STATEMENT, expressions: [] };
        if (!this.next())
            return forStatement; // consume for
        if (!this.next())
            return forStatement; // consume first space
        const spaceOrExpression = this.activeToken;
        if (spaceOrExpression.isSpace()) {
            this.next();
            return forStatement; // argumentless for
        }
        const expression = this.parseExpression();
        if (expression)
            forStatement.expressions.push(expression);
        return forStatement;
    }
    parseSetExpression() {
        if (!this.activeToken)
            return;
        const postCondition = this.parsePostCondition();
        const assignment = this.parseAssignment(() => {
            const setVariables = this.parseSetVariables();
            if (this.activeToken && this.activeToken.isWhiteSpace())
                this.next(true);
            if (postCondition && !setVariables) {
                postCondition.expression = setVariables;
                return postCondition;
            }
            return setVariables;
        });
        if (assignment && postCondition) {
            postCondition.expression = assignment;
            return postCondition;
        }
        return assignment;
    }
    parsePostCondition() {
        if (!this.activeToken)
            return;
        if (this.activeToken.isColon()) {
            const colon = this.activeToken;
            const postCondition = { kind: SyntaxKind.POST_CONDITION, colon };
            this.next(true);
            const condition = this.parseExpression();
            if (!condition)
                return postCondition;
            postCondition.condition = condition;
            return postCondition;
        }
    }
    parseSetVariables() {
        if (!this.activeToken)
            return;
        if (this.activeToken.isOpenParen()) {
            this.next(true);
            const variables = this.loadCommaSeparated(() => this.parseExpression());
            if (this.activeToken && this.activeToken.isCloseParen())
                this.next(true);
            return { variables, kind: SyntaxKind.MULTIPLE_VARIABLE_SET };
        }
        else {
            return this.parseExpression(true);
        }
    }
    parseExpression(ignoreEquals, includeRet) {
        const postCondition = this.parsePostCondition();
        let rootNode = this.parseValue(undefined, includeRet);
        if (!rootNode) {
            if (postCondition)
                return postCondition;
            else
                return;
        }
        if (this.activeToken && this.activeToken.isEqualSign() && ignoreEquals) {
            return rootNode;
        }
        rootNode = this.parseOperatorSeparatedValues(rootNode, ignoreEquals);
        if (!rootNode)
            return;
        rootNode = this.parseColonSeparatedValues(rootNode);
        if (postCondition) {
            postCondition.expression = rootNode;
            rootNode = postCondition;
        }
        return rootNode;
    }
    parseColonSeparatedValues(rootNode) {
        while (this.activeToken && this.activeToken.isColon()) {
            const colonToken = this.activeToken;
            this.next(true);
            const colon = {
                kind: SyntaxKind.BINARY_OPERATOR,
                left: rootNode,
                operator: [colonToken],
                right: this.parseValue(),
            };
            rootNode = colon;
        }
        return rootNode;
    }
    parseOperatorSeparatedValues(rootNode, ignoreEquals) {
        while (this.activeToken && getBinaryOperator(this.activeToken.value)) {
            if (this.activeToken.isEqualSign() && ignoreEquals)
                break;
            const operator = this.parseBinaryOperator();
            if (!operator)
                return;
            operator.left = rootNode;
            operator.right = this.parseValue();
            rootNode = operator;
        }
        return rootNode;
    }
    parseBinaryOperator() {
        if (!this.activeToken)
            return;
        const binaryOperator = {
            kind: SyntaxKind.BINARY_OPERATOR,
            operator: [this.activeToken],
        };
        if (!this.next(true))
            return binaryOperator;
        let operator;
        do {
            operator = getBinaryOperator(this.activeToken.value);
            if (!operator)
                break;
            if (operator) {
                const not = operator.value === OPERATOR_VALUE.NOT_LITERAL
                    || operator.value === OPERATOR_VALUE.APOSTROPHE;
                if (not && binaryOperator.operator.length)
                    break;
            }
            binaryOperator.operator.push(this.activeToken);
            this.next(true);
        } while (operator && operator.appendable);
        return binaryOperator;
    }
    parseUnaryOperator(includeRet) {
        if (!this.activeToken)
            return [];
        const leftOperator = [];
        let unaryOperator = getUnaryOperator(this.activeToken.value, includeRet);
        if (unaryOperator) {
            leftOperator.push(this.activeToken);
            this.next(true);
            while (unaryOperator && unaryOperator.appendable) {
                unaryOperator = getUnaryOperator(this.activeToken.value);
                if (unaryOperator) {
                    leftOperator.push(this.activeToken);
                    this.next(true);
                }
            }
        }
        return leftOperator;
    }
    parseValue(tree, includeRet) {
        let value;
        if (!this.activeToken || this.activeToken.isWhiteSpace()) {
            if (tree)
                return tree;
            else
                return;
        }
        const unaryOperator = this.parseUnaryOperator(includeRet);
        if (!this.activeToken) {
            return {
                id: new tokenizer_1.Token(-1 /* Type.Undefined */, '', { character: 0, line: 0 }),
                kind: SyntaxKind.IDENTIFIER,
                unaryOperator,
            };
        }
        if (this.activeToken.type === 1 /* Type.Alphanumeric */) {
            value = this.parseIdentifier();
            if (!value)
                return;
            if (unaryOperator.length)
                value.unaryOperator = unaryOperator;
        }
        else if (this.activeToken.type === 9 /* Type.DoubleQuotes */) {
            value = this.parseStringLiteral();
            if (!value)
                return;
            if (unaryOperator.length)
                value.unaryOperator = unaryOperator;
        }
        else if (this.activeToken.type === 2 /* Type.Numeric */) {
            value = { id: this.activeToken, kind: SyntaxKind.NUMERIC_LITERAL };
            if (unaryOperator.length)
                value.unaryOperator = unaryOperator;
            this.next(true);
        }
        else if (this.activeToken.type === 40 /* Type.OpenParen */) {
            this.next(true);
            value = this.parseExpression();
            this.next(true);
        }
        if (tree && value) {
            tree.right = value;
            value = tree;
        }
        if (this.activeToken && (this.activeToken.type === 46 /* Type.Period */ || this.activeToken.type === 94 /* Type.Caret */)) {
            const operator = {
                kind: SyntaxKind.BINARY_OPERATOR,
                left: value,
                operator: [this.activeToken],
            };
            this.next();
            return this.parseValue(operator);
        }
        if (value && this.activeToken && this.activeToken.isWhiteSpace())
            this.next(true);
        return value;
    }
    parseIdentifier() {
        if (!this.activeToken)
            return;
        const id = this.activeToken;
        if (this.next() && this.activeToken.isOpenParen()) {
            const openParen = this.activeToken;
            if (this.next(true)) {
                const args = this.parseArgs();
                if (this.activeToken.isCloseParen()) {
                    const closeParen = this.activeToken;
                    this.next();
                    return { id, kind: SyntaxKind.IDENTIFIER, args, openParen, closeParen };
                }
            }
        }
        return { id, kind: SyntaxKind.IDENTIFIER };
    }
    parseStringLiteral() {
        if (!this.activeToken)
            return;
        const openQuote = this.activeToken;
        this.next();
        const id = this.activeToken;
        if (!id || !this.next())
            return;
        if (this.activeToken.isDoubleQuotes()) {
            const closeQuote = this.activeToken;
            this.next(true);
            return { id, kind: SyntaxKind.STRING_LITERAL, openQuote, closeQuote };
        }
    }
    parseArgs() {
        return this.loadCommaSeparated(() => this.parseExpression(false, true));
    }
    loadCommaSeparated(getArg) {
        const args = [];
        do {
            if (!this.activeToken)
                break;
            if (this.activeToken.isWhiteSpace()) {
                if (!this.next(true))
                    break;
            }
            if (this.activeToken.isComma()) {
                args.push(undefined); // desired behavior?
                continue;
            }
            const arg = getArg();
            if (arg && !Array.isArray(arg))
                args.push(arg);
            else if (arg)
                args.push(...arg);
            else if (!arg && args.length > 0)
                args.push(undefined);
            else
                break;
            if (!this.activeToken)
                break;
        } while (this.activeToken.isComma() && this.next(true));
        return args;
    }
    next(skipSpaceOrTab) {
        if (this.activeToken)
            this.previousToken = this.activeToken;
        do {
            this.activeToken = this.tokenizer.next().value;
            if (this.activeToken)
                this.tokens.push(this.activeToken);
        } while (skipSpaceOrTab && this.activeToken && (this.activeToken.isSpace() || this.activeToken.isTab()));
        return this.activeToken !== undefined;
    }
}
exports.StatementParser = StatementParser;
function getBinaryOperator(tokenValue) {
    return BINARY_OPERATORS.find(o => o.value === tokenValue);
}
function getUnaryOperator(tokenValue, includeRet) {
    const operator = UNARY_OPERATORS.find(o => o.value === tokenValue);
    if (!operator)
        return;
    if (operator.value === OPERATOR_VALUE.RET && !includeRet)
        return;
    return operator;
}
function forEachChild(node, f) {
    let goDeeper = false;
    if (!node)
        return;
    switch (node.kind) {
        case SyntaxKind.DO_STATEMENT:
        case SyntaxKind.IF_STATEMENT:
        case SyntaxKind.QUIT_STATEMENT:
        case SyntaxKind.RETURN_STATEMENT:
        case SyntaxKind.SET_STATEMENT:
        case SyntaxKind.WHILE_STATEMENT:
        case SyntaxKind.FOR_STATEMENT:
        case SyntaxKind.CATCH_STATEMENT:
        case SyntaxKind.TYPE_STATEMENT:
            goDeeper = f(node);
            if (!goDeeper)
                return;
            const statement = node;
            statement.expressions.forEach(expression => {
                if (!expression)
                    return;
                forEachChild(expression, f);
            });
            break;
        case SyntaxKind.ASSIGNMENT:
        case SyntaxKind.BINARY_OPERATOR:
            goDeeper = f(node);
            if (!goDeeper)
                return;
            const assignment = node;
            const left = assignment.left;
            if (left && left.kind === SyntaxKind.MULTIPLE_VARIABLE_SET) {
                left.variables.forEach(n => {
                    forEachChild(n, f);
                });
            }
            else if (left) {
                forEachChild(left, f);
            }
            const right = assignment.right;
            if (right) {
                forEachChild(right, f);
            }
            break;
        case SyntaxKind.POST_CONDITION:
            goDeeper = f(node);
            if (!goDeeper)
                return;
            const postCondition = node;
            if (postCondition.condition)
                forEachChild(postCondition.condition, f);
            if (postCondition.expression) {
                const expression = postCondition.expression;
                if (Array.isArray(expression)) {
                    expression.forEach(n => {
                        forEachChild(n, f);
                    });
                }
                else if (expression) {
                    forEachChild(expression, f);
                }
            }
            break;
        case SyntaxKind.IDENTIFIER:
        case SyntaxKind.TYPE_IDENTIFIER:
            goDeeper = f(node);
            if (!goDeeper)
                return;
            const identifier = node;
            if (identifier.args)
                identifier.args.forEach(arg => forEachChild(arg, f));
            break;
        case SyntaxKind.VARIABLE_DECLARATION:
            goDeeper = f(node);
            if (!goDeeper)
                return;
            const declaration = node;
            if (declaration.args)
                declaration.args.forEach(arg => forEachChild(arg, f));
            f(declaration.type);
        case SyntaxKind.NUMERIC_LITERAL:
        case SyntaxKind.STRING_LITERAL:
            f(node);
            break;
        default:
            break;
    }
}
exports.forEachChild = forEachChild;
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoic3RhdGVtZW50UGFyc2VyLmpzIiwic291cmNlUm9vdCI6IiIsInNvdXJjZXMiOlsiLi4vLi4vLi4vc3JjL3BhcnNlci9zdGF0ZW1lbnRQYXJzZXIudHMiXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6Ijs7O0FBQUEsMkNBQXFEO0FBRXJELElBQVksVUFtQlg7QUFuQkQsV0FBWSxVQUFVO0lBQ3JCLHVEQUFVLENBQUE7SUFDVixpRUFBZSxDQUFBO0lBQ2YsaUVBQWUsQ0FBQTtJQUNmLDJEQUFZLENBQUE7SUFDWiw2REFBYSxDQUFBO0lBQ2IsdURBQVUsQ0FBQTtJQUNWLDJEQUFZLENBQUE7SUFDWixpRUFBZSxDQUFBO0lBQ2YsK0RBQWMsQ0FBQTtJQUNkLCtEQUFjLENBQUE7SUFDZCxvRUFBZ0IsQ0FBQTtJQUNoQiw4REFBYSxDQUFBO0lBQ2IsOEVBQXFCLENBQUE7SUFDckIsZ0VBQWMsQ0FBQTtJQUNkLGtFQUFlLENBQUE7SUFDZixnRUFBYyxDQUFBO0lBQ2QsNEVBQW9CLENBQUE7SUFDcEIsa0VBQWUsQ0FBQTtBQUNoQixDQUFDLEVBbkJXLFVBQVUsR0FBVixrQkFBVSxLQUFWLGtCQUFVLFFBbUJyQjtBQUVELElBQUssY0F5Qko7QUF6QkQsV0FBSyxjQUFjO0lBQ2xCLHFDQUFtQixDQUFBO0lBQ25CLGtDQUFpQixDQUFBO0lBQ2pCLDBCQUFRLENBQUE7SUFDUixtQ0FBaUIsQ0FBQTtJQUNqQiw4QkFBWSxDQUFBO0lBQ1osNkJBQVcsQ0FBQTtJQUNYLDhCQUFZLENBQUE7SUFDWiwyQkFBUyxDQUFBO0lBQ1QsNkJBQVcsQ0FBQTtJQUNYLG1DQUFpQixDQUFBO0lBQ2pCLG9DQUFrQixDQUFBO0lBQ2xCLDRCQUFVLENBQUE7SUFDVixvQ0FBa0IsQ0FBQTtJQUNsQixpQ0FBZSxDQUFBO0lBQ2YsNkJBQVcsQ0FBQTtJQUNYLHFDQUFtQixDQUFBO0lBQ25CLG1DQUFpQixDQUFBO0lBQ2pCLDRCQUFVLENBQUE7SUFDVixxQ0FBbUIsQ0FBQTtJQUNuQixxQ0FBbUIsQ0FBQTtJ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