"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.parseFile = exports.parseText = exports.NON_TYPE_MODIFIERS = exports.MemberClass = void 0;
const fs = require("fs");
const statementParser_1 = require("./statementParser");
const tokenizer_1 = require("./tokenizer");
const utilities_1 = require("./utilities");
/**
 * Used for checking the type of Member at runtime
 */
var MemberClass;
(function (MemberClass) {
    MemberClass[MemberClass["method"] = 1] = "method";
    MemberClass[MemberClass["parameter"] = 2] = "parameter";
    MemberClass[MemberClass["property"] = 3] = "property";
    MemberClass[MemberClass["declaration"] = 4] = "declaration";
    MemberClass[MemberClass["column"] = 5] = "column";
    MemberClass[MemberClass["table"] = 6] = "table";
    MemberClass[MemberClass["proc"] = 7] = "proc";
})(MemberClass = exports.MemberClass || (exports.MemberClass = {}));
// tslint:disable-next-line:class-name
class _Method {
    constructor() {
        this.types = [];
        this.modifiers = [];
        this.parameters = [];
        this.line = -1;
        this.declarations = [];
        this.endLine = -1;
        this.memberClass = MemberClass.method;
        this.documentation = '';
        this.statements = [];
    }
}
// tslint:disable-next-line:class-name
class _Parameter {
    constructor() {
        this.modifiers = [];
        this.req = false;
        this.ret = false;
        this.literal = false;
        this.memberClass = MemberClass.parameter;
    }
}
const NON_METHOD_KEYWORDS = [
    'do', 'd', 'set', 's', 'if', 'i', 'for', 'f', 'while', 'w',
];
exports.NON_TYPE_MODIFIERS = [
    'public', 'static', 'private',
];
function parseText(sourceText) {
    const parser = new Parser();
    return parser.parseDocument(sourceText);
}
exports.parseText = parseText;
function parseFile(sourcePath) {
    return new Promise((resolve, reject) => {
        fs.readFile(sourcePath, (err, data) => {
            if (err) {
                reject(err);
            }
            else {
                const parser = new Parser();
                resolve(parser.parseDocument(data.toString()));
            }
        });
    });
}
exports.parseFile = parseFile;
class Parser {
    constructor(tokenizer) {
        this.methods = [];
        this.properties = [];
        this.declarations = [];
        this.tokens = [];
        this.comments = [];
        if (tokenizer)
            this.tokenizer = tokenizer;
    }
    parseDocument(documentText) {
        this.tokenizer = (0, tokenizer_1.getTokens)(documentText);
        while (this.next()) {
            if (this.activeToken.isAlphanumeric() || this.activeToken.isMinusSign()) {
                const method = this.parseMethod();
                if (!method)
                    continue;
                this.methods.push(method);
                this.activeMethod = method;
            }
            else if (this.activeToken.isTab() || this.activeToken.isSpace()) {
                const lineNumber = this.activeToken.position.line;
                const tokenBuffer = this.loadTokenBuffer();
                const propertyDef = this.lookForPropertyDef(tokenBuffer);
                if (propertyDef) {
                    if (propertyDef.id)
                        this.properties.push(propertyDef);
                    this.activeProperty = propertyDef;
                    continue;
                }
                const typeDec = this.lookForTypeDeclaration(tokenBuffer);
                if (typeDec.length > 0) {
                    const activeDeclarations = this.activeMethod ? this.activeMethod.declarations : this.declarations;
                    for (const dec of typeDec)
                        activeDeclarations.push(dec);
                    continue;
                }
                const extending = this.checkForExtends(tokenBuffer);
                if (extending)
                    this.extending = extending;
                const pslPackage = this.checkForPSLPackage(tokenBuffer);
                if (pslPackage)
                    this.pslPackage = pslPackage;
                if (this.activeMethod && this.activeMethod.batch && this.activeMethod.id.value === 'REVHIST') {
                    continue;
                }
                const statements = this.parseStatementsOnLine(tokenBuffer);
                if (statements && this.activeMethod)
                    this.activeMethod.statements = this.activeMethod.statements.concat(statements);
                if (this.activeProperty && this.activeProperty.id.position.line + 1 === lineNumber) {
                    const documentation = this.checkForDocumentation(tokenBuffer);
                    if (documentation)
                        this.activeProperty.documentation = documentation;
                }
                else if (this.activeMethod && (0, utilities_1.getLineAfter)(this.activeMethod) === lineNumber) {
                    const documentation = this.checkForDocumentation(tokenBuffer);
                    if (documentation)
                        this.activeMethod.documentation = documentation;
                }
            }
            else if (this.activeToken.isNewLine())
                continue;
            else
                this.throwAwayTokensTil(13 /* Type.NewLine */);
        }
        return {
            comments: this.comments,
            declarations: this.declarations,
            extending: this.extending,
            pslPackage: this.pslPackage,
            methods: this.methods,
            properties: this.properties,
            tokens: this.tokens,
        };
    }
    next() {
        this.activeToken = this.tokenizer.next().value;
        if (this.activeToken) {
            this.tokens.push(this.activeToken);
            if (this.activeToken.isLineComment() || this.activeToken.isBlockComment()) {
                this.comments.push(this.activeToken);
            }
        }
        return this.activeToken !== undefined;
    }
    checkForDocumentation(tokenBuffer) {
        let i = 0;
        while (i < tokenBuffer.length) {
            const token = tokenBuffer[i];
            if (token.isTab() || token.isSpace()) {
                i++;
                continue;
            }
            if (token.isBlockCommentInit() && tokenBuffer[i + 1] && tokenBuffer[i + 1].isBlockComment()) {
                return tokenBuffer[i + 1].value;
            }
            return '';
        }
    }
    lookForTypeDeclaration(tokenBuffer) {
        let i = 0;
        const tokens = [];
        while (i < tokenBuffer.length) {
            const token = tokenBuffer[i];
            if (token.isTab() || token.isSpace()) {
                i++;
                continue;
            }
            if (token.isAlphanumeric() && token.value === 'type') {
                for (let j = i + 1; j < tokenBuffer.length; j++) {
                    const loadToken = tokenBuffer[j];
                    if (loadToken.isSpace() || loadToken.isTab())
                        continue;
                    // if (loadToken.isEqualSign()) break;
                    tokens.push(loadToken);
                }
            }
            else if (token.isAlphanumeric() && token.value === 'catch') {
                for (let j = i + 1; j < tokenBuffer.length; j++) {
                    const loadToken = tokenBuffer[j];
                    if (loadToken.isSpace() || loadToken.isTab())
                        continue;
                    // if (loadToken.isEqualSign()) break;
                    tokens.push(new tokenizer_1.Token(1 /* Type.Alphanumeric */, 'Error', { character: 0, line: 0 }));
                    tokens.push(loadToken);
                    break;
                }
            }
            break;
        }
        const memberClass = MemberClass.declaration;
        const declarations = [];
        let type;
        let tokenIndex = 0;
        let id;
        let hasType;
        const modifiers = [];
        while (tokenIndex < tokens.length) {
            const token = tokens[tokenIndex];
            tokenIndex++;
            if (this.isDeclarationKeyword(token)) {
                modifiers.push(token);
                continue;
            }
            if (!hasType) {
                if (token.type !== 1 /* Type.Alphanumeric */)
                    break;
                if (token.value === 'static') {
                    modifiers.push(token);
                    hasType = true;
                }
                else {
                    type = token;
                    hasType = true;
                }
                continue;
            }
            else if (token.isAlphanumeric()) {
                id = token;
                if (hasType && !type)
                    type = token;
                // declarations.push({types: [type], identifier});
            }
            else if (token.isEqualSign()) {
                tokenIndex = this.skipToNextDeclaration(tokens, tokenIndex);
                if (id && type)
                    declarations.push({ types: [type], id, memberClass, modifiers });
                id = undefined;
            }
            else if (token.isOpenParen()) {
                const types = [];
                const myIdentifier = tokens[tokenIndex - 2];
                while (tokenIndex < tokens.length) {
                    const arrayTypeToken = tokens[tokenIndex];
                    tokenIndex++;
                    if (arrayTypeToken.isOpenParen())
                        continue;
                    else if (arrayTypeToken.isAlphanumeric()) {
                        types.push(arrayTypeToken);
                    }
                    else if (arrayTypeToken.isComma()) {
                        continue;
                    }
                    else if (arrayTypeToken.isCloseParen()) {
                        if (type)
                            declarations.push({ id: myIdentifier, types: [type].concat(types), memberClass, modifiers });
                        id = undefined;
                        break;
                    }
                }
            }
            // Cheating!!
            // else if (token.isPercentSign()) continue;
            else if (token.isComma()) {
                if (id && type)
                    declarations.push({ types: [type], id, memberClass, modifiers });
                id = undefined;
                continue;
            }
            else if (token.value === '\r')
                continue;
            else if (token.isBlockComment())
                continue;
            else if (token.isBlockCommentInit())
                continue;
            else if (token.isBlockCommentTerm())
                continue;
            else if (token.isNewLine()) {
                if (id && type)
                    declarations.push({ types: [type], id, memberClass, modifiers });
                id = undefined;
                break;
            }
            else
                break;
        }
        if (id && type)
            declarations.push({ types: [type], id, memberClass, modifiers });
        return declarations;
    }
    checkForExtends(tokenBuffer) {
        let i = 0;
        let classDef = false;
        let extending = false;
        let equals = false;
        while (i < tokenBuffer.length) {
            const token = tokenBuffer[i];
            if (token.isTab() || token.isSpace()) {
                i++;
                continue;
            }
            else if (token.isNumberSign() && !classDef) {
                const nextToken = tokenBuffer[i + 1];
                if (!nextToken)
                    return;
                if (nextToken.value === 'CLASSDEF') {
                    classDef = true;
                    i += 2;
                }
                else
                    break;
            }
            else if (token.value === 'extends' && !extending) {
                extending = true;
                i++;
            }
            else if (token.isEqualSign() && !equals) {
                equals = true;
                i++;
            }
            else if (token.isAlphanumeric() && classDef && extending && equals) {
                return token;
            }
            else {
                i++;
            }
        }
        return;
    }
    checkForPSLPackage(tokenBuffer) {
        let i = 0;
        let foundPackageToken = false;
        let fullPackage = '';
        while (i < tokenBuffer.length) {
            const token = tokenBuffer[i];
            if (token.isTab() || token.isSpace()) {
                i++;
                continue;
            }
            else if (token.isNumberSign() && !foundPackageToken) {
                const nextToken = tokenBuffer[i + 1];
                if (!nextToken)
                    return;
                if (nextToken.value === 'PACKAGE') {
                    foundPackageToken = true;
                    i += 2;
                }
                else
                    break;
            }
            else if (token.isAlphanumeric() && foundPackageToken) {
                // TODO: Maybe this should return an ordered list of tokens?
                if (fullPackage === '') {
                    fullPackage = token.value;
                }
                else {
                    fullPackage += ('.' + token.value);
                }
                i++;
            }
            else {
                i++;
            }
        }
        if (fullPackage !== '') {
            return fullPackage;
        }
        return;
    }
    skipToNextDeclaration(identifiers, tokenIndex) {
        let parenStack = 0;
        while (tokenIndex < identifiers.length) {
            const token = identifiers[tokenIndex];
            tokenIndex++;
            if (token.isOpenParen()) {
                parenStack++;
            }
            else if (token.isCloseParen()) {
                parenStack--;
            }
            else if (token.isComma() && parenStack === 0) {
                break;
            }
        }
        return tokenIndex;
    }
    isDeclarationKeyword(token) {
        if (token.type !== 1 /* Type.Alphanumeric */)
            return false;
        const keywords = ['public', 'private', 'new', 'literal'];
        return keywords.indexOf(token.value) !== -1;
    }
    throwAwayTokensTil(type) {
        while (this.next() && this.activeToken.type !== type)
            ;
    }
    loadTokenBuffer() {
        const tokenBuffer = [];
        while (this.next() && this.activeToken.type !== 13 /* Type.NewLine */) {
            tokenBuffer.push(this.activeToken);
        }
        return tokenBuffer;
    }
    lookForPropertyDef(tokenBuffer) {
        let i = 0;
        // TODO better loop
        while (i < tokenBuffer.length) {
            let token = tokenBuffer[i];
            if (token.isTab() || token.isSpace()) {
                i++;
                continue;
            }
            if (token.isNumberSign()) {
                token = tokenBuffer[i + 1];
                if (token && token.value === 'PROPERTYDEF') {
                    const tokens = tokenBuffer.filter(t => {
                        if (t.isNumberSign())
                            return false;
                        if (t.value === 'PROPERTYDEF')
                            return false;
                        return t.type !== 32 /* Type.Space */ && t.type !== 11 /* Type.Tab */;
                    });
                    const classTypes = [];
                    const classIndex = tokens.findIndex(t => t.value === 'class');
                    if (tokens[classIndex + 1]
                        && tokens[classIndex + 1].value === '='
                        && tokens[classIndex + 2]
                        && tokens[classIndex + 2].isAlphanumeric()) {
                        classTypes.push(tokens[classIndex + 2]);
                    }
                    return {
                        id: tokens[0],
                        memberClass: MemberClass.property,
                        modifiers: this.findPropertyModifiers(tokens.slice(1)),
                        types: classTypes,
                    };
                }
                else {
                    break;
                }
            }
            else {
                break;
            }
        }
        return;
    }
    findPropertyModifiers(tokens) {
        return tokens.filter(t => {
            return t.value === 'private' || t.value === 'literal' || t.value === 'public';
        });
    }
    parseMethod() {
        let batchLabel = false;
        const method = new _Method();
        do {
            if (!this.activeToken)
                continue;
            if (this.activeToken.isTab() || this.activeToken.isSpace())
                continue;
            else if (this.activeToken.isNewLine())
                break;
            else if (this.activeToken.isOpenParen()) {
                const processed = this.processParameters(method);
                if (!processed)
                    return undefined;
                method.parameters = processed;
                break;
            }
            else if (this.activeToken.isAlphanumeric() || this.activeToken.isNumeric()) {
                if (batchLabel) {
                    method.modifiers.push(this.activeToken);
                    method.batch = true;
                    break;
                }
                if (method.line === -1) {
                    method.line = this.activeToken.position.line;
                }
                method.modifiers.push(this.activeToken);
            }
            else if (this.activeToken.isMinusSign()) {
                batchLabel = true;
                continue;
            }
            else if (this.activeToken.isLineCommentInit()
                || this.activeToken.isLineComment()
                || this.activeToken.isBlockCommentInit()
                || this.activeToken.isBlockComment()
                || this.activeToken.isBlockCommentTerm()) {
                continue;
            }
            else if (this.activeToken.value === '\r')
                continue;
            else if (this.activeToken.isCloseParen()) {
                if (!method.closeParen) {
                    method.closeParen = this.activeToken;
                }
            }
            else {
                this.throwAwayTokensTil(13 /* Type.NewLine */);
                if (method.modifiers.length > 1) {
                    break;
                }
                return undefined;
            }
        } while (this.next());
        return this.finalizeMethod(method);
    }
    finalizeMethod(method) {
        for (const keyword of NON_METHOD_KEYWORDS) {
            const index = method.modifiers.map(i => i.value.toLowerCase()).indexOf(keyword.toLowerCase());
            if (index > -1 && index <= method.modifiers.length - 1) {
                method.modifiers = [method.modifiers[0]];
                method.parameters = [];
                break;
            }
        }
        // better way...
        method.id = method.modifiers.pop();
        if (this.activeMethod) {
            this.activeMethod.endLine = method.id.position.line - 1;
        }
        const lastModifier = method.modifiers[method.modifiers.length - 1];
        if (lastModifier && exports.NON_TYPE_MODIFIERS.indexOf(lastModifier.value) < 0) {
            method.types = [method.modifiers.pop()];
        }
        this.activeMethod = method;
        return method;
    }
    processParameters(method) {
        const args = [];
        let param;
        let open = false;
        while (this.next()) {
            if (this.activeToken.isTab() || this.activeToken.isSpace() || this.activeToken.isNewLine())
                continue;
            else if (this.activeToken.isOpenParen()) {
                open = true;
                if (!param)
                    return undefined;
                if (param.types.length === 1 && !param.id) {
                    param.id = param.types[0];
                    param.types[0] = this.getDummy();
                }
                const objectArgs = this.processObjectArgs();
                if (!objectArgs)
                    return undefined;
                param.types = param.types.concat(objectArgs);
                continue;
            }
            else if (this.activeToken.isCloseParen()) {
                open = false;
                method.closeParen = this.activeToken;
                if (!param)
                    break;
                if (param.types.length === 1 && !param.id) {
                    param.id = param.types[0];
                    param.types[0] = this.getDummy();
                }
                args.push(param);
                break;
            }
            else if (this.activeToken.isAlphanumeric()) {
                if (!param)
                    param = new _Parameter();
                // let value = this.activeToken.value;
                if (this.activeToken.value === 'req') {
                    param.modifiers.push(this.activeToken);
                    param.req = true;
                }
                else if (this.activeToken.value === 'ret') {
                    param.modifiers.push(this.activeToken);
                    param.ret = true;
                }
                else if (this.activeToken.value === 'literal') {
                    param.modifiers.push(this.activeToken);
                    param.literal = true;
                }
                else if (!param.types)
                    param.types = [this.activeToken];
                else {
                    param.id = this.activeToken;
                }
            }
            else if (this.activeToken.isLineComment()) {
                if (param) {
                    param.comment = this.activeToken;
                }
                else if (args.length >= 1) {
                    args[args.length - 1].comment = this.activeToken;
                }
            }
            else if (this.activeToken.isComma()) {
                if (!param)
                    return undefined;
                if (param.types.length === 1 && !param.id) {
                    param.id = param.types[0];
                    param.types[0] = this.getDummy();
                }
                args.push(param);
                param = undefined;
            }
        }
        if (open)
            return undefined;
        return args;
    }
    processObjectArgs() {
        const types = [];
        let found = false;
        while (this.next()) {
            const dummy = this.getDummy();
            if (this.activeToken.isTab() || this.activeToken.isSpace())
                continue;
            else if (this.activeToken.isCloseParen()) {
                if (types.length === 0)
                    types.push(dummy);
                return types;
            }
            else if (this.activeToken.isAlphanumeric()) {
                if (!found) {
                    types.push(this.activeToken);
                    found = true;
                }
                else
                    return undefined;
            }
            else if (this.activeToken.isComma()) {
                if (!found) {
                    if (types.length === 0) {
                        types.push(dummy);
                    }
                    types.push(dummy);
                }
                found = false;
                continue;
            }
        }
        return undefined;
    }
    parseStatementsOnLine(tokenBuffer) {
        const statementParser = new statementParser_1.StatementParser(tokenBuffer);
        try {
            return statementParser.parseLine();
        }
        catch (_a) {
            return [];
        }
    }
    getDummy() {
        return new tokenizer_1.Token(-1 /* Type.Undefined */, '', this.activeToken.position);
    }
}
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoicGFyc2VyLmpzIiwic291cmNlUm9vdCI6IiIsInNvdXJjZXMiOlsiLi4vLi4vLi4vc3JjL3BhcnNlci9wYXJzZXIudHMiXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6Ijs7O0FBQUEseUJBQXlCO0FBQ3pCLHVEQUErRDtBQUMvRCwyQ0FBcUQ7QUFDckQsMkNBQTJDO0FBRTNDOztHQUVHO0FBQ0gsSUFBWSxXQVFYO0FBUkQsV0FBWSxXQUFXO0lBQ3RCLGlEQUFVLENBQUE7SUFDVix1REFBYSxDQUFBO0lBQ2IscURBQVksQ0FBQTtJQUNaLDJEQUFlLENBQUE7SUFDZixpREFBVSxDQUFBO0lBQ1YsK0NBQVMsQ0FBQTtJQUNULDZDQUFRLENBQUE7QUFDVCxDQUFDLEVBUlcsV0FBVyxHQUFYLG1CQUFXLEtBQVgsbUJBQVcsUUFRdEI7QUE2S0Qsc0NBQXNDO0FBQ3RDLE1BQU0sT0FBTztJQWVaO1FBQ0MsSUFBSSxDQUFDLEtBQUssR0FBRyxFQUFFLENBQUM7UUFDaEIsSUFBSSxDQUFDLFNBQVMsR0FBRyxFQUFFLENBQUM7UUFDcEIsSUFBSSxDQUFDLFVBQVUsR0FBRyxFQUFFLENBQUM7UUFDckIsSUFBSSxDQUFDLElBQUksR0FBRyxDQUFDLENBQUMsQ0FBQztRQUNmLElBQUksQ0FBQyxZQUFZLEdBQUcsRUFBRSxDQUFDO1FBQ3ZCLElBQUksQ0FBQyxPQUFPLEdBQUcsQ0FBQyxDQUFDLENBQUM7UUFDbEIsSUFBSSxDQUFDLFdBQVcsR0FBRyxXQUFXLENBQUMsTUFBTSxDQUFDO1FBQ3RDLElBQUksQ0FBQyxhQUFhLEdBQUcsRUFBRSxDQUFDO1FBQ3hCLElBQUksQ0FBQyxVQUFVLEdBQUcsRUFBRSxDQUFDO0lBQ3RCLENBQUM7Q0FDRDtBQUVELHNDQUFzQztBQUN0QyxNQUFNLFVBQVU7SUFVZjtRQUNDLElBQUksQ0FBQyxTQUFTLEdBQUcsRUFBRSxDQUFDO1FBQ3BCLElBQUksQ0FBQyxHQUFHLEdBQUcsS0FBSyxDQUFDO1FBQ2pCLElBQUksQ0FBQyxHQUFHLEdBQUcsS0FBSyxDQUFDO1FBQ2pCLElBQUksQ0FBQyxPQUFPLEdBQUcsS0FBSyxDQUFDO1FBQ3JCLElBQUksQ0FBQyxXQUFXLEdBQUcsV0FBVyxDQUFDLFNBQVMsQ0FBQztJQUMxQyxDQUFDO0NBQ0Q7QUFFRCxNQUFNLG1CQUFtQixHQUFHO0lBQzNCLElBQUksRUFBRSxHQUFHLEVBQUUsS0FBSyxFQUFFLEdBQUcsRUFBRSxJQUFJLEVBQUUsR0FBRyxFQUFFLEtBQUssRUFBRSxHQUFHLEVBQUUsT0FBTyxFQUFFLEdBQUc7Q0FDMUQsQ0FBQztBQUVXLFFBQUEsa0JBQWtCLEdBQUc7SUFDakMsUUFBUSxFQUFFLFFBQVEsRUFBRSxTQUFTO0NBQzdCLENBQUM7QUFFRixTQUFnQixTQUFTLENBQUMsVUFBa0I7SUFDM0MsTUFBTSxNQUFNLEdBQUcsSUFBSSxNQUFNLEVBQUUsQ0FBQztJQUM1QixPQUFPLE1BQU0sQ0FBQyxhQUFhLENBQUMsVUFBVSxDQUFDLENBQUM7QUFDekMsQ0FBQztBQUhELDhCQUdDO0FBRUQsU0FBZ0IsU0FBUyxDQUFDLFVBQWtCO0lBQzNDLE9BQU8sSUFBSSxPQUFPLENBQUMsQ0FBQyxPQUFPLEVBQUUsTUFBTSxFQUFFLEVBQUU7UUFDdEMsRUFBRSxDQUFDLFFBQVEsQ0FBQyxVQUFVLEVBQUUsQ0FBQyxHQUFHLEVBQUUsSUFBSSxFQUFFLEVBQUU7WUFDckMsSUFBSSxHQUFHLEVBQUU7Z0JBQ1IsTUFBTSxDQUFDLEdBQUcsQ0FBQyxDQUFDO2FBQ1o7aUJBQ0k7Z0JBQ0osTUFBTSxNQUFNLEdBQUcsSUFBSSxNQUFNLEVBQUUsQ0FBQztnQkFDNUIsT0FBTyxDQUFDLE1BQU0sQ0FBQyxhQUFhLENBQUMsSUFBSSxDQUFDLFFBQVEsRUFBRSxDQUFDLENBQUMsQ0FBQzthQUMvQztRQUNGLENBQUMsQ0FBQyxDQUFDO0lBQ0osQ0FBQyxDQUFDLENBQUM7QUFDSixDQUFDO0FBWkQsOEJBWUM7QUFFRCxNQUFNLE1BQU07SUFjWCxZQUFZLFNBQW1DO1FBQzlDLElBQUksQ0FBQyxPQUFPLEdBQUcsRUFBRSxDQUFDO1FBQ2xCLElBQUksQ0FBQyxVQUFVLEdBQUcsRUFBRSxDQUFDO1FBQ3JCLElBQUksQ0FBQyxZQUFZLEdBQUcsRUFBRSxDQUFDO1FBQ3ZCLElBQUksQ0FBQyxNQUFNLEdBQUcsRUFBRSxDQUFDO1FBQ2pCLElBQUksQ0FBQyxRQUFRLEdBQUcsRUFBRSxDQUFDO1FBQ25CLElBQUksU0FBUztZQUFFLElBQUksQ0FBQyxTQUFTLEdBQUcsU0FBUyxDQUFDO0lBQzNDLENBQUM7SUFFRCxhQUFhLENBQUMsWUFBb0I7UUFDakMsSUFBSSxDQUFDLFNBQVMsR0FBRyxJQUFBLHFCQUFTLEVBQUMsWUFBWSxDQUFDLENBQUM7UUFDekMsT0FBTyxJQUFJLENBQUMsSUFBSSxFQUFFLEVBQUU7WUFDbkIsSUFBSSxJQUFJLENBQUMsV0FBVyxDQUFDLGNBQWMsRUFBRSxJQUFJLElBQUksQ0FBQyxXQUFXLENBQUMsV0FBVyxFQUFFLEVBQUU7Z0JBQ3hFLE1BQU0sTUFBTSxHQUFHLElBQUksQ0FBQyxXQUFXLEVBQUUsQ0FBQztnQkFDbEMsSUFBSSxDQUFDLE1BQU07b0JBQUUsU0FBUztnQkFDdEIsSUFBSSxDQUFDLE9BQU8sQ0FBQyxJQUFJLENBQUMsTUFBTSxDQUFDLENBQUM7Z0JBQzFCLElBQUksQ0FBQyxZQUFZLEdBQUcsTUFBTSxDQUFDO2FBQzNCO2lCQUNJLElBQUksSUFBSSxDQUFDLFdBQVcsQ0FBQyxLQUFLLEVBQUUsSUFBSSxJQUFJLENBQUMsV0FBVyxDQUFDLE9BQU8sRUFBRSxFQUFFO2dCQUNoRSxNQUFNLFVBQVUsR0FBRyxJQUFJLENBQUMsV0FBVyxDQUFDLFFBQVEsQ0FBQyxJQUFJLENBQUM7Z0JBQ2xELE1BQU0sV0FBVyxHQUFHLElBQUksQ0FBQyxlQUFlLEVBQUUsQ0FBQztnQkFDM0MsTUFBTSxXQUFXLEdBQUcsSUFBSSxDQUFDLGtCQUFrQixDQUFDLFdBQVcsQ0FBQyxDQUFDO2dCQUN6RCxJQUFJLFdBQVcsRUFBRTtvQkFDaEIsSUFBSSxXQUFXLENBQUMsRUFBRTt3QkFBRSxJQUFJLENBQUMsVUFBVSxDQUFDLElBQUksQ0FBQyxXQUFXLENBQUMsQ0FBQztvQkFDdEQsSUFBSSxDQUFDLGNBQWMsR0FBRyxXQUFXLENBQUM7b0JBQ2xDLFNBQVM7aUJBQ1Q7Z0JBQ0QsTUFBTSxPQUFPLEdBQUcsSUFBSSxDQUFDLHNCQUFzQixDQUFDLFdBQVcsQ0FBQyxDQUFDO2dCQUN6RCxJQUFJLE9BQU8sQ0FBQyxNQUFNLEdBQUcsQ0FBQyxFQUFFO29CQUN2QixNQUFNLGtCQUFrQixHQUFHLElBQUksQ0FBQyxZQUFZLENBQUMsQ0FBQyxDQUFDLElBQUksQ0FBQyxZQUFZLENBQUMsWUFBWSxDQUFDLENBQUMsQ0FBQyxJQUFJLENBQUMsWUFBWSxDQUFDO29CQUNsRyxLQUFLLE1BQU0sR0FBRyxJQUFJLE9BQU87d0JBQUUsa0JBQWtCLENBQUMsSUFBSSxDQUFDLEdBQUcsQ0FBQyxDQUFDO29CQUN4RCxTQUFTO2lCQUNUO2dCQUNELE1BQU0sU0FBUyxHQUFHLElBQUksQ0FBQyxlQUFlLENBQUMsV0FBVyxDQUFDLENBQUM7Z0JBQ3BELElBQUksU0FBUztvQkFBRSxJQUFJLENBQUMsU0FBUyxHQUFHLFNBQVMsQ0FBQztnQkFDMUMsTUFBTSxVQUFVLEdBQUcsSUFBSSxDQUFDLGtCQUFrQixDQUFDLFdBQVcsQ0FBQyxDQUFDO2dCQUN4RCxJQUFJLFVBQVU7b0JBQUUsSUFBSSxDQUFDLFVBQVUsR0FBRyxVQUFVLENBQUM7Z0JBQzdDLElBQUksSUFBSSxDQUFDLFlBQVksSUFBSSxJQUFJLENBQUMsWUFBWSxDQUFDLEtBQUssSUFBSSxJQUFJLENBQUMsWUFBWSxDQUFDLEVBQUUsQ0FBQyxLQUFLLEtBQUssU0FBUyxFQUFFO29CQUM3RixTQUFTO2lCQUNUO2dCQUNELE1BQU0sVUFBVSxHQUFHLElBQUksQ0FBQyxxQkFBcUIsQ0FBQyxXQUFXLENBQUMsQ0FBQztnQkFDM0QsSUFBSSxVQUFVLElBQUksSUFBSSxDQUFDLFlBQVk7b0JBQUUsSUFBSSxDQUFDLFlBQVksQ0FBQyxVQUFVLEdBQUcsSUFBSSxDQUFDLFlBQVksQ0FBQyxVQUFVLENBQUMsTUFBTSxDQUFDLFVBQVUsQ0FBQyxDQUFDO2dCQUNwSCxJQUFJLElBQUksQ0FBQyxjQUFjLElBQUksSUFBSSxDQUFDLGNBQWMsQ0FBQyxFQUFFLENBQUMsUUFBUSxDQUFDLElBQUksR0FBRyxDQUFDLEtBQUssVUFBVSxFQUFFO29CQUNuRixNQUFNLGFBQWEsR0FBRyxJQUFJLENBQUMscUJBQXFCLENBQUMsV0FBVyxDQUFDLENBQUM7b0JBQzlELElBQUksYUFBYTt3QkFBRSxJQUFJLENBQUMsY0FBYyxDQUFDLGFBQWEsR0FBRyxhQUFhLENBQUM7aUJBQ3JFO3FCQUNJLElBQUksSUFBSSxDQUFDLFlBQVksSUFBSSxJQUFBLHdCQUFZLEVBQUMsSUFBSSxDQUFDLFlBQVksQ0FBQyxLQUFLLFVBQVUsRUFBRTtvQkFDN0UsTUFBTSxhQUFhLEdBQUcsSUFBSSxDQUFDLHFCQUFxQixDQUFDLFdBQVcsQ0FBQyxDQUFDO29CQUM5RCxJQUFJLGFBQWE7d0JBQUUsSUFBSSxDQUFDLFlBQVksQ0FBQyxhQUFhLEdBQUcsYUFBYSxDQUFDO2lCQUNuRTthQUNEO2lCQUNJLElBQUksSUFBSSxDQUFDLFdBQVcsQ0FBQyxTQUFTLEVBQUU7Z0JBQUUsU0FBUzs7Z0JBQzNDLElBQUksQ0FBQyxrQkFBa0IsdUJBQWMsQ0FBQztTQUMzQztRQUNELE9BQU87WUFDTixRQUFRLEVBQUUsSUFBSSxDQUFDLFFBQVE7WUFDdkIsWUFBWSxFQUFFLElBQUksQ0FBQyxZQUFZO1lBQy9CLFNBQVMsRUFBRSxJQUFJLENBQUMsU0FBUztZQUN6QixVQUFVLEVBQUUsSUFBSSxDQUFDLFVBQVU7WUFDM0IsT0FBTyxFQUFFLElBQUksQ0FBQyxPQUFPO1lBQ3JCLFVBQVUsRUFBRSxJQUFJLENBQUMsVUFBVTtZQUMzQixNQUFNLEVBQUUsSUFBSSxDQUFDLE1BQU07U0FDbkIsQ0FBQztJQUNILENBQUM7SUFFTyxJQUFJO1FBQ1gsSUFBSSxDQUFDLFdBQVcsR0FBRyxJQUFJLENBQUMsU0FBUyxDQUFDLElBQUksRUFBRSxDQUFDLEtBQUssQ0FBQztRQUMvQyxJQUFJLElBQUksQ0FBQyxXQUFXLEVBQUU7WUFDckIsSUFBSSxDQUFDLE1BQU0sQ0FBQyxJQUFJLENBQUMsSUFBSSxDQUFDLFdBQVcsQ0FBQyxDQUFDO1lBQ25DLElBQUksSUFBSSxDQUFDLFdBQVcsQ0FBQyxhQUFhLEVBQUUsSUFBSSxJQUFJLENBQUMsV0FBVyxDQUFDLGNBQWMsRUFBRSxFQUFFO2dCQUMxRSxJQUFJLENBQUMsUUFBUSxDQUFDLElBQUksQ0FBQyxJQUFJLENBQUMsV0FBVyxDQUFDLENBQUM7YUFDckM7U0FDRDtRQUNELE9BQU8sSUFBSSxDQUFDLFdBQVcsS0FBSyxTQUFTLENBQUM7SUFDdkMsQ0FBQztJQUVPLHFCQUFxQixDQUFDLFdBQW9CO1FBQ2pELElBQUksQ0FBQyxHQUFHLENBQUMsQ0FBQztRQUNWLE9BQU8sQ0FBQyxHQUFHLFdBQVcsQ0FBQyxNQUFNLEVBQUU7WUFDOUIsTUFBTSxLQUFLLEdBQUcsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDO1lBQzdCLElBQUksS0FBSyxDQUFDLEtBQUssRUFBRSxJQUFJLEtBQUssQ0FBQyxPQUFPLEVBQUUsRUFBRTtnQkFDckMsQ0FBQyxFQUFFLENBQUM7Z0JBQ0osU0FBUzthQUNUO1lBQ0QsSUFBSSxLQUFLLENBQUMsa0JBQWtCLEVBQUUsSUFBSSxXQUFXLENBQUMsQ0FBQyxHQUFHLENBQUMsQ0FBQyxJQUFJLFdBQVcsQ0FBQyxDQUFDLEdBQUcsQ0FBQyxDQUFDLENBQUMsY0FBYyxFQUFFLEVBQUU7Z0JBQzVGLE9BQU8sV0FBVyxDQUFDLENBQUMsR0FBRyxDQUFDLENBQUMsQ0FBQyxLQUFLLENBQUM7YUFDaEM7WUFDRCxPQUFPLEVBQUUsQ0FBQztTQUNWO0lBQ0YsQ0FBQztJQUVPLHNCQUFzQixDQUFDLFdBQW9CO1FBQ2xELElBQUksQ0FBQyxHQUFHLENBQUMsQ0FBQztRQUNWLE1BQU0sTUFBTSxHQUFZLEVBQUUsQ0FBQztRQUMzQixPQUFPLENBQUMsR0FBRyxXQUFXLENBQUMsTUFBTSxFQUFFO1lBQzlCLE1BQU0sS0FBSyxHQUFHLFdBQVcsQ0FBQyxDQUFDLENBQUMsQ0FBQztZQUM3QixJQUFJLEtBQUssQ0FBQyxLQUFLLEVBQUUsSUFBSSxLQUFLLENBQUMsT0FBTyxFQUFFLEVBQUU7Z0JBQ3JDLENBQUMsRUFBRSxDQUFDO2dCQUNKLFNBQVM7YUFDVDtZQUNELElBQUksS0FBSyxDQUFDLGNBQWMsRUFBRSxJQUFJLEtBQUssQ0FBQyxLQUFLLEtBQUssTUFBTSxFQUFFO2dCQUNyRCxLQUFLLElBQUksQ0FBQyxHQUFHLENBQUMsR0FBRyxDQUFDLEVBQUUsQ0FBQyxHQUFHLFdBQVcsQ0FBQyxNQUFNLEVBQUUsQ0FBQyxFQUFFLEVBQUU7b0JBQ2hELE1BQU0sU0FBUyxHQUFHLFdBQVcsQ0FBQyxDQUFDLENBQUMsQ0FBQztvQkFDakMsSUFBSSxTQUFTLENBQUMsT0FBTyxFQUFFLElBQUksU0FBUyxDQUFDLEtBQUssRUFBRTt3QkFBRSxTQUFTO29CQUN2RCxzQ0FBc0M7b0JBQ3RDLE1BQU0sQ0FBQyxJQUFJLENBQUMsU0FBUyxDQUFDLENBQUM7aUJBQ3ZCO2FBQ0Q7aUJBQ0ksSUFBSSxLQUFLLENBQUMsY0FBYyxFQUFFLElBQUksS0FBSyxDQUFDLEtBQUssS0FBSyxPQUFPLEVBQUU7Z0JBQzNELEtBQUssSUFBSSxDQUFDLEdBQUcsQ0FBQyxHQUFHLENBQUMsRUFBRSxDQUFDLEdBQUcsV0FBVyxDQUFDLE1BQU0sRUFBRSxDQUFDLEVBQUUsRUFBRTtvQkFDaEQsTUFBTSxTQUFTLEdBQUcsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDO29CQUNqQyxJQUFJLFNBQVMsQ0FBQyxPQUFPLEVBQUUsSUFBSSxTQUFTLENBQUMsS0FBSyxFQUFFO3dCQUFFLFNBQVM7b0JBQ3ZELHNDQUFzQztvQkFDdEMsTUFBTSxDQUFDLElBQUksQ0FBQyxJQUFJLGlCQUFLLDRCQUFvQixPQUFPLEVBQUUsRUFBRSxTQUFTLEVBQUUsQ0FBQyxFQUFFLElBQUksRUFBRSxDQUFDLEVBQUUsQ0FBQyxDQUFDLENBQUM7b0JBQzlFLE1BQU0sQ0FBQyxJQUFJLENBQUMsU0FBUyxDQUFDLENBQUM7b0JBQ3ZCLE1BQU07aUJBQ047YUFDRDtZQUNELE1BQU07U0FDTjtRQUNELE1BQU0sV0FBVyxHQUFHLFdBQVcsQ0FBQyxXQUFXLENBQUM7UUFDNUMsTUFBTSxZQUFZLEdBQWtCLEVBQUUsQ0FBQztRQUN2QyxJQUFJLElBQUksQ0FBQztRQUNULElBQUksVUFBVSxHQUFHLENBQUMsQ0FBQztRQUNuQixJQUFJLEVBQUUsQ0FBQztRQUNQLElBQUksT0FBTyxDQUFDO1FBQ1osTUFBTSxTQUFTLEdBQVksRUFBRSxDQUFDO1FBQzlCLE9BQU8sVUFBVSxHQUFHLE1BQU0sQ0FBQyxNQUFNLEVBQUU7WUFDbEMsTUFBTSxLQUFLLEdBQUcsTUFBTSxDQUFDLFVBQVUsQ0FBQyxDQUFDO1lBQ2pDLFVBQVUsRUFBRSxDQUFDO1lBQ2IsSUFBSSxJQUFJLENBQUMsb0JBQW9CLENBQUMsS0FBSyxDQUFDLEVBQUU7Z0JBQ3JDLFNBQVMsQ0FBQyxJQUFJLENBQUMsS0FBSyxDQUFDLENBQUM7Z0JBQ3RCLFNBQVM7YUFDVDtZQUNELElBQUksQ0FBQyxPQUFPLEVBQUU7Z0JBQ2IsSUFBSSxLQUFLLENBQUMsSUFBSSw4QkFBc0I7b0JBQUUsTUFBTTtnQkFDNUMsSUFBSSxLQUFLLENBQUMsS0FBSyxLQUFLLFFBQVEsRUFBRTtvQkFDN0IsU0FBUyxDQUFDLElBQUksQ0FBQyxLQUFLLENBQUMsQ0FBQztvQkFDdEIsT0FBTyxHQUFHLElBQUksQ0FBQztpQkFDZjtxQkFDSTtvQkFDSixJQUFJLEdBQUcsS0FBSyxDQUFDO29CQUNiLE9BQU8sR0FBRyxJQUFJLENBQUM7aUJBQ2Y7Z0JBQ0QsU0FBUzthQUNUO2lCQUNJLElBQUksS0FBSyxDQUFDLGNBQWMsRUFBRSxFQUFFO2dCQUNoQyxFQUFFLEdBQUcsS0FBSyxDQUFDO2dCQUNYLElBQUksT0FBTyxJQUFJLENBQUMsSUFBSTtvQkFBRSxJQUFJLEdBQUcsS0FBSyxDQUFDO2dCQUNuQyxrREFBa0Q7YUFDbEQ7aUJBQ0ksSUFBSSxLQUFLLENBQUMsV0FBVyxFQUFFLEVBQUU7Z0JBQzdCLFVBQVUsR0FBRyxJQUFJLENBQUMscUJBQXFCLENBQUMsTUFBTSxFQUFFLFVBQVUsQ0FBQyxDQUFDO2dCQUM1RCxJQUFJLEVBQUUsSUFBSSxJQUFJO29CQUFFLFlBQVksQ0FBQyxJQUFJLENBQUMsRUFBRSxLQUFLLEVBQUUsQ0FBQyxJQUFJLENBQUMsRUFBRSxFQUFFLEVBQUUsV0FBVyxFQUFFLFNBQVMsRUFBRSxDQUFDLENBQUM7Z0JBQ2pGLEVBQUUsR0FBRyxTQUFTLENBQUM7YUFDZjtpQkFDSSxJQUFJLEtBQUssQ0FBQyxXQUFXLEVBQUUsRUFBRTtnQkFDN0IsTUFBTSxLQUFLLEdBQUcsRUFBRSxDQUFDO2dCQUNqQixNQUFNLFlBQVksR0FBRyxNQUFNLENBQUMsVUFBVSxHQUFHLENBQUMsQ0FBQyxDQUFDO2dCQUM1QyxPQUFPLFVBQVUsR0FBRyxNQUFNLENBQUMsTUFBTSxFQUFFO29CQUNsQyxNQUFNLGNBQWMsR0FBRyxNQUFNLENBQUMsVUFBVSxDQUFDLENBQUM7b0JBQzFDLFVBQVUsRUFBRSxDQUFDO29CQUNiLElBQUksY0FBYyxDQUFDLFdBQVcsRUFBRTt3QkFBRSxTQUFTO3lCQUN0QyxJQUFJLGNBQWMsQ0FBQyxjQUFjLEVBQUUsRUFBRTt3QkFDekMsS0FBSyxDQUFDLElBQUksQ0FBQyxjQUFjLENBQUMsQ0FBQztxQkFDM0I7eUJBQ0ksSUFBSSxjQUFjLENBQUMsT0FBTyxFQUFFLEVBQUU7d0JBQ2xDLFNBQVM7cUJBQ1Q7eUJBQ0ksSUFBSSxjQUFjLENBQUMsWUFBWSxFQUFFLEVBQUU7d0JBQ3ZDLElBQUksSUFBSTs0QkFBRSxZQUFZLENBQUMsSUFBSSxDQUFDLEVBQUUsRUFBRSxFQUFFLFlBQVksRUFBRSxLQUFLLEVBQUUsQ0FBQyxJQUFJLENBQUMsQ0FBQyxNQUFNLENBQUMsS0FBSyxDQUFDLEVBQUUsV0FBVyxFQUFFLFNBQVMsRUFBRSxDQUFDLENBQUM7d0JBQ3ZHLEVBQUUsR0FBRyxTQUFTLENBQUM7d0JBQ2YsTUFBTTtxQkFDTjtpQkFDRDthQUNEO1lBQ0QsYUFBYTtZQUNiLDRDQUE0QztpQkFDdkMsSUFBSSxLQUFLLENBQUMsT0FBTyxFQUFFLEVBQUU7Z0JBQ3pCLElBQUksRUFBRSxJQUFJLElBQUk7b0JBQUUsWUFBWSxDQUFDLElBQUksQ0FBQyxFQUFFLEtBQUssRUFBRSxDQUFDLElBQUksQ0FBQyxFQUFFLEVBQUUsRUFBRSxXQUFXLEVBQUUsU0FBUyxFQUFFLENBQUMsQ0FBQztnQkFDakYsRUFBRSxHQUFHLFNBQVMsQ0FBQztnQkFDZixTQUFTO2FBQ1Q7aUJBQ0ksSUFBSSxLQUFLLENBQUMsS0FBSyxLQUFLLElBQUk7Z0JBQUUsU0FBUztpQkFDbkMsSUFBSSxLQUFLLENBQUMsY0FBYyxFQUFFO2dCQUFFLFNBQVM7aUJBQ3JDLElBQUksS0FBSyxDQUFDLGtCQUFrQixFQUFFO2dCQUFFLFNBQVM7aUJBQ3pDLElBQUksS0FBSyxDQUFDLGtCQUFrQixFQUFFO2dCQUFFLFNBQVM7aUJBQ3pDLElBQUksS0FBSyxDQUFDLFNBQVMsRUFBRSxFQUFFO2dCQUMzQixJQUFJLEVBQUUsSUFBSSxJQUFJO29CQUFFLFlBQVksQ0FBQyxJQUFJLENBQUMsRUFBRSxLQUFLLEVBQUUsQ0FBQyxJQUFJLENBQUMsRUFBRSxFQUFFLEVBQUUsV0FBVyxFQUFFLFNBQVMsRUFBRSxDQUFDLENBQUM7Z0JBQ2pGLEVBQUUsR0FBRyxTQUFTLENBQUM7Z0JBQ2YsTUFBTTthQUNOOztnQkFDSSxNQUFNO1NBQ1g7UUFDRCxJQUFJLEVBQUUsSUFBSSxJQUFJO1lBQUUsWUFBWSxDQUFDLElBQUksQ0FBQyxFQUFFLEtBQUssRUFBRSxDQUFDLElBQUksQ0FBQyxFQUFFLEVBQUUsRUFBRSxXQUFXLEVBQUUsU0FBUyxFQUFFLENBQUMsQ0FBQztRQUNqRixPQUFPLFlBQVksQ0FBQztJQUNyQixDQUFDO0lBRU8sZUFBZSxDQUFDLFdBQW9CO1FBQzNDLElBQUksQ0FBQyxHQUFHLENBQUMsQ0FBQztRQUNWLElBQUksUUFBUSxHQUFHLEtBQUssQ0FBQztRQUNyQixJQUFJLFNBQVMsR0FBRyxLQUFLLENBQUM7UUFDdEIsSUFBSSxNQUFNLEdBQUcsS0FBSyxDQUFDO1FBQ25CLE9BQU8sQ0FBQyxHQUFHLFdBQVcsQ0FBQyxNQUFNLEVBQUU7WUFDOUIsTUFBTSxLQUFLLEdBQUcsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDO1lBQzdCLElBQUksS0FBSyxDQUFDLEtBQUssRUFBRSxJQUFJLEtBQUssQ0FBQyxPQUFPLEVBQUUsRUFBRTtnQkFDckMsQ0FBQyxFQUFFLENBQUM7Z0JBQ0osU0FBUzthQUNUO2lCQUNJLElBQUksS0FBSyxDQUFDLFlBQVksRUFBRSxJQUFJLENBQUMsUUFBUSxFQUFFO2dCQUMzQyxNQUFNLFNBQVMsR0FBRyxXQUFXLENBQUMsQ0FBQyxHQUFHLENBQUMsQ0FBQyxDQUFDO2dCQUNyQyxJQUFJLENBQUMsU0FBUztvQkFBRSxPQUFPO2dCQUN2QixJQUFJLFNBQVMsQ0FBQyxLQUFLLEtBQUssVUFBVSxFQUFFO29CQUNuQyxRQUFRLEdBQUcsSUFBSSxDQUFDO29CQUNoQixDQUFDLElBQUksQ0FBQyxDQUFDO2lCQUNQOztvQkFDSSxNQUFNO2FBQ1g7aUJBQ0ksSUFBSSxLQUFLLENBQUMsS0FBSyxLQUFLLFNBQVMsSUFBSSxDQUFDLFNBQVMsRUFBRTtnQkFDakQsU0FBUyxHQUFHLElBQUksQ0FBQztnQkFDakIsQ0FBQyxFQUFFLENBQUM7YUFDSjtpQkFDSSxJQUFJLEtBQUssQ0FBQyxXQUFXLEVBQUUsSUFBSSxDQUFDLE1BQU0sRUFBRTtnQkFDeEMsTUFBTSxHQUFHLElBQUksQ0FBQztnQkFDZCxDQUFDLEVBQUUsQ0FBQzthQUNKO2lCQUNJLElBQUksS0FBSyxDQUFDLGNBQWMsRUFBRSxJQUFJLFFBQVEsSUFBSSxTQUFTLElBQUksTUFBTSxFQUFFO2dCQUNuRSxPQUFPLEtBQUssQ0FBQzthQUNiO2lCQUNJO2dCQUNKLENBQUMsRUFBRSxDQUFDO2FBQ0o7U0FDRDtRQUNELE9BQU87SUFDUixDQUFDO0lBRU8sa0JBQWtCLENBQUMsV0FBb0I7UUFDOUMsSUFBSSxDQUFDLEdBQUcsQ0FBQyxDQUFDO1FBQ1YsSUFBSSxpQkFBaUIsR0FBRyxLQUFLLENBQUM7UUFFOUIsSUFBSSxXQUFXLEdBQUcsRUFBRSxDQUFDO1FBRXJCLE9BQU8sQ0FBQyxHQUFHLFdBQVcsQ0FBQyxNQUFNLEVBQUU7WUFDOUIsTUFBTSxLQUFLLEdBQUcsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDO1lBRTdCLElBQUksS0FBSyxDQUFDLEtBQUssRUFBRSxJQUFJLEtBQUssQ0FBQyxPQUFPLEVBQUUsRUFBRTtnQkFDckMsQ0FBQyxFQUFFLENBQUM7Z0JBQ0osU0FBUzthQUNUO2lCQUNJLElBQUksS0FBSyxDQUFDLFlBQVksRUFBRSxJQUFJLENBQUMsaUJBQWlCLEVBQUU7Z0JBQ3BELE1BQU0sU0FBUyxHQUFHLFdBQVcsQ0FBQyxDQUFDLEdBQUcsQ0FBQyxDQUFDLENBQUM7Z0JBQ3JDLElBQUksQ0FBQyxTQUFTO29CQUFFLE9BQU87Z0JBQ3ZCLElBQUksU0FBUyxDQUFDLEtBQUssS0FBSyxTQUFTLEVBQUU7b0JBQ2xDLGlCQUFpQixHQUFHLElBQUksQ0FBQztvQkFDekIsQ0FBQyxJQUFJLENBQUMsQ0FBQztpQkFDUDs7b0JBQ0ksTUFBTTthQUNYO2lCQUNJLElBQUksS0FBSyxDQUFDLGNBQWMsRUFBRSxJQUFJLGlCQUFpQixFQUFFO2dCQUNyRCw0REFBNEQ7Z0JBQzVELElBQUksV0FBVyxLQUFLLEVBQUUsRUFBRTtvQkFDdkIsV0FBVyxHQUFHLEtBQUssQ0FBQyxLQUFLLENBQUM7aUJBQzFCO3FCQUNJO29CQUNKLFdBQVcsSUFBSSxDQUFDLEdBQUcsR0FBRyxLQUFLLENBQUMsS0FBSyxDQUFDLENBQUM7aUJBQ25DO2dCQUNELENBQUMsRUFBRSxDQUFDO2FBQ0o7aUJBQ0k7Z0JBQ0osQ0FBQyxFQUFFLENBQUM7YUFDSjtTQUNEO1FBQ0QsSUFBSSxXQUFXLEtBQUssRUFBRSxFQUFFO1lBQ3ZCLE9BQU8sV0FBVyxDQUFDO1NBQ25CO1FBQ0QsT0FBTztJQUNSLENBQUM7SUFFTyxxQkFBcUIsQ0FBQyxXQUFvQixFQUFFLFVBQWtCO1FBQ3JFLElBQUksVUFBVSxHQUFHLENBQUMsQ0FBQztRQUNuQixPQUFPLFVBQVUsR0FBRyxXQUFXLENBQUMsTUFBTSxFQUFFO1lBQ3ZDLE1BQU0sS0FBSyxHQUFHLFdBQVcsQ0FBQyxVQUFVLENBQUMsQ0FBQztZQUN0QyxVQUFVLEVBQUUsQ0FBQztZQUNiLElBQUksS0FBSyxDQUFDLFdBQVcsRUFBRSxFQUFFO2dCQUN4QixVQUFVLEVBQUUsQ0FBQzthQUNiO2lCQUNJLElBQUksS0FBSyxDQUFDLFlBQVksRUFBRSxFQUFFO2dCQUM5QixVQUFVLEVBQUUsQ0FBQzthQUNiO2lCQUNJLElBQUksS0FBSyxDQUFDLE9BQU8sRUFBRSxJQUFJLFVBQVUsS0FBSyxDQUFDLEVBQUU7Z0JBQzdDLE1BQU07YUFDTjtTQUNEO1FBQ0QsT0FBTyxVQUFVLENBQUM7SUFDbkIsQ0FBQztJQUVPLG9CQUFvQixDQUFDLEtBQVk7UUFDeEMsSUFBSSxLQUFLLENBQUMsSUFBSSw4QkFBc0I7WUFBRSxPQUFPLEtBQUssQ0FBQztRQUNuRCxNQUFNLFFBQVEsR0FBRyxDQUFDLFFBQVEsRUFBRSxTQUFTLEVBQUUsS0FBSyxFQUFFLFNBQVMsQ0FBQyxDQUFDO1FBQ3pELE9BQU8sUUFBUSxDQUFDLE9BQU8sQ0FBQyxLQUFLLENBQUMsS0FBSyxDQUFDLEtBQUssQ0FBQyxDQUFDLENBQUM7SUFDN0MsQ0FBQztJQUVPLGtCQUFrQixDQUFDLElBQVU7UUFDcEMsT0FBTyxJQUFJLENBQUMsSUFBSSxFQUFFLElBQUksSUFBSSxDQUFDLFdBQVcsQ0FBQyxJQUFJLEtBQUssSUFBSTtZQUFDLENBQUM7SUFDdkQsQ0FBQztJQUVPLGVBQWU7UUFDdEIsTUFBTSxXQUFXLEdBQUcsRUFBRSxDQUFDO1FBQ3ZCLE9BQU8sSUFBSSxDQUFDLElBQUksRUFBRSxJQUFJLElBQUksQ0FBQyxXQUFXLENBQUMsSUFBSSwwQkFBaUIsRUFBRTtZQUM3RCxXQUFXLENBQUMsSUFBSSxDQUFDLElBQUksQ0FBQyxXQUFXLENBQUMsQ0FBQztTQUNuQztRQUNELE9BQU8sV0FBVyxDQUFDO0lBRXBCLENBQUM7SUFFTyxrQkFBa0IsQ0FBQyxXQUFvQjtRQUM5QyxJQUFJLENBQUMsR0FBRyxDQUFDLENBQUM7UUFDVixtQkFBbUI7UUFDbkIsT0FBTyxDQUFDLEdBQUcsV0FBVyxDQUFDLE1BQU0sRUFBRTtZQUM5QixJQUFJLEtBQUssR0FBRyxXQUFXLENBQUMsQ0FBQyxDQUFDLENBQUM7WUFDM0IsSUFBSSxLQUFLLENBQUMsS0FBSyxFQUFFLElBQUksS0FBSyxDQUFDLE9BQU8sRUFBRSxFQUFFO2dCQUNyQyxDQUFDLEVBQUUsQ0FBQztnQkFDSixTQUFTO2FBQ1Q7WUFDRCxJQUFJLEtBQUssQ0FBQyxZQUFZLEVBQUUsRUFBRTtnQkFDekIsS0FBSyxHQUFHLFdBQVcsQ0FBQyxDQUFDLEdBQUcsQ0FBQyxDQUFDLENBQUM7Z0JBQzNCLElBQUksS0FBSyxJQUFJLEtBQUssQ0FBQyxLQUFLLEtBQUssYUFBYSxFQUFFO29CQUMzQyxNQUFNLE1BQU0sR0FBRyxXQUFXLENBQUMsTUFBTSxDQUFDLENBQUMsQ0FBQyxFQUFFO3dCQUNyQyxJQUFJLENBQUMsQ0FBQyxZQUFZLEVBQUU7NEJBQUUsT0FBTyxLQUFLLENBQUM7d0JBQ25DLElBQUksQ0FBQyxDQUFDLEtBQUssS0FBSyxhQUFhOzRCQUFFLE9BQU8sS0FBSyxDQUFDO3dCQUM1QyxPQUFPLENBQUMsQ0FBQyxJQUFJLHdCQUFlLElBQUksQ0FBQyxDQUFDLElBQUksc0JBQWEsQ0FBQztvQkFDckQsQ0FBQyxDQUNBLENBQUM7b0JBQ0YsTUFBTSxVQUFVLEdBQVksRUFBRSxDQUFDO29CQUMvQixNQUFNLFVBQVUsR0FBRyxNQUFNLENBQUMsU0FBUyxDQUFDLENBQUMsQ0FBQyxFQUFFLENBQUMsQ0FBQyxDQUFDLEtBQUssS0FBSyxPQUFPLENBQUMsQ0FBQztvQkFDOUQsSUFDQyxNQUFNLENBQUMsVUFBVSxHQUFHLENBQUMsQ0FBQzsyQkFDbkIsTUFBTSxDQUFDLFVBQVUsR0FBRyxDQUFDLENBQUMsQ0FBQyxLQUFLLEtBQUssR0FBRzsyQkFDcEMsTUFBTSxDQUFDLFVBQVUsR0FBRyxDQUFDLENBQUM7MkJBQ3RCLE1BQU0sQ0FBQyxVQUFVLEdBQUcsQ0FBQyxDQUFDLENBQUMsY0FBYyxFQUFFLEVBQ3pDO3dCQUNELFVBQVUsQ0FBQyxJQUFJLENBQUMsTUFBTSxDQUFDLFVBQVUsR0FBRyxDQUFDLENBQUMsQ0FBQyxDQUFDO3FCQUN4QztvQkFDRCxPQUFPO3dCQUNOLEVBQUUsRUFBRSxNQUFNLENBQUMsQ0FBQyxDQUFDO3dCQUNiLFdBQVcsRUFBRSxXQUFXLENBQUMsUUFBUTt3QkFDakMsU0FBUyxFQUFFLElBQUksQ0FBQyxxQkFBcUIsQ0FBQyxNQUFNLENBQUMsS0FBSyxDQUFDLENBQUMsQ0FBQyxDQUFDO3dCQUN0RCxLQUFLLEVBQUUsVUFBVTtxQkFDakIsQ0FBQztpQkFFRjtxQkFDSTtvQkFDSixNQUFNO2lCQUNOO2FBQ0Q7aUJBQ0k7Z0JBQ0osTUFBTTthQUNOO1NBQ0Q7UUFDRCxPQUFPO0lBQ1IsQ0FBQztJQUVPLHFCQUFxQixDQUFDLE1BQWU7UUFDNUMsT0FBTyxNQUFNLENBQUMsTUFBTSxDQUFDLENBQUMsQ0FBQyxFQUFFO1lBQ3hCLE9BQU8sQ0FBQyxDQUFDLEtBQUssS0FBSyxTQUFTLElBQUksQ0FBQyxDQUFDLEtBQUssS0FBSyxTQUFTLElBQUksQ0FBQyxDQUFDLEtBQUssS0FBSyxRQUFRLENBQUM7UUFDL0UsQ0FBQyxDQUFDLENBQUM7SUFDSixDQUFDO0lBRU8sV0FBVztRQUNsQixJQUFJLFVBQVUsR0FBRyxLQUFLLENBQUM7UUFDdkIsTUFBTSxNQUFNLEdBQVksSUFBSSxPQUFPLEVBQUUsQ0FBQztRQUN0QyxHQUFHO1lBQ0YsSUFBSSxDQUFDLElBQUksQ0FBQyxXQUFXO2dCQUFFLFNBQVM7WUFDaEMsSUFBSSxJQUFJLENBQUMsV0FBVyxDQUFDLEtBQUssRUFBRSxJQUFJLElBQUksQ0FBQyxXQUFXLENBQUMsT0FBTyxFQUFFO2dCQUFFLFNBQVM7aUJBQ2hFLElBQUksSUFBSSxDQUFDLFdBQVcsQ0FBQyxTQUFTLEVBQUU7Z0JBQUUsTUFBTTtpQkFDeEMsSUFBSSxJQUFJLENBQUMsV0FBVyxDQUFDLFdBQVcsRUFBRSxFQUFFO2dCQUN4QyxNQUFNLFNBQVMsR0FBRyxJQUFJLENBQUMsaUJBQWlCLENBQUMsTUFBTSxDQUFDLENBQUM7Z0JBQ2pELElBQUksQ0FBQyxTQUFTO29CQUFFLE9BQU8sU0FBUyxDQUFDO2dCQUNqQyxNQUFNLENBQUMsVUFBVSxHQUFHLFNBQVMsQ0FBQztnQkFDOUIsTUFBTTthQUNOO2lCQUNJLElBQUksSUFBSSxDQUFDLFdBQVcsQ0FBQyxjQUFjLEVBQUUsSUFBSSxJQUFJLENBQUMsV0FBVyxDQUFDLFNBQVMsRUFBRSxFQUFFO2dCQUMzRSxJQUFJLFVBQVUsRUFBRTtvQkFDZixNQUFNLENBQUMsU0FBUyxDQUFDLElBQUksQ0FBQyxJQUFJLENBQUMsV0FBVyxDQUFDLENBQUM7b0JBQ3hDLE1BQU0sQ0FBQyxLQUFLLEdBQUcsSUFBSSxDQUFDO29CQUNwQixNQUFNO2lCQUNOO2dCQUNELElBQUksTUFBTSxDQUFDLElBQUksS0FBSyxDQUFDLENBQUMsRUFBRTtvQkFDdkIsTUFBTSxDQUFDLElBQUksR0FBRyxJQUFJLENBQUMsV0FBVyxDQUFDLFFBQVEsQ0FBQyxJQUFJLENBQUM7aUJBQzdDO2dCQUNELE1BQU0sQ0FBQyxTQUFTLENBQUMsSUFBSSxDQUFDLElBQUksQ0FBQyxXQUFXLENBQUMsQ0FBQzthQUN4QztpQkFDSSxJQUFJLElBQUksQ0FBQyxXQUFXLENBQUMsV0FBVyxFQUFFLEVBQUU7Z0JBQ3hDLFVBQVUsR0FBRyxJQUFJLENBQUM7Z0JBQ2xCLFNBQVM7YUFDVDtpQkFDSSxJQUNKLElBQUksQ0FBQyxXQUFXLENBQUMsaUJBQWlCLEVBQUU7bUJBQ2pDLElBQUksQ0FBQyxXQUFXLENBQUMsYUFBYSxFQUFFO21CQUNoQyxJQUFJLENBQUMsV0FBVyxDQUFDLGtCQUFrQixFQUFFO21CQUNyQyxJQUFJLENBQUMsV0FBVyxDQUFDLGNBQWMsRUFBRTttQkFDakMsSUFBSSxDQUFDLFdBQVcsQ0FBQyxrQkFBa0IsRUFBRSxFQUN2QztnQkFDRCxTQUFTO2FBQ1Q7aUJBQ0ksSUFBSSxJQUFJLENBQUMsV0FBVyxDQUFDLEtBQUssS0FBSyxJQUFJO2dCQUFFLFNBQVM7aUJBQzlDLElBQUksSUFBSSxDQUFDLFdBQVcsQ0FBQyxZQUFZLEVBQUUsRUFBRTtnQkFDekMsSUFBSSxDQUFDLE1BQU0sQ0FBQyxVQUFVLEVBQUU7b0JBQ3ZCLE1BQU0sQ0FBQyxVQUFVLEdBQUcsSUFBSSxDQUFDLFdBQVcsQ0FBQztpQkFDckM7YUFDRDtpQkFDSTtnQkFDSixJQUFJLENBQUMsa0JBQWtCLHVCQUFjLENBQUM7Z0JBQ3RDLElBQUksTUFBTSxDQUFDLFNBQVMsQ0FBQyxNQUFNLEdBQUcsQ0FBQyxFQUFFO29CQUNoQyxNQUFNO2lCQUNOO2dCQUNELE9BQU8sU0FBUyxDQUFDO2FBQ2pCO1NBQ0QsUUFBUSxJQUFJLENBQUMsSUFBSSxFQUFFLEVBQUU7UUFFdEIsT0FBTyxJQUFJLENBQUMsY0FBYyxDQUFDLE1BQU0sQ0FBQyxDQUFDO0lBQ3BDLENBQUM7SUFFTyxjQUFjLENBQUMsTUFBZTtRQUNyQyxLQUFLLE1BQU0sT0FBTyxJQUFJLG1CQUFtQixFQUFFO1lBQzFDLE1BQU0sS0FBSyxHQUFHLE1BQU0sQ0FBQyxTQUFTLENBQUMsR0FBRyxDQUFDLENBQUMsQ0FBQyxFQUFFLENBQUMsQ0FBQyxDQUFDLEtBQUssQ0FBQyxXQUFXLEVBQUUsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxPQUFPLENBQUMsV0FBVyxFQUFFLENBQUMsQ0FBQztZQUM5RixJQUFJLEtBQUssR0FBRyxDQUFDLENBQUMsSUFBSSxLQUFLLElBQUksTUFBTSxDQUFDLFNBQVMsQ0FBQyxNQUFNLEdBQUcsQ0FBQyxFQUFFO2dCQUN2RCxNQUFNLENBQUMsU0FBUyxHQUFHLENBQUMsTUFBTSxDQUFDLFNBQVMsQ0FBQyxDQUFDLENBQUMsQ0FBQyxDQUFDO2dCQUN6QyxNQUFNLENBQUMsVUFBVSxHQUFHLEVBQUUsQ0FBQztnQkFDdkIsTUFBTTthQUNOO1NBQ0Q7UUFDRCxnQkFBZ0I7UUFDaEIsTUFBTSxDQUFDLEVBQUUsR0FBRyxNQUFNLENBQUMsU0FBUyxDQUFDLEdBQUcsRUFBRSxDQUFDO1FBQ25DLElBQUksSUFBSSxDQUFDLFlBQVksRUFBRTtZQUN0QixJQUFJLENBQUMsWUFBWSxDQUFDLE9BQU8sR0FBRyxNQUFNLENBQUMsRUFBRSxDQUFDLFFBQVEsQ0FBQyxJQUFJLEdBQUcsQ0FBQyxDQUFDO1NBQ3hEO1FBQ0QsTUFBTSxZQUFZLEdBQUcsTUFBTSxDQUFDLFNBQVMsQ0FBQyxNQUFNLENBQUMsU0FBUyxDQUFDLE1BQU0sR0FBRyxDQUFDLENBQUMsQ0FBQztRQUNuRSxJQUFJLFlBQVksSUFBSSwwQkFBa0IsQ0FBQyxPQUFPLENBQUMsWUFBWSxDQUFDLEtBQUssQ0FBQyxHQUFHLENBQUMsRUFBRTtZQUN2RSxNQUFNLENBQUMsS0FBSyxHQUFHLENBQUMsTUFBTSxDQUFDLFNBQVMsQ0FBQyxHQUFHLEVBQUUsQ0FBQyxDQUFDO1NBQ3hDO1FBQ0QsSUFBSSxDQUFDLFlBQVksR0FBRyxNQUFNLENBQUM7UUFDM0IsT0FBTyxNQUFNLENBQUM7SUFDZixDQUFDO0lBRU8saUJBQWlCLENBQUMsTUFBZTtRQUV4QyxNQUFNLElBQUksR0FBaUIsRUFBRSxDQUFDO1FBQzlCLElBQUksS0FBNkIsQ0FBQztRQUNsQyxJQUFJLElBQUksR0FBRyxLQUFLLENBQUM7UUFDakIsT0FBTyxJQUFJLENBQUMsSUFBSSxFQUFFLEVBQUU7WUFDbkIsSUFBSSxJQUFJLENBQUMsV0FBVyxDQUFDLEtBQUssRUFBRSxJQUFJLElBQUksQ0FBQyxXQUFXLENBQUMsT0FBTyxFQUFFLElBQUksSUFBSSxDQUFDLFdBQVcsQ0FBQyxTQUFTLEVBQUU7Z0JBQUUsU0FBUztpQkFDaEcsSUFBSSxJQUFJLENBQUMsV0FBVyxDQUFDLFdBQVcsRUFBRSxFQUFFO2dCQUN4QyxJQUFJLEdBQUcsSUFBSSxDQUFDO2dCQUNaLElBQUksQ0FBQyxLQUFLO29CQUFFLE9BQU8sU0FBUyxDQUFDO2dCQUM3QixJQUFJLEtBQUssQ0FBQyxLQUFLLENBQUMsTUFBTSxLQUFLLENBQUMsSUFBSSxDQUFDLEtBQUssQ0FBQyxFQUFFLEVBQUU7b0JBQzFDLEtBQUssQ0FBQyxFQUFFLEdBQUcsS0FBSyxDQUFDLEtBQUssQ0FBQyxDQUFDLENBQUMsQ0FBQztvQkFDMUIsS0FBSyxDQUFDLEtBQUssQ0FBQyxDQUFDLENBQUMsR0FBRyxJQUFJLENBQUMsUUFBUSxFQUFFLENBQUM7aUJBQ2pDO2dCQUNELE1BQU0sVUFBVSxHQUFHLElBQUksQ0FBQyxpQkFBaUIsRUFBRSxDQUFDO2dCQUM1QyxJQUFJLENBQUMsVUFBVTtvQkFBRSxPQUFPLFNBQVMsQ0FBQztnQkFDbEMsS0FBSyxDQUFDLEtBQUssR0FBRyxLQUFLLENBQUMsS0FBSyxDQUFDLE1BQU0sQ0FBQyxVQUFVLENBQUMsQ0FBQztnQkFDN0MsU0FBUzthQUNUO2lCQUNJLElBQUksSUFBSSxDQUFDLFdBQVcsQ0FBQyxZQUFZLEVBQUUsRUFBRTtnQkFDekMsSUFBSSxHQUFHLEtBQUssQ0FBQztnQkFDYixNQUFNLENBQUMsVUFBVSxHQUFHLElBQUksQ0FBQyxXQUFXLENBQUM7Z0JBQ3JDLElBQUksQ0FBQyxLQUFLO29CQUFFLE1BQU07Z0JBQ2xCLElBQUksS0FBSyxDQUFDLEtBQUssQ0FBQyxNQUFNLEtBQUssQ0FBQyxJQUFJLENBQUMsS0FBSyxDQUFDLEVBQUUsRUFBRTtvQkFDMUMsS0FBSyxDQUFDLEVBQUUsR0FBRyxLQUFLLENBQUMsS0FBSyxDQUFDLENBQUMsQ0FBQyxDQUFDO29CQUMxQixLQUFLLENBQUMsS0FBSyxDQUFDLENBQUMsQ0FBQyxHQUFHLElBQUksQ0FBQyxRQUFRLEVBQUUsQ0FBQztpQkFDakM7Z0JBQ0QsSUFBSSxDQUFDLElBQUksQ0FBQyxLQUFLLENBQUMsQ0FBQztnQkFDakIsTUFBTTthQUNOO2lCQUNJLElBQUksSUFBSSxDQUFDLFdBQVcsQ0FBQyxjQUFjLEVBQUUsRUFBRTtnQkFDM0MsSUFBSSxDQUFDLEtBQUs7b0JBQUUsS0FBSyxHQUFHLElBQUksVUFBVSxFQUFFLENBQUM7Z0JBQ3JDLHNDQUFzQztnQkFDdEMsSUFBSSxJQUFJLENBQUMsV0FBVyxDQUFDLEtBQUssS0FBSyxLQUFLLEVBQUU7b0JBQ3JDLEtBQUssQ0FBQyxTQUFTLENBQUMsSUFBSSxDQUFDLElBQUksQ0FBQyxXQUFXLENBQUMsQ0FBQztvQkFDdkMsS0FBSyxDQUFDLEdBQUcsR0FBRyxJQUFJLENBQUM7aUJBQ2pCO3FCQUNJLElBQUksSUFBSSxDQUFDLFdBQVcsQ0FBQyxLQUFLLEtBQUssS0FBSyxFQUFFO29CQUMxQyxLQUFLLENBQUMsU0FBUyxDQUFDLElBQUksQ0FBQyxJQUFJLENBQUMsV0FBVyxDQUFDLENBQUM7b0JBQ3ZDLEtBQUssQ0FBQyxHQUFHLEdBQUcsSUFBSSxDQUFDO2lCQUNqQjtxQkFDSSxJQUFJLElBQUksQ0FBQyxXQUFXLENBQUMsS0FBSyxLQUFLLFNBQVMsRUFBRTtvQkFDOUMsS0FBSyxDQUFDLFNBQVMsQ0FBQyxJQUFJLENBQUMsSUFBSSxDQUFDLFdBQVcsQ0FBQyxDQUFDO29CQUN2QyxLQUFLLENBQUMsT0FBTyxHQUFHLElBQUksQ0FBQztpQkFDckI7cUJBQ0ksSUFBSSxDQUFDLEtBQUssQ0FBQyxLQUFLO29CQUFFLEtBQUssQ0FBQyxLQUFLLEdBQUcsQ0FBQyxJQUFJLENBQUMsV0FBVyxDQUFDLENBQUM7cUJBQ25EO29CQUNKLEtBQUssQ0FBQyxFQUFFLEdBQUcsSUFBSSxDQUFDLFdBQVcsQ0FBQztpQkFDNUI7YUFDRDtpQkFDSSxJQUFJLElBQUksQ0FBQyxXQUFXLENBQUMsYUFBYSxFQUFFLEVBQUU7Z0JBQzFDLElBQUksS0FBSyxFQUFFO29CQUNWLEtBQUssQ0FBQyxPQUFPLEdBQUcsSUFBSSxDQUFDLFdBQVcsQ0FBQztpQkFDakM7cUJBQ0ksSUFBSSxJQUFJLENBQUMsTUFBTSxJQUFJLENBQUMsRUFBRTtvQkFDMUIsSUFBSSxDQUFDLElBQUksQ0FBQyxNQUFNLEdBQUcsQ0FBQyxDQUFDLENBQUMsT0FBTyxHQUFHLElBQUksQ0FBQyxXQUFXLENBQUM7aUJBQ2pEO2FBQ0Q7aUJBQ0ksSUFBSSxJQUFJLENBQUMsV0FBVyxDQUFDLE9BQU8sRUFBRSxFQUFFO2dCQUNwQyxJQUFJLENBQUMsS0FBSztvQkFBRSxPQUFPLFNBQVMsQ0FBQztnQkFDN0IsSUFBSSxLQUFLLENBQUMsS0FBSyxDQUFDLE1BQU0sS0FBSyxDQUFDLElBQUksQ0FBQyxLQUFLLENBQUMsRUFBRSxFQUFFO29CQUMxQyxLQUFLLENBQUMsRUFBRSxHQUFHLEtBQUssQ0FBQyxLQUFLLENBQUMsQ0FBQyxDQUFDLENBQUM7b0JBQzFCLEtBQUssQ0FBQyxLQUFLLENBQUMsQ0FBQyxDQUFDLEdBQUcsSUFBSSxDQUFDLFFBQVEsRUFBRSxDQUFDO2lCQUNqQztnQkFDRCxJQUFJLENBQUMsSUFBSSxDQUFDLEtBQUssQ0FBQyxDQUFDO2dCQUNqQixLQUFLLEdBQUcsU0FBUyxDQUFDO2FBQ2xCO1NBQ0Q7UUFDRCxJQUFJLElBQUk7WUFBRSxPQUFPLFNBQVMsQ0FBQztRQUMzQixPQUFPLElBQUksQ0FBQztJQUNiLENBQUM7SUFFTyxpQkFBaUI7UUFDeEIsTUFBTSxLQUFLLEdBQVksRUFBRSxDQUFDO1FBQzFCLElBQUksS0FBSyxHQUFHLEtBQUssQ0FBQztRQUNsQixPQUFPLElBQUksQ0FBQyxJQUFJLEVBQUUsRUFBRTtZQUNuQixNQUFNLEtBQUssR0FBRyxJQUFJLENBQUMsUUFBUSxFQUFFLENBQUM7WUFDOUIsSUFBSSxJQUFJLENBQUMsV0FBVyxDQUFDLEtBQUssRUFBRSxJQUFJLElBQUksQ0FBQyxXQUFXLENBQUMsT0FBTyxFQUFFO2dCQUFFLFNBQVM7aUJBQ2hFLElBQUksSUFBSSxDQUFDLFdBQVcsQ0FBQyxZQUFZLEVBQUUsRUFBRTtnQkFDekMsSUFBSSxLQUFLLENBQUMsTUFBTSxLQUFLLENBQUM7b0JBQUUsS0FBSyxDQUFDLElBQUksQ0FBQyxLQUFLLENBQUMsQ0FBQztnQkFDMUMsT0FBTyxLQUFLLENBQUM7YUFDYjtpQkFDSSxJQUFJLElBQUksQ0FBQyxXQUFXLENBQUMsY0FBYyxFQUFFLEVBQUU7Z0JBQzNDLElBQUksQ0FBQyxLQUFLLEVBQUU7b0JBQ1gsS0FBSyxDQUFDLElBQUksQ0FBQyxJQUFJLENBQUMsV0FBVyxDQUFDLENBQUM7b0JBQzdCLEtBQUssR0FBRyxJQUFJLENBQUM7aUJBQ2I7O29CQUNJLE9BQU8sU0FBUyxDQUFDO2FBQ3RCO2lCQUNJLElBQUksSUFBSSxDQUFDLFdBQVcsQ0FBQyxPQUFPLEVBQUUsRUFBRTtnQkFDcEMsSUFBSSxDQUFDLEtBQUssRUFBRTtvQkFDWCxJQUFJLEtBQUssQ0FBQyxNQUFNLEtBQUssQ0FBQyxFQUFFO3dCQUN2QixLQUFLLENBQUMsSUFBSSxDQUFDLEtBQUssQ0FBQyxDQUFDO3FCQUNsQjtvQkFDRCxLQUFLLENBQUMsSUFBSSxDQUFDLEtBQUssQ0FBQyxDQUFDO2lCQUNsQjtnQkFDRCxLQUFLLEdBQUcsS0FBSyxDQUFDO2dCQUNkLFNBQVM7YUFDVDtTQUNEO1FBQ0QsT0FBTyxTQUFTLENBQUM7SUFDbEIsQ0FBQztJQUVPLHFCQUFxQixDQUFDLFdBQW9CO1FBQ2pELE1BQU0sZUFBZSxHQUFHLElBQUksaUNBQWUsQ0FBQyxXQUFXLENBQUMsQ0FBQztRQUN6RCxJQUFJO1lBQ0gsT0FBTyxlQUFlLENBQUMsU0FBUyxFQUFFLENBQUM7U0FDbkM7UUFDRCxXQUFNO1lBQ0wsT0FBTyxFQUFFLENBQUM7U0FDVjtJQUNGLENBQUM7SUFFTyxRQUFRO1FBQ2YsT0FBTyxJQUFJLGlCQUFLLDBCQUFpQixFQUFFLEVBQUUsSUFBSSxDQUFDLFdBQVcsQ0FBQyxRQUFRLENBQUMsQ0FBQztJQUNqRSxDQUFDO0NBQ0QifQ==