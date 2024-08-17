"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.Position = exports.Range = exports.Token = exports.getTokens = void 0;
function* getTokens(documentContents) {
    const t = new Tokenizer();
    for (const char of documentContents) {
        t.charType = getType(char);
        if (t.tokenType === 0) {
            t.tokenType = t.charType;
        }
        while (!t.parsed) {
            if (t.parseCharacter(char)) {
                yield t.token;
            }
        }
        t.parsed = false;
    }
    if (t.tokenType !== 0) { // if there is an unfinished token left
        t.finalizeToken(0);
        yield t.token;
    }
}
exports.getTokens = getTokens;
class Token {
    constructor(type, value, position) {
        this.position = position;
        this.value = value;
        this.type = type;
    }
    getRange() {
        const startPosition = this.position;
        const endPosition = { line: this.position.line, character: this.position.character + this.value.length };
        return new Range(startPosition, endPosition);
    }
    isWhiteSpace() {
        return this.type === 32 /* Type.Space */
            || this.type === 11 /* Type.Tab */
            || this.type === 13 /* Type.NewLine */
            || this.type === -1 /* Type.Undefined */;
    }
    isAlphanumeric() {
        return this.type === 1 /* Type.Alphanumeric */;
    }
    isNumeric() {
        return this.type === 2 /* Type.Numeric */;
    }
    isLineComment() {
        return this.type === 3 /* Type.LineComment */;
    }
    isBlockComment() {
        return this.type === 4 /* Type.BlockComment */;
    }
    isString() {
        return this.type === 5 /* Type.String */;
    }
    isLineCommentInit() {
        return this.type === 6 /* Type.LineCommentInit */;
    }
    isBlockCommentInit() {
        return this.type === 7 /* Type.BlockCommentInit */;
    }
    isBlockCommentTerm() {
        return this.type === 8 /* Type.BlockCommentTerm */;
    }
    isDoubleQuotes() {
        return this.type === 9 /* Type.DoubleQuotes */;
    }
    isSlash() {
        return this.type === 10 /* Type.Slash */;
    }
    isTab() {
        return this.type === 11 /* Type.Tab */;
    }
    isNewLine() {
        return this.type === 13 /* Type.NewLine */;
    }
    isSpace() {
        return this.type === 32 /* Type.Space */;
    }
    isExclamationMark() {
        return this.type === 33 /* Type.ExclamationMark */;
    }
    isNumberSign() {
        return this.type === 35 /* Type.NumberSign */;
    }
    isDollarSign() {
        return this.type === 36 /* Type.DollarSign */;
    }
    isAmpersand() {
        return this.type === 38 /* Type.Ampersand */;
    }
    isSingleQuote() {
        return this.type === 39 /* Type.SingleQuote */;
    }
    isOpenParen() {
        return this.type === 40 /* Type.OpenParen */;
    }
    isCloseParen() {
        return this.type === 41 /* Type.CloseParen */;
    }
    isAsterisk() {
        return this.type === 42 /* Type.Asterisk */;
    }
    isPlusSign() {
        return this.type === 43 /* Type.PlusSign */;
    }
    isComma() {
        return this.type === 44 /* Type.Comma */;
    }
    isMinusSign() {
        return this.type === 45 /* Type.MinusSign */;
    }
    isPeriod() {
        return this.type === 46 /* Type.Period */;
    }
    isColon() {
        return this.type === 58 /* Type.Colon */;
    }
    isSemiColon() {
        return this.type === 59 /* Type.SemiColon */;
    }
    isLessThan() {
        return this.type === 60 /* Type.LessThan */;
    }
    isEqualSign() {
        return this.type === 61 /* Type.EqualSign */;
    }
    isGreaterThan() {
        return this.type === 62 /* Type.GreaterThan */;
    }
    isQuestionMark() {
        return this.type === 63 /* Type.QuestionMark */;
    }
    isAtSymbol() {
        return this.type === 64 /* Type.AtSymbol */;
    }
    isOpenBracket() {
        return this.type === 91 /* Type.OpenBracket */;
    }
    isBackslash() {
        return this.type === 92 /* Type.Backslash */;
    }
    isCloseBracket() {
        return this.type === 93 /* Type.CloseBracket */;
    }
    isCaret() {
        return this.type === 94 /* Type.Caret */;
    }
    isUnderscore() {
        return this.type === 95 /* Type.Underscore */;
    }
    isBackQuote() {
        return this.type === 96 /* Type.BackQuote */;
    }
    isOpenBrace() {
        return this.type === 123 /* Type.OpenBrace */;
    }
    isPipe() {
        return this.type === 124 /* Type.Pipe */;
    }
    isCloseBrace() {
        return this.type === 125 /* Type.CloseBrace */;
    }
    isTilde() {
        return this.type === 126 /* Type.Tilde */;
    }
}
exports.Token = Token;
class Range {
    constructor(a, b, c, d) {
        if (typeof a === 'number' && typeof b === 'number' && typeof c === 'number' && typeof d === 'number') {
            this.start = new Position(a, b);
            this.end = new Position(c, d);
        }
        else {
            this.start = a;
            this.end = b;
        }
    }
}
exports.Range = Range;
class Position {
    /**
     * @param line A zero-based line value.
     * @param character A zero-based character value.
     */
    constructor(line, character) {
        this.line = line;
        this.character = character;
    }
}
exports.Position = Position;
class Tokenizer {
    constructor() {
        this.documentLine = 0;
        this.documentColumn = 0;
        this.charType = 0;
        this.tokenType = 0;
        this.tokenValue = '';
        this.tokenPosition = { line: this.documentLine, character: this.documentColumn };
        this.parsed = false;
        this.stringOpen = false;
        this.firstSlash = false;
        this.asterisk = false;
    }
    parseCharacter(char) {
        if (this.tokenType === 1 /* Type.Alphanumeric */) {
            if (this.charType === 1 /* Type.Alphanumeric */ || this.charType === 2 /* Type.Numeric */) {
                this.tokenValue = this.tokenValue + char;
                this.parsed = true;
                this.documentColumn++;
                return false;
            }
            else {
                this.finalizeToken(this.charType);
                return true;
            }
        }
        else if (this.tokenType === 2 /* Type.Numeric */) {
            if (this.charType === 2 /* Type.Numeric */) {
                this.tokenValue = this.tokenValue + char;
                this.parsed = true;
                this.documentColumn++;
                return false;
            }
            else {
                this.finalizeToken(this.charType);
                return true;
            }
        }
        else if (this.tokenType === 3 /* Type.LineComment */) {
            if (this.charType === 13 /* Type.NewLine */) {
                this.finalizeToken(13 /* Type.NewLine */);
                return true;
            }
            else {
                this.tokenValue = this.tokenValue + char;
                this.parsed = true;
                this.documentColumn++;
                return false;
            }
        }
        else if (this.tokenType === 4 /* Type.BlockComment */) {
            if (this.asterisk) { // the previous char is *
                this.asterisk = false;
                if (this.charType === 10 /* Type.Slash */) { // the last two chars are * /
                    this.finalizeToken(8 /* Type.BlockCommentTerm */);
                    this.tokenValue = this.tokenValue + '*'; // add the * that was not yet added to the token
                    this.documentColumn++;
                    return true;
                }
                else {
                    this.tokenValue = this.tokenValue + '*'; // add the * that was not yet added to the token
                    this.documentColumn++;
                }
            }
            // do not add a * to the token immediately, it could be the end of a block comment
            if (this.charType === 42 /* Type.Asterisk */) {
                this.asterisk = true;
            }
            else {
                this.tokenValue = this.tokenValue + char;
                if (this.charType === 13 /* Type.NewLine */) {
                    this.documentLine++;
                    this.documentColumn = 0;
                }
                else {
                    this.documentColumn++;
                }
            }
            this.parsed = true;
            return false;
        }
        else if (this.tokenType === 5 /* Type.String */) {
            if (this.charType === 9 /* Type.DoubleQuotes */) {
                this.finalizeToken(9 /* Type.DoubleQuotes */);
                return true;
            }
            else {
                this.tokenValue = this.tokenValue + char;
                this.parsed = true;
                if (this.charType === 13 /* Type.NewLine */) {
                    this.documentLine++;
                    this.documentColumn = 0;
                }
                else {
                    this.documentColumn++;
                }
                return false;
            }
        }
        else if (this.tokenType === 6 /* Type.LineCommentInit */) {
            this.tokenValue = this.tokenValue + char;
            this.parsed = true;
            this.documentColumn++;
            this.finalizeToken(3 /* Type.LineComment */);
            return true;
        }
        else if (this.tokenType === 7 /* Type.BlockCommentInit */) {
            this.tokenValue = this.tokenValue + char;
            this.parsed = true;
            this.documentColumn++;
            this.finalizeToken(4 /* Type.BlockComment */);
            return true;
        }
        else if (this.tokenType === 8 /* Type.BlockCommentTerm */) {
            this.tokenValue = this.tokenValue + char;
            this.parsed = true;
            this.documentColumn++;
            this.finalizeToken(0);
            return true;
        }
        else if (this.tokenType === 9 /* Type.DoubleQuotes */) {
            this.tokenValue = this.tokenValue + char;
            this.parsed = true;
            this.documentColumn++;
            if (this.stringOpen) {
                this.stringOpen = false;
                this.finalizeToken(0);
            }
            else {
                this.stringOpen = true;
                this.finalizeToken(5 /* Type.String */);
            }
            return true;
        }
        else if (this.tokenType === 10 /* Type.Slash */ || this.tokenType === 59 /* Type.SemiColon */) {
            if (this.tokenType === 59 /* Type.SemiColon */) {
                this.tokenType = 6 /* Type.LineCommentInit */;
                return false;
            }
            else if (this.firstSlash) {
                this.firstSlash = false;
                if (this.charType === 10 /* Type.Slash */) {
                    this.tokenType = 6 /* Type.LineCommentInit */;
                    return false;
                }
                else if (this.charType === 42 /* Type.Asterisk */) {
                    this.tokenType = 7 /* Type.BlockCommentInit */;
                    return false;
                }
                else {
                    this.finalizeToken(this.charType);
                    return true;
                }
            }
            else {
                this.firstSlash = true;
                this.tokenValue = this.tokenValue + char;
                this.parsed = true;
                this.documentColumn++;
                return false;
            }
        }
        else if (this.tokenType === 13 /* Type.NewLine */) {
            this.tokenValue = this.tokenValue + char;
            this.parsed = true;
            this.documentLine++;
            this.documentColumn = 0;
            this.finalizeToken(0);
            return true;
        }
        else if (this.tokenType > 10) { // all other token types
            this.tokenValue = this.tokenValue + char;
            this.parsed = true;
            this.documentColumn++;
            this.finalizeToken(0);
            return true;
        }
        else if (this.tokenType === -1) { // undefined
            this.tokenValue = this.tokenValue + char;
            this.parsed = true;
            this.documentColumn++;
            this.finalizeToken(0);
            return true;
        }
        return false;
    }
    finalizeToken(newType) {
        this.token = new Token(this.tokenType, this.tokenValue, this.tokenPosition);
        this.tokenType = newType;
        this.tokenValue = '';
        this.tokenPosition = { line: this.documentLine, character: this.documentColumn };
    }
}
function getType(c) {
    const charCode = c.charCodeAt(0);
    // Find a better way to incorporate the %
    if (charCode >= 65 && charCode <= 90 || charCode >= 97 && charCode <= 122 || charCode === 37) {
        return 1 /* Type.Alphanumeric */;
    }
    else if (charCode >= 48 && charCode <= 57) {
        return 2 /* Type.Numeric */;
    }
    else if (charCode === 34) {
        return 9 /* Type.DoubleQuotes */;
    }
    else if (charCode === 47) {
        return 10 /* Type.Slash */;
    }
    else if (charCode === 9) {
        return 11 /* Type.Tab */;
    }
    else if (charCode === 10) {
        return 13 /* Type.NewLine */;
    }
    else if (charCode === 32) {
        return 32 /* Type.Space */;
    }
    else if (charCode === 33) {
        return 33 /* Type.ExclamationMark */;
    }
    else if (charCode === 35) {
        return 35 /* Type.NumberSign */;
    }
    else if (charCode === 36) {
        return 36 /* Type.DollarSign */;
        // } else if (charCode === 37) {
        // 	return Type.PercentSign;
    }
    else if (charCode === 38) {
        return 38 /* Type.Ampersand */;
    }
    else if (charCode === 39) {
        return 39 /* Type.SingleQuote */;
    }
    else if (charCode === 40) {
        return 40 /* Type.OpenParen */;
    }
    else if (charCode === 41) {
        return 41 /* Type.CloseParen */;
    }
    else if (charCode === 42) {
        return 42 /* Type.Asterisk */;
    }
    else if (charCode === 43) {
        return 43 /* Type.PlusSign */;
    }
    else if (charCode === 44) {
        return 44 /* Type.Comma */;
    }
    else if (charCode === 45) {
        return 45 /* Type.MinusSign */;
    }
    else if (charCode === 46) {
        return 46 /* Type.Period */;
    }
    else if (charCode === 58) {
        return 58 /* Type.Colon */;
    }
    else if (charCode === 59) {
        return 59 /* Type.SemiColon */;
    }
    else if (charCode === 60) {
        return 60 /* Type.LessThan */;
    }
    else if (charCode === 61) {
        return 61 /* Type.EqualSign */;
    }
    else if (charCode === 62) {
        return 62 /* Type.GreaterThan */;
    }
    else if (charCode === 63) {
        return 63 /* Type.QuestionMark */;
    }
    else if (charCode === 64) {
        return 64 /* Type.AtSymbol */;
    }
    else if (charCode === 91) {
        return 91 /* Type.OpenBracket */;
    }
    else if (charCode === 92) {
        return 92 /* Type.Backslash */;
    }
    else if (charCode === 93) {
        return 93 /* Type.CloseBracket */;
    }
    else if (charCode === 94) {
        return 94 /* Type.Caret */;
    }
    else if (charCode === 95) {
        return 95 /* Type.Underscore */;
    }
    else if (charCode === 96) {
        return 96 /* Type.BackQuote */;
    }
    else if (charCode === 123) {
        return 123 /* Type.OpenBrace */;
    }
    else if (charCode === 124) {
        return 124 /* Type.Pipe */;
    }
    else if (charCode === 125) {
        return 125 /* Type.CloseBrace */;
    }
    else if (charCode === 126) {
        return 126 /* Type.Tilde */;
    }
    else {
        return -1 /* Type.Undefined */;
    }
}
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoidG9rZW5pemVyLmpzIiwic291cmNlUm9vdCI6IiIsInNvdXJjZXMiOlsiLi4vLi4vLi4vc3JjL3BhcnNlci90b2tlbml6ZXIudHMiXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6Ijs7O0FBQUEsUUFBZSxDQUFDLENBQUMsU0FBUyxDQUFDLGdCQUF3QjtJQUNsRCxNQUFNLENBQUMsR0FBYyxJQUFJLFNBQVMsRUFBRSxDQUFDO0lBRXJDLEtBQUssTUFBTSxJQUFJLElBQUksZ0JBQWdCLEVBQUU7UUFDcEMsQ0FBQyxDQUFDLFFBQVEsR0FBRyxPQUFPLENBQUMsSUFBSSxDQUFDLENBQUM7UUFDM0IsSUFBSSxDQUFDLENBQUMsU0FBUyxLQUFLLENBQUMsRUFBRTtZQUN0QixDQUFDLENBQUMsU0FBUyxHQUFHLENBQUMsQ0FBQyxRQUFRLENBQUM7U0FDekI7UUFDRCxPQUFPLENBQUMsQ0FBQyxDQUFDLE1BQU0sRUFBRTtZQUNqQixJQUFJLENBQUMsQ0FBQyxjQUFjLENBQUMsSUFBSSxDQUFDLEVBQUU7Z0JBQzNCLE1BQU0sQ0FBQyxDQUFDLEtBQUssQ0FBQzthQUNkO1NBQ0Q7UUFDRCxDQUFDLENBQUMsTUFBTSxHQUFHLEtBQUssQ0FBQztLQUNqQjtJQUNELElBQUksQ0FBQyxDQUFDLFNBQVMsS0FBSyxDQUFDLEVBQUUsRUFBRSx1Q0FBdUM7UUFDL0QsQ0FBQyxDQUFDLGFBQWEsQ0FBQyxDQUFDLENBQUMsQ0FBQztRQUNuQixNQUFNLENBQUMsQ0FBQyxLQUFLLENBQUM7S0FDZDtBQUNGLENBQUM7QUFuQkQsOEJBbUJDO0FBRUQsTUFBYSxLQUFLO0lBS2pCLFlBQVksSUFBVSxFQUFFLEtBQWEsRUFBRSxRQUFrQjtRQUN4RCxJQUFJLENBQUMsUUFBUSxHQUFHLFFBQVEsQ0FBQztRQUN6QixJQUFJLENBQUMsS0FBSyxHQUFHLEtBQUssQ0FBQztRQUNuQixJQUFJLENBQUMsSUFBSSxHQUFHLElBQUksQ0FBQztJQUNsQixDQUFDO0lBRUQsUUFBUTtRQUNQLE1BQU0sYUFBYSxHQUFh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