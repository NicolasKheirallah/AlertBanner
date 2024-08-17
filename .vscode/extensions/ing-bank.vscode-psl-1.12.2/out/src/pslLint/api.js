"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.ProfileComponent = exports.DeclarationRule = exports.ParameterRule = exports.MethodRule = exports.PropertyRule = exports.MemberRule = exports.PslRule = exports.FileDefinitionRule = exports.ProfileComponentRule = exports.DiagnosticRelatedInformation = exports.Diagnostic = exports.DiagnosticSeverity = void 0;
const path = require("path");
const tokenizer_1 = require("./../parser/tokenizer");
var DiagnosticSeverity;
(function (DiagnosticSeverity) {
    /**
     * Something not allowed by the rules of a language or other means.
     */
    DiagnosticSeverity[DiagnosticSeverity["Error"] = 0] = "Error";
    /**
     * Something suspicious but allowed.
     */
    DiagnosticSeverity[DiagnosticSeverity["Warning"] = 1] = "Warning";
    /**
     * Something to inform about but not a problem.
     */
    DiagnosticSeverity[DiagnosticSeverity["Information"] = 2] = "Information";
    /**
     * Something to hint to a better way of doing it, like proposing
     * a refactoring.
     */
    DiagnosticSeverity[DiagnosticSeverity["Hint"] = 3] = "Hint";
})(DiagnosticSeverity = exports.DiagnosticSeverity || (exports.DiagnosticSeverity = {}));
class Diagnostic {
    /**
     * Creates a new diagnostic object.
     *
     * @param range The range to which this diagnostic applies.
     * @param message The human-readable message.
     * @param severity The severity, default is [error](#DiagnosticSeverity.Error).
     */
    constructor(range, message, ruleName, severity, member) {
        this.range = range;
        this.message = message;
        this.ruleName = ruleName;
        if (severity)
            this.severity = severity;
        if (member)
            this.member = member;
    }
}
exports.Diagnostic = Diagnostic;
/**
 * Represents a related message and source code location for a diagnostic. This should be
 * used to point to code locations that cause or related to a diagnostics, e.g when duplicating
 * a symbol in a scope.
 */
class DiagnosticRelatedInformation {
    /**
     * Creates a new related diagnostic information object.
     *
     * @param range The range.
     * @param message The message.
     */
    constructor(range, message) {
        this.range = range;
        this.message = message;
    }
}
exports.DiagnosticRelatedInformation = DiagnosticRelatedInformation;
class ProfileComponentRule {
    constructor() {
        this.ruleName = this.constructor.name;
    }
}
exports.ProfileComponentRule = ProfileComponentRule;
class FileDefinitionRule extends ProfileComponentRule {
}
exports.FileDefinitionRule = FileDefinitionRule;
class PslRule extends ProfileComponentRule {
}
exports.PslRule = PslRule;
class MemberRule extends PslRule {
}
exports.MemberRule = MemberRule;
class PropertyRule extends PslRule {
}
exports.PropertyRule = PropertyRule;
class MethodRule extends PslRule {
}
exports.MethodRule = MethodRule;
class ParameterRule extends PslRule {
}
exports.ParameterRule = ParameterRule;
class DeclarationRule extends PslRule {
}
exports.DeclarationRule = DeclarationRule;
/**
 * A ProfileComponent contains information about a file used in Profile.
 * The file may be PSL or non-PSL (such as a TBL or COL).
 */
class ProfileComponent {
    constructor(fsPath, textDocument, getTextAtLine) {
        this.textDocument = textDocument;
        this.fsPath = fsPath;
        if (getTextAtLine)
            this.getTextAtLine = getTextAtLine;
    }
    static isPsl(fsPath) {
        return path.extname(fsPath) === '.PROC'
            || path.extname(fsPath) === '.BATCH'
            || path.extname(fsPath) === '.TRIG'
            || path.extname(fsPath).toUpperCase() === '.PSL';
    }
    static isFileDefinition(fsPath) {
        return path.extname(fsPath) === '.TBL'
            || path.extname(fsPath) === '.COL';
    }
    static isProfileComponent(fsPath) {
        return ProfileComponent.isPsl(fsPath)
            || ProfileComponent.isFileDefinition(fsPath);
    }
    /**
     * A utility method to get the text at a specified line of the document.
     * @param lineNumber The zero-based line number of the document where the text is.
     */
    getTextAtLine(lineNumber) {
        if (lineNumber < 0) {
            throw new Error('Cannot get text at negative line number.');
        }
        if (!this.indexedDocument) {
            this.indexedDocument = this.createIndexedDocument();
        }
        return this.indexedDocument.get(lineNumber) || '';
    }
    /**
     * Converts a zero-based offset to a position.
     *
     * @param offset A zero-based offset.
     * @return A valid [position](#Position).
     */
    positionAt(offset) {
        const before = this.textDocument.slice(0, offset);
        const newLines = before.match(/\n/g);
        const line = newLines ? newLines.length : 0;
        const preCharacters = before.match(/(\n|^).*$/g);
        return new tokenizer_1.Position(line, preCharacters ? preCharacters[0].length : 0);
    }
    createIndexedDocument() {
        const indexedDocument = new Map();
        let line = '';
        let index = 0;
        for (const char of this.textDocument) {
            line += char;
            if (char === '\n') {
                indexedDocument.set(index, line);
                index++;
                line = '';
            }
        }
        return indexedDocument;
    }
}
exports.ProfileComponent = ProfileComponent;
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoiYXBpLmpzIiwic291cmNlUm9vdCI6IiIsInNvdXJjZXMiOlsiLi4vLi4vLi4vc3JjL3BzbExpbnQvYXBpLnRzIl0sIm5hbWVzIjpbXSwibWFwcGluZ3MiOiI7OztBQUFBLDZCQUE2QjtBQUU3QixxREFBd0Q7QUFFeEQsSUFBWSxrQkFzQlg7QUF0QkQsV0FBWSxrQkFBa0I7SUFFN0I7O09BRUc7SUFDSCw2REFBUyxDQUFBO0lBRVQ7O09BRUc7SUFDSCxpRUFBVyxDQUFBO0lBRVg7O09BRUc7SUFDSCx5RUFBZSxDQUFBO0lBRWY7OztPQUdHO0lBQ0gsMkRBQVEsQ0FBQTtBQUNULENBQUMsRUF0Qlcsa0JBQWtCLEdBQWxCLDBCQUFrQixLQUFsQiwwQkFBa0IsUUFzQjdCO0FBRUQsTUFBYSxVQUFVO0lBd0N0Qjs7Ozs7O09BTUc7SUFDSCxZQUFZLEtBQVksRUFBRSxPQUFlLEVBQUUsUUFBZ0IsRUFBRSxRQUE2QixFQUFFLE1BQWU7UUFDMUcsSUFBSSxDQUFDLEtBQUssR0FBRyxLQUFLLENBQUM7UUFDbkIsSUFBSSxDQUFDLE9BQU8sR0FBRyxPQUFPLENBQUM7UUFDdkIsSUFBSSxDQUFDLFFBQVEsR0FBRyxRQUFRLENBQUM7UUFDekIsSUFBSSxRQUFRO1lBQUUsSUFBSSxDQUFDLFFBQVEsR0FBRyxRQUFRLENBQUM7UUFDdkMsSUFBSSxNQUFNO1lBQUUsSUFBSSxDQUFDLE1BQU0sR0FBRyxNQUFNLENBQUM7SUFDbEMsQ0FBQztDQUNEO0FBdERELGdDQXNEQztBQUVEOzs7O0dBSUc7QUFDSCxNQUFhLDRCQUE0QjtJQVl4Qzs7Ozs7T0FLRztJQUNILFlBQVksS0FBWSxFQUFFLE9BQWU7UUFDeEMsSUFBSSxDQUFDLEtBQUssR0FBRyxLQUFLLENBQUM7UUFDbkIsSUFBSSxDQUFDLE9BQU8sR0FBRyxPQUFPLENBQUM7SUFDeEIsQ0FBQztDQUNEO0FBdEJELG9FQXNCQztBQUVELE1BQXNCLG9CQUFvQjtJQUExQztRQUVVLGFBQVEsR0FBVyxJQUFJLENBQUMsV0FBVyxDQUFDLElB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