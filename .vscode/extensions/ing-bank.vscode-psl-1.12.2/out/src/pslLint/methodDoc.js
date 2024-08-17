"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.TwoEmptyLines = exports.MethodSeparator = exports.MethodDocumentation = exports.Code = void 0;
const utilities_1 = require("../parser/utilities");
const api_1 = require("./api");
var Code;
(function (Code) {
    Code[Code["ONE_EMPTY_LINE"] = 1] = "ONE_EMPTY_LINE";
    Code[Code["TWO_EMPTY_LINES"] = 2] = "TWO_EMPTY_LINES";
})(Code = exports.Code || (exports.Code = {}));
/**
 * Checks if method has a documentation block below it.
 */
class MethodDocumentation extends api_1.MethodRule {
    report(method) {
        if (method.batch)
            return [];
        const diagnostics = [];
        if (!hasBlockComment(method, this.parsedDocument)) {
            const idToken = method.id;
            const message = `Documentation missing for label "${idToken.value}".`;
            diagnostics.push(addDiagnostic(idToken, method, message, this.ruleName));
        }
        return diagnostics;
    }
}
exports.MethodDocumentation = MethodDocumentation;
class MethodSeparator extends api_1.MethodRule {
    report(method) {
        if (method.batch)
            return [];
        const diagnostics = [];
        if (!hasSeparator(method, this.parsedDocument)) {
            const idToken = method.id;
            const message = `Separator missing for label "${idToken.value}".`;
            diagnostics.push(addDiagnostic(idToken, method, message, this.ruleName));
        }
        return diagnostics;
    }
}
exports.MethodSeparator = MethodSeparator;
class TwoEmptyLines extends api_1.MethodRule {
    report(method) {
        if (method.batch)
            return [];
        const diagnostics = [];
        const idToken = method.id;
        const lineAbove = hasSeparator(method, this.parsedDocument) ?
            method.id.position.line - 2 : method.id.position.line - 1;
        if (lineAbove < 2) {
            const message = `There should be two empty lines above label "${idToken.value}".`;
            return [addDiagnostic(idToken, method, message, this.ruleName, Code.TWO_EMPTY_LINES)];
        }
        const hasOneSpaceAbove = this.profileComponent.getTextAtLine(lineAbove).trim() === '';
        const hasTwoSpacesAbove = this.profileComponent.getTextAtLine(lineAbove - 1).trim() === '';
        const hasThreeSpacesAbove = this.profileComponent.getTextAtLine(lineAbove - 2).trim() === '';
        let code;
        if (!hasTwoSpacesAbove)
            code = Code.ONE_EMPTY_LINE;
        if (!hasOneSpaceAbove)
            code = Code.TWO_EMPTY_LINES;
        // Checks two empty lines above a method
        if (!hasOneSpaceAbove || !hasTwoSpacesAbove || lineAbove <= 0) {
            const message = `There should be two empty lines above label "${idToken.value}".`;
            diagnostics.push(addDiagnostic(idToken, method, message, this.ruleName, code));
        }
        // Check more than 2 empty lines above a method
        if (hasOneSpaceAbove && hasTwoSpacesAbove && hasThreeSpacesAbove) {
            const message = `There are more than two empty lines above label "${idToken.value}".`;
            diagnostics.push(addDiagnostic(idToken, method, message, this.ruleName, code));
        }
        return diagnostics;
    }
}
exports.TwoEmptyLines = TwoEmptyLines;
function addDiagnostic(idToken, method, message, ruleName, code) {
    const range = idToken.getRange();
    const diagnostic = new api_1.Diagnostic(range, message, ruleName, api_1.DiagnosticSeverity.Information);
    diagnostic.source = 'lint';
    diagnostic.member = method;
    if (code)
        diagnostic.code = code;
    return diagnostic;
}
function hasSeparator(method, parsedDocument) {
    const nextLineCommentTokens = (0, utilities_1.getCommentsOnLine)(parsedDocument, method.id.position.line - 1);
    return nextLineCommentTokens[0] && nextLineCommentTokens[0].isLineComment();
}
function hasBlockComment(method, parsedDocument) {
    const nextLineCommentTokens = (0, utilities_1.getCommentsOnLine)(parsedDocument, (0, utilities_1.getLineAfter)(method));
    return nextLineCommentTokens[0] && nextLineCommentTokens[0].isBlockComment();
}
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoibWV0aG9kRG9jLmpzIiwic291cmNlUm9vdCI6IiIsInNvdXJjZXMiOlsiLi4vLi4vLi4vc3JjL3BzbExpbnQvbWV0aG9kRG9jLnRzIl0sIm5hbWVzIjpbXSwibWFwcGluZ3MiOiI7OztBQUVBLG1EQUFxRTtBQUNyRSwrQkFBbUU7QUFFbkUsSUFBWSxJQUdYO0FBSEQsV0FBWSxJQUFJO0lBQ2YsbURBQWtCLENBQUE7SUFDbEIscURBQW1CLENBQUE7QUFDcEIsQ0FBQyxFQUhXLElBQUksR0FBSixZQUFJLEtBQUosWUFBSSxRQUdmO0FBRUQ7O0dBRUc7QUFDSCxNQUFhLG1CQUFvQixTQUFRLGdCQUFVO0lBRWxELE1BQU0sQ0FBQyxNQUFjO1FBRXBCLElBQUksTUFBTSxDQUFDLEtBQUs7WUFBRSxPQUFPLEVBQUUsQ0FBQztRQUU1QixNQUFNLFdBQVcsR0FBaUIsRUFBRSxDQUFDO1FBRXJDLElBQUksQ0FBQyxlQUFlLENBQUMsTUFBTSxFQUFFLElBQUksQ0FBQyxjQUFjLENBQUMsRUFBRTtZQUNsRCxNQUFNLE9BQU8sR0FBRyxNQUFNLENBQUMsRUFBRSxDQUFDO1lBQzFCLE1BQU0sT0FBTyxHQUFHLG9DQUFvQyxPQUFPLENBQUMsS0FBSyxJQUFJLENBQUM7WUFDdEUsV0FBVyxDQUFDLElBQUksQ0FBQyxhQUFhLENBQUMsT0FBTyxFQUFFLE1BQU0sRUFBRSxPQUFPLEVBQUUsSUFBSSxDQUFDLFFBQVEsQ0FBQyxDQUFDLENBQUM7U0FDekU7UUFFRCxPQUFPLFdBQVcsQ0FBQztJQUNwQixDQUFDO0NBQ0Q7QUFoQkQsa0RBZ0JDO0FBQ0QsTUFBYSxlQUFnQixTQUFRLGdCQUFVO0lBRTlDLE1BQU0sQ0FBQyxNQUFjO1FBRXBCLElBQUksTUFBTSxDQUFDLEtBQUs7WUFBRSxPQUFPLEVBQUUsQ0FBQztRQUU1QixNQUFNLFdBQVcsR0FBaUIsRUFBRSxDQUFDO1FBRXJDLElBQUksQ0FBQyxZQUFZLENBQUMsTUFBTSxFQUFFLElBQUksQ0FBQyxjQUFjLENBQUMsRUFBRTtZQUMvQyxN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