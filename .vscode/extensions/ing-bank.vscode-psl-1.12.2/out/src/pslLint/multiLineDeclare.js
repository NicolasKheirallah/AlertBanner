"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.MultiLineDeclare = void 0;
const parser_1 = require("../parser/parser");
const tokenizer_1 = require("../parser/tokenizer");
const api_1 = require("./api");
class MultiLineDeclare extends api_1.MethodRule {
    report(method) {
        const diagnostics = [];
        let reportVariable = false;
        const multiLineDeclarations = this.getMultiLineDeclarations(method.declarations);
        multiLineDeclarations.forEach((declarationsOnLine, lineNumber) => {
            const fullLine = this.profileComponent.getTextAtLine(lineNumber);
            if (!(fullLine.includes('=') && fullLine.includes(',')))
                return;
            for (const declaration of declarationsOnLine) {
                reportVariable = false;
                let conditionOpen = false;
                let commaFound = false;
                let conditionClose = false;
                let typePresent = false;
                for (const token of (0, tokenizer_1.getTokens)(fullLine)) {
                    if (token.isWhiteSpace())
                        continue;
                    if (token.isDoubleQuotes())
                        continue;
                    if (token.isBlockCommentInit())
                        continue;
                    if ((token.value) === 'type') {
                        typePresent = true;
                        continue;
                    }
                    if (parser_1.NON_TYPE_MODIFIERS.indexOf(token.value) > -1) {
                        continue;
                    }
                    if (declaration.types.map(t => t.value).indexOf(token.value) > -1) {
                        continue;
                    }
                    if (token.isOpenParen()) {
                        conditionOpen = true;
                        conditionClose = false;
                        continue;
                    }
                    if (conditionOpen && token.isCloseParen()) {
                        conditionClose = true;
                        continue;
                    }
                    if (token.isComma()) {
                        commaFound = true;
                        continue;
                    }
                    if (commaFound && token.isEqualSign() && typePresent && conditionOpen === conditionClose) {
                        conditionOpen = false;
                        conditionClose = false;
                        commaFound = false;
                        reportVariable = true;
                    }
                }
                if (reportVariable) {
                    const diagnostic = new api_1.Diagnostic(declaration.id.getRange(), `Declaration ${declaration.id.value} should be initialized on a new line.`, this.ruleName, api_1.DiagnosticSeverity.Warning);
                    diagnostic.source = 'lint';
                    diagnostics.push(diagnostic);
                }
            }
        });
        return diagnostics;
    }
    getMultiLineDeclarations(declarations) {
        const data = new Map();
        declarations.forEach(declaration => {
            const lineNumber = declaration.id.position.line;
            const declarationsOnLine = data.get(lineNumber);
            if (declarationsOnLine)
                declarationsOnLine.push(declaration);
            else
                data.set(lineNumber, [declaration]);
        });
        data.forEach((declarationArray, lineNumber) => {
            if (declarationArray.length <= 1) {
                data.delete(lineNumber);
            }
        });
        return data;
    }
}
exports.MultiLineDeclare = MultiLineDeclare;
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoibXVsdGlMaW5lRGVjbGFyZS5qcyIsInNvdXJjZVJvb3QiOiIiLCJzb3VyY2VzIjpbIi4uLy4uLy4uL3NyYy9wc2xMaW50L211bHRpTGluZURlY2xhcmUudHMiXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6Ijs7O0FBQUEsNkNBQTJFO0FBQzNFLG1EQUFnRDtBQUNoRCwrQkFBbUU7QUFFbkUsTUFBYSxnQkFBaUIsU0FBUSxnQkFBVTtJQUUvQyxNQUFNLENBQUMsTUFBYztRQUVwQixNQUFNLFdBQVcsR0FBaUIsRUFBRSxDQUFDO1FBQ3JDLElBQUksY0FBYyxHQUFZLEtBQUssQ0FBQztRQUVwQyxNQUFNLHFCQUFxQixHQUFHLElBQUksQ0FBQyx3QkFBd0IsQ0FBQyxNQUFNLENBQUMsWUFBWSxDQUFDLENBQUM7UUFFakYscUJBQXFCLENBQUMsT0FBTyxDQUFDLENBQUMsa0JBQWtCLEVBQUUsVUFBVSxFQUFFLEVBQUU7WUFDaEUsTUFBTSxRQUFRLEdBQUcsSUFBSSxDQUFDLGdCQUFnQixDQUFDLGFBQWEsQ0FBQyxVQUFVLENBQUMsQ0FBQztZQUNqRSxJQUFJLENBQUMsQ0FBQyxRQUFRLENBQUMsUUFBUSxDQUFDLEdBQUcsQ0FBQyxJQUFJLFFBQVEsQ0FBQyxRQUFRLENBQUMsR0FBRyxDQUFDLENBQUM7Z0JBQUUsT0FBTztZQUNoRSxLQUFLLE1BQU0sV0FBVyxJQUFJLGtCQUFrQixFQUFFO2dCQUM3QyxjQUFjLEdBQUcsS0FBSyxDQUFDO2dCQUN2QixJQUFJLGFBQWEsR0FBWSxLQUFLLENBQUM7Z0JBQ25DLElBQUksVUFBVSxHQUFZLEtBQUssQ0FBQztnQkFDaEMsSUFBSSxjQUFjLEdBQVksS0FBSyxDQUFDO2dCQUNwQyxJQUFJLFdBQVcsR0FBWSxLQUFLLENBQUM7Z0JBQ2pDLEtBQUssTUFBTSxLQUFLLElBQUksSUFBQSxxQkFBUyxFQUFDLFFBQVEsQ0FBQyxFQUFFO29CQUN4QyxJQUFJLEtBQUssQ0FBQyxZQUFZLEVBQUU7d0JBQUUsU0FBUztvQkFDbkMsSUFBSSxLQUFL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