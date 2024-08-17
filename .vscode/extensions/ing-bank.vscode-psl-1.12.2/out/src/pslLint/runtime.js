"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.RuntimeStart = void 0;
const parser_1 = require("../parser/parser");
const statementParser_1 = require("../parser/statementParser");
const tokenizer_1 = require("../parser/tokenizer");
const utilities_1 = require("../parser/utilities");
const api_1 = require("./api");
class RuntimeStart extends api_1.MethodRule {
    report(method) {
        const runtimeCalls = [];
        method.statements.filter(statement => {
            return statement.action.value === 'do';
        }).forEach(statement => {
            statement.expressions.forEach(expression => {
                const dotOperator = expression;
                const classIdentifier = this.getClass(dotOperator);
                if (!classIdentifier)
                    return;
                if (classIdentifier.id.value === 'Runtime')
                    runtimeCalls.push(dotOperator);
            });
        });
        if (!runtimeCalls.length)
            return [];
        const diagnostics = [];
        this.tpFence(diagnostics, runtimeCalls, method);
        return diagnostics;
    }
    getClass(dotOperator) {
        if (dotOperator.kind !== statementParser_1.SyntaxKind.BINARY_OPERATOR)
            return;
        if (Array.isArray(dotOperator.left))
            return;
        if (!dotOperator.left || dotOperator.left.kind === statementParser_1.SyntaxKind.BINARY_OPERATOR)
            return;
        return dotOperator.left;
    }
    getMethod(dotOperator) {
        if (dotOperator.kind !== statementParser_1.SyntaxKind.BINARY_OPERATOR)
            return;
        return dotOperator.right;
    }
    tpFence(diagnostics, runtimeCalls, method) {
        let lastStart;
        let variables;
        let acceptVariables = [];
        for (const runtimeCall of runtimeCalls) {
            const runtimeMethod = this.getMethod(runtimeCall);
            if (runtimeMethod.id.value === 'start') {
                if (lastStart) {
                    variables.forEach((identifiers, variable) => {
                        this.createDiagnostic(lastStart, variable, identifiers, diagnostics);
                    });
                }
                lastStart = runtimeMethod;
                variables = new Map();
                acceptVariables = this.addToWhitelist(runtimeMethod);
            }
            else if (runtimeMethod.id.value === 'commit') {
                if (!lastStart)
                    continue;
                else {
                    const startLine = lastStart.id.position.line;
                    const commitLine = runtimeMethod.id.position.line;
                    const identifierTokens = this.getAllIdentifersInRange(this.parsedDocument.tokens, startLine, commitLine);
                    const variablesOutsideStart = method.declarations.concat(method.parameters)
                        .filter(variable => {
                        return variable.id.position.line <= startLine && acceptVariables.indexOf(variable.id.value) === -1;
                    });
                    for (const token of identifierTokens) {
                        this.addVariable(variablesOutsideStart, token, lastStart, variables);
                    }
                }
            }
        }
        if (variables) {
            variables.forEach((identifiers, variable) => {
                this.createDiagnostic(lastStart, variable, identifiers, diagnostics);
            });
        }
    }
    getAllIdentifersInRange(tokens, startLine, commitLine) {
        return tokens.filter(token => {
            return token.position.line > startLine && token.position.line < commitLine;
        });
    }
    createDiagnostic(lastStart, variable, identifiers, diagnostics) {
        const range = this.getDiagnosticRange(lastStart);
        const word = variable.memberClass === parser_1.MemberClass.parameter ? 'Parameter' : 'Declaration';
        const diag = new api_1.Diagnostic(range, `${word} "${variable.id.value}" referenced inside Runtime.start but not in variable list.`, this.ruleName, api_1.DiagnosticSeverity.Warning, variable);
        const relatedSource = new api_1.DiagnosticRelatedInformation(variable.id.getRange(), `Source of "${variable.id.value}"`);
        const relatedReferences = identifiers.map(i => {
            return new api_1.DiagnosticRelatedInformation(i.getRange(), `Reference to "${i.value}"`);
        });
        diag.relatedInformation = [
            relatedSource,
            ...relatedReferences,
        ];
        diag.source = 'tpfence';
        diagnostics.push(diag);
    }
    addVariable(localVariablesOutsideStart, identifierToken, start, variables) {
        const variable = localVariablesOutsideStart.find(v => v.id.value === identifierToken.value);
        if (variable
            && variable.id !== variable.types[0]
            && variable.modifiers.map(m => m.value).indexOf('literal') === -1) { // no static and literal
            const varList = start.args[1];
            if (!varList || varList.id.value.split(',').indexOf(variable.id.value) === -1) {
                const tokens = variables.get(variable);
                if (!tokens) {
                    variables.set(variable, [identifierToken]);
                }
                else if (tokens.indexOf(identifierToken) === -1) {
                    variables.set(variable, tokens.concat([identifierToken]));
                }
            }
        }
    }
    getDiagnosticRange(start) {
        const startPos = start.id.position.character - 'do Runtime.'.length;
        const endPos = start.closeParen.position.character + 1;
        return new tokenizer_1.Range(start.id.position.line, startPos, start.id.position.line, endPos);
    }
    addToWhitelist(runtimeMethod) {
        let acceptVariables = [];
        const commentsAbove = (0, utilities_1.getCommentsOnLine)(this.parsedDocument, runtimeMethod.id.position.line - 1);
        const whiteListComment = commentsAbove[0];
        if (!whiteListComment || !whiteListComment.isLineComment())
            return [];
        const comment = whiteListComment.value.trim();
        if (!comment.startsWith('@psl-lint.RuntimeStart'))
            return [];
        const args = comment.replace(/^@psl-lint\.RuntimeStart\s+/, '').split('=');
        for (let i = 0; i < args.length; i += 2) {
            const arg = args[i];
            const value = args[i + 1];
            if (arg === 'accept' && value) {
                const strippedValue = value.replace(/"/g, '');
                acceptVariables = strippedValue.split(',');
            }
        }
        return acceptVariables;
    }
}
exports.RuntimeStart = RuntimeStart;
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoicnVudGltZS5qcyIsInNvdXJjZVJvb3QiOiIiLCJzb3VyY2VzIjpbIi4uLy4uLy4uL3NyYy9wc2xMaW50L3J1bnRpbWUudHMiXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6Ijs7O0FBQUEsNkNBQStEO0FBQy9ELCtEQUdtQztBQUNuQyxtREFBbUQ7QUFDbkQsbURBQXdEO0FBQ3hELCtCQUFpRztBQUVqRyxNQUFhLFlBQWEsU0FBUSxnQkFBVTtJQUUzQyxNQUFNLENBQUMsTUFBYztRQUVwQixNQUFNLFlBQVksR0FBcUIsRUFBRSxDQUFDO1FBRTFDLE1BQU0sQ0FBQyxVQUFVLENBQUMsTUFBTSxDQUFDLFNBQVMsQ0FBQyxFQUFFO1lBQ3BDLE9BQU8sU0FBUyxDQUFDLE1BQU0sQ0FBQyxLQUFLLEtBQUssSUFBSSxDQUFDO1FBQ3hDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxTQUFTLENBQUMsRUFBRTtZQUN0QixTQUFTLENBQUMsV0FBVyxDQUFDLE9BQU8sQ0FBQyxVQUFVLENBQUMsRUFBRTtnQkFDMUMsTUFBTSxXQUFXLEdBQUcsVUFBNEIsQ0FBQztnQkFDakQsTUFBTSxlQUFlLEdBQUcsSUFBSSxDQUFDLFFBQVEsQ0FBQyxXQUFXLENBQUMsQ0FBQztnQkFDbkQsSUFBSSxDQUFDLGVBQWU7b0JBQUUsT0FBTztnQkFDN0IsSUFBSSxlQUFlLENBQUMsRUFBRSxDQUFDLEtBQUssS0FBSyxTQUFTO29CQUFFLFlBQVksQ0FBQyxJQUFJLENBQUMsV0FBVyxDQUFDLENBQUM7WUFDNUUsQ0FBQyxDQUFDLENBQUM7UUFDSixDQUFDLENBQUMsQ0FBQztRQUVILElBQUksQ0FBQyxZQUFZLENBQUMsTUFBTTtZQUFFLE9BQU8sRUFBRSxDQUFDO1FBRXBDLE1BQU0sV0FBVyxHQUFpQixFQUFFLENBQUM7UUFDckMsSUFBSSxDQUFDLE9BQU8sQ0FBQyxXQUFXLEVBQUUsWUFBWSxFQUFFLE1BQU0sQ0FBQyxDQUFDO1FBQ2hELE9BQU8sV0FBVyxDQUFDO0lBQ3BCLENBQUM7SUFFRCxRQUFR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