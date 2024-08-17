"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.PSLActionProvider = void 0;
const vscode = require("vscode");
const utilities_1 = require("../parser/utilities");
const methodDoc_1 = require("../pslLint/methodDoc");
function initializeAction(title, ...diagnostics) {
    const action = new vscode.CodeAction(title, vscode.CodeActionKind.QuickFix);
    action.edit = new vscode.WorkspaceEdit();
    if (diagnostics)
        action.diagnostics = diagnostics;
    return action;
}
class PSLActionProvider {
    provideCodeActions(document, _range, context, _token) {
        return __awaiter(this, void 0, void 0, function* () {
            if (context.diagnostics.length === 0)
                return;
            const newLine = document.eol === vscode.EndOfLine.LF ? '\n' : '\r\n';
            const actions = [];
            const allDiagnostics = [];
            const allTextEdits = [];
            const fixAll = initializeAction('Fix all.');
            for (const diagnostic of context.diagnostics) {
                if (!diagnostic.member)
                    continue;
                const method = diagnostic.member;
                if (diagnostic.ruleName === methodDoc_1.MethodSeparator.name) {
                    const separatorAction = initializeAction('Add separator.', diagnostic);
                    const textEdit = vscode.TextEdit.insert(new vscode.Position(method.id.position.line - 1, Number.MAX_VALUE), `${newLine}\t// ---------------------------------------------------------------------`);
                    separatorAction.edit.set(document.uri, [textEdit]);
                    actions.push(separatorAction);
                    allDiagnostics.push(diagnostic);
                    allTextEdits.push({ edit: textEdit, priority: 2 });
                }
                if (diagnostic.ruleName === methodDoc_1.MethodDocumentation.name) {
                    const documentationAction = initializeAction('Add documentation block.', diagnostic);
                    let docText = `\t/* DOC -----------------------------------------------------------------${newLine}\t`
                        + `TODO: description of label ${method.id.value}${newLine}${newLine}`;
                    const terminator = `\t** ENDDOC */${newLine}`;
                    if (method.parameters.length > 0) {
                        const spacing = method.parameters.slice().sort((p1, p2) => {
                            return p2.id.value.length - p1.id.value.length;
                        })[0].id.value.length + 2;
                        docText += method.parameters.map(p => {
                            return `\t@param ${p.id.value}${' '.repeat(spacing - p.id.value.length)}TODO: description of param ${p.id.value}`;
                        }).join(`${newLine}${newLine}`) + `${newLine}`;
                    }
                    docText += terminator;
                    const textEdit = vscode.TextEdit.insert(new vscode.Position((0, utilities_1.getLineAfter)(method), 0), docText);
                    documentationAction.edit.set(document.uri, [textEdit]);
                    actions.push(documentationAction);
                    allDiagnostics.push(diagnostic);
                    allTextEdits.push({ edit: textEdit, priority: 2 });
                }
            }
            if (actions.length > 1) {
                fixAll.edit.set(document.uri, allTextEdits.sort((a, b) => a.priority - b.priority).map(edits => edits.edit));
                fixAll.diagnostics = allDiagnostics;
                actions.push(fixAll);
            }
            return actions;
        });
    }
}
exports.PSLActionProvider = PSLActionProvider;
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoiY29kZUFjdGlvbi5qcyIsInNvdXJjZVJvb3QiOiIiLCJzb3VyY2VzIjpbIi4uLy4uLy4uL3NyYy9sYW5ndWFnZS9jb2RlQWN0aW9uLnRzIl0sIm5hbWVzIjpbXSwibWFwcGluZ3MiOiI7Ozs7Ozs7Ozs7OztBQUFBLGlDQUFpQztBQUdqQyxtREFBbUQ7QUFDbkQsb0RBQTRFO0FBRTVFLFNBQVMsZ0JBQWdCLENBQUMsS0FBYSxFQUFFLEdBQUcsV0FBK0I7SUFDMUUsTUFBTSxNQUFNLEdBQUcsSUFBSSxNQUFNLENBQUMsVUFBVSxDQUFDLEtBQUssRUFBRSxNQUFNLENBQUMsY0FBYyxDQUFDLFFBQVEsQ0FBQyxDQUFDO0lBQzVFLE1BQU0sQ0FBQyxJQUFJLEdBQUcsSUFBSSxNQUFNLENBQUMsYUFBYSxFQUFFLENBQUM7SUFDekMsSUFBSSxXQUFXO1FBQUUsTUFBTSxDQUFDLFdBQVcsR0FBRyxXQUFXLENBQUM7SUFDbEQsT0FBTyxNQUFNLENBQUM7QUFDZixDQUFDO0FBa0JELE1BQWEsaUJBQWlCO0lBQ2hCLGtCQUFrQixDQUM5QixRQUE2QixFQUM3QixNQUF1QyxFQUN2QyxPQUFpQyxFQUNqQyxNQUFnQzs7WUFHaEMsSUFBSSxPQUFPLENBQUMsV0FBVyxDQUFDLE1BQU0sS0FBSyxDQUFDO2dCQUFFLE9BQU87WUFFN0MsTUFBTSxPQUFPLEdBQUcsUUFBUSxDQUFDLEdBQUcsS0FBSyxNQUFNLENBQUMsU0FBUyxDQUFDLEVBQUUsQ0FBQyxDQUFDLENBQUMsSUFBSSxDQUFDLENBQUMsQ0FBQyxNQUFNLENBQUM7WUFDckUsTUFBTSxPQUFPLEdBQXdCLEVBQUUsQ0FBQztZQUN4QyxNQUFNLGNBQWMsR0FBdUIsRUFBRSxDQUFDO1lBQzlDLE1BQU0sWUFBWSxHQUF1RCxFQUFFLENBQUM7WUFFNUUsTUFBTSxNQUFNLEdBQXNCLGdCQUFnQixDQUFDLFVBQVUsQ0FBQyxDQUFDO1lBRS9ELEtBQUssTUFBTSxVQUFVLElBQUksT0FBTyxDQUFDLFdBQVcsRUFBRTtnQkFDN0MsSUFBSSxD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