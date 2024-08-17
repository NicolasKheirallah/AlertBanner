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
exports.getWorkspaceDocumentText = exports.getDocumentation = void 0;
const fs = require("fs-extra");
const jsonc = require("jsonc-parser");
const path = require("path");
const vscode = require("vscode");
const parser = require("../parser/parser");
function getDocumentation(result, finder) {
    return __awaiter(this, void 0, void 0, function* () {
        const { fsPath, member } = result;
        if (!member) {
            // handle tables here
            if (fsPath.endsWith('.TBL')) {
                const text = yield getWorkspaceDocumentText(fsPath);
                const parsed = jsonc.parse(text);
                const doc = text.split('}')[1];
                const tableName = path.basename(fsPath).split('.')[0];
                return { code: '(table) ' + tableName, markdown: `${parsed.DES}\n\n${doc}` };
            }
        }
        else if (member.memberClass === parser.MemberClass.column) {
            const typs = {
                $: ['Number', 'column type: $ (Currency)'],
                B: ['String', 'column type: B (Blob)'],
                C: ['Number', 'column type: C (Time)'],
                D: ['Date', 'column type: D (Date)'],
                F: ['Number', 'column type: F (Frequency)'],
                L: ['Boolean', 'column type: L (Logical)'],
                M: ['String', 'column type: M (Memo)'],
                N: ['Number', 'column type: N (Number)'],
                T: ['String', 'column type: T (Text)'],
                U: ['String', 'column type: U (Uppercase text)'],
            };
            const text = yield getWorkspaceDocumentText(fsPath);
            const parsed = jsonc.parse(text);
            const typ = parsed.TYP;
            const doc = text.split('}')[1];
            return {
                code: `(column) ${typs[typ][0]} ${member.id.value}`,
                markdown: `${parsed.DES}\n\n${typs[typ][1]}\n\n${doc}`,
            };
        }
        else if (member.memberClass === parser.MemberClass.method) {
            const method = member;
            const sigArray = [...method.modifiers, method.types[0], method.id];
            const sig = sigArray.filter(Boolean).map(t => t.value).join(' ');
            const argString = method.parameters
                .map(param => `${param.types[0].value} ${param.id.value}`)
                .join('\n\u200B , ');
            let code = '';
            if (method.parameters.length === 0)
                code = `${sig}(${argString})`;
            else
                code = `${sig}(\n\u200B \u200B \u200B ${argString}\n\u200B )`;
            const markdown = method.documentation ? method.documentation : '';
            return { code, markdown };
        }
        else {
            let code = '';
            if (member.types.length === 0)
                code = `void ${member.id.value}`;
            else if (member.types.length === 1) {
                if (member.types[0] === member.id)
                    code = `static ${member.id.value}`;
                else
                    code = `${member.types[0].value} ${member.id.value}`;
            }
            else {
                code = `${member.types[0].value} ${member.id.value}( ${member.types.slice(1).map((t) => t.value).join(', ')})`;
            }
            switch (member.memberClass) {
                case parser.MemberClass.declaration:
                    code = ' type ' + code;
                    break;
                case parser.MemberClass.parameter:
                    code = '(parameter) ' + code;
                    break;
                case parser.MemberClass.property:
                    code = ' #PROPERTYDEF ' + code;
                    break;
                default:
                    return;
            }
            let markdown = result.member.documentation ? result.member.documentation : '';
            if (member.types[0].value.startsWith('Record')) {
                const tableName = member.types[0].value.replace('Record', '');
                const tableDirectory = yield finder.resolveFileDefinitionDirectory(tableName);
                if (tableDirectory) {
                    const tableLocation = path.join(tableDirectory, tableName.toUpperCase() + '.TBL');
                    const text = yield getWorkspaceDocumentText(tableLocation);
                    const parsed = jsonc.parse(text);
                    const doc = text.split('}')[1];
                    markdown = `${parsed.DES}\n\n${doc}`;
                }
            }
            return { code, markdown };
        }
    });
}
exports.getDocumentation = getDocumentation;
function getWorkspaceDocumentText(fsPath) {
    return __awaiter(this, void 0, void 0, function* () {
        return fs.stat(fsPath).then(_ => {
            return vscode.workspace.openTextDocument(fsPath).then(textDocument => textDocument.getText(), () => '');
        }).catch(() => '');
    });
}
exports.getWorkspaceDocumentText = getWorkspaceDocumentText;
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoibGFuZy5qcyIsInNvdXJjZVJvb3QiOiIiLCJzb3VyY2VzIjpbIi4uLy4uLy4uL3NyYy9sYW5ndWFnZS9sYW5nLnRzIl0sIm5hbWVzIjpbXSwibWFwcGluZ3MiOiI7Ozs7Ozs7Ozs7OztBQUFBLCtCQUErQjtBQUMvQixzQ0FBc0M7QUFDdEMsNkJBQTZCO0FBQzdCLGlDQUFpQztBQUNqQywyQ0FBMkM7QUFTM0MsU0FBc0IsZ0JBQWdCLENBQUMsTUFBMEIsRUFBRSxNQUE2Qjs7UUFDL0YsTUFBTSxFQUFFLE1BQU0sRUFBRSxNQUFNLEVBQUUsR0FBRyxNQUFNLENBQUM7UUFDbEMsSUFBSSxDQUFDLE1BQU0sRUFBRTtZQUNaLHFCQUFxQjtZQUNyQixJQUFJLE1BQU0sQ0FBQyxRQUFRLENBQUMsTUFBTSxDQUFDLEVBQUU7Z0JBQzVCLE1BQU0sSUFBSSxHQUFHLE1BQU0sd0JBQXdCLENBQUMsTUFBTSxDQUFDLENBQUM7Z0JBQ3BELE1BQU0sTUFBTSxHQUFHLEtBQUssQ0FBQyxLQUFLLENBQUMsSUFBSSxDQUFDLENBQUM7Z0JBQ2pDLE1BQU0sR0FBRyxHQUFHLElBQUksQ0FBQyxLQUFLLENBQUMsR0FBRyxDQUFDLENBQUMsQ0FBQyxDQUFDLENBQUM7Z0JBQy9CLE1BQU0sU0FBUyxHQUFHLElBQUksQ0FBQyxRQUFRLENBQUMsTUFBTSxDQUFDLENBQUMsS0FBSyxDQUFDLEdBQUcsQ0FBQyxDQUFDLENBQUMsQ0FBQyxDQUFDO2dCQUV0RCxPQUFPLEVBQUUsSUFBSSxFQUFFLFVBQVUsR0FBRyxTQUFTLEVBQUUsUUFBUSxFQUFFLEdBQUcsTUFBTSxDQUFDLEdBQUcsT0FBTyxHQUFHLEVBQUUsRUFBRSxDQUFDO2FBQzdFO1NBQ0Q7YUFDSSxJQUFJLE1BQU0sQ0FBQyxXQUFXLEtBQUssTUFBTSxDQUFDLFdBQVcsQ0FBQyxNQUFNLEVBQUU7WUFDMUQsTUFBTSxJQUFJLEdBQUc7Z0JBQ1osQ0FBQyxFQUFFLENBQUMsUUFBUSxFQUFFLDJCQUEyQixDQUFDO2dCQUMxQyxDQUFDLEVBQUUs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