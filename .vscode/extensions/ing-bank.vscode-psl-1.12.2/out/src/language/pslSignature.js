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
exports.PSLSignatureHelpProvider = void 0;
const vscode = require("vscode");
const config_1 = require("../parser/config");
const parser = require("../parser/parser");
const utils = require("../parser/utilities");
const lang = require("./lang");
class PSLSignatureHelpProvider {
    provideSignatureHelp(document, position) {
        return __awaiter(this, void 0, void 0, function* () {
            const workspaceDirectory = vscode.workspace.getWorkspaceFolder(document.uri);
            if (!workspaceDirectory)
                return;
            let parsedDoc = parser.parseText(document.getText());
            // get tokens on line and current token
            let tokenSearchResults = ((tokens, position) => {
                const tokensOnLine = tokens.filter(t => t.position.line === position.line);
                if (tokensOnLine.length === 0)
                    return undefined;
                const index = tokensOnLine.findIndex(t => {
                    const start = t.position;
                    const end = { line: t.position.line, character: t.position.character + t.value.length };
                    const isBetween = (lb, t, ub) => {
                        return lb.line <= t.line &&
                            lb.character <= t.character &&
                            ub.line >= t.line &&
                            ub.character >= t.character;
                    };
                    return isBetween(start, position, end);
                });
                return { tokensOnLine, index };
            })(parsedDoc.tokens, position);
            if (!tokenSearchResults)
                return;
            let { tokensOnLine, index } = tokenSearchResults;
            let { callTokens, parameterIndex } = utils.findCallable(tokensOnLine, index);
            if (callTokens.length === 0)
                return;
            let paths = (0, config_1.getFinderPaths)(workspaceDirectory.uri.fsPath, document.fileName);
            let finder = new utils.ParsedDocFinder(parsedDoc, paths, lang.getWorkspaceDocumentText);
            let resolvedResult = yield finder.resolveResult(callTokens);
            if (!resolvedResult.member || resolvedResult.member.memberClass !== parser.MemberClass.method)
                return;
            if (resolvedResult)
                return getSignature(resolvedResult, parameterIndex, finder);
        });
    }
}
exports.PSLSignatureHelpProvider = PSLSignatureHelpProvider;
function getSignature(result, parameterIndex, finder) {
    return __awaiter(this, void 0, void 0, function* () {
        let { code, markdown } = yield lang.getDocumentation(result, finder);
        let clean = markdown.replace(/\s*(DOC)?\s*\-+/, '').replace(/\*+\s+ENDDOC/, '').trim();
        clean = clean
            .split(/\r?\n/g).map(l => l.trim()).join('\n')
            .replace(/(@\w+)/g, '*$1*')
            .replace(/(\*(@(param|publicnew|public|throws?))\*)\s+([A-Za-z\-0-9%_\.]+)/g, '$1 `$4`');
        let method = result.member;
        let argString = method.parameters.map((param) => `${param.types[0].value} ${param.id.value}`).join(', ');
        code = `${method.id.value}(${argString})`;
        let info = new vscode.SignatureInformation(code, new vscode.MarkdownString().appendMarkdown(clean));
        info.parameters = method.parameters.map(parameter => new vscode.ParameterInformation(`${parameter.types[0].value} ${parameter.id.value}`));
        let signatureHelp = new vscode.SignatureHelp();
        signatureHelp.signatures = [info];
        signatureHelp.activeSignature = 0;
        signatureHelp.activeParameter = parameterIndex;
        return signatureHelp;
    });
}
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoicHNsU2lnbmF0dXJlLmpzIiwic291cmNlUm9vdCI6IiIsInNvdXJjZXMiOlsiLi4vLi4vLi4vc3JjL2xhbmd1YWdlL3BzbFNpZ25hdHVyZS50cyJdLCJuYW1lcyI6W10sIm1hcHBpbmdzIjoiOzs7Ozs7Ozs7Ozs7QUFBQSxpQ0FBaUM7QUFDakMsNkNBQStEO0FBQy9ELDJDQUEyQztBQUUzQyw2Q0FBNkM7QUFDN0MsK0JBQStCO0FBRS9CLE1BQWEsd0JBQXdCO0lBQ3ZCLG9CQUFvQixDQUFDLFFBQTZCLEVBQUUsUUFBeUI7O1lBQ3pGLE1BQU0sa0JBQWtCLEdBQUcsTUFBTSxDQUFDLFNBQVMsQ0FBQyxrQkFBa0IsQ0FBQyxRQUFRLENBQUMsR0FBRyxDQUFDLENBQUM7WUFDN0UsSUFBSSxDQUFDLGtCQUFrQjtnQkFBRSxPQUFPO1lBRWhDLElBQUksU0FBUyxHQUFHLE1BQU0sQ0FBQyxTQUFTLENBQUMsUUFBUSxDQUFDLE9BQU8sRUFBRSxDQUFDLENBQUM7WUFDckQsdUNBQXVDO1lBQ3ZDLElBQUksa0JBQWtCLEdBQUcsQ0FBQyxDQUFDLE1BQWUsRUFBRSxRQUFrQixFQUFFLEVBQUU7Z0JBQ2pFLE1BQU0sWUFBWSxHQUFHLE1BQU0sQ0FBQyxNQUFNLENBQUMsQ0FBQyxDQUFDLEVBQUUsQ0FBQyxDQUFDLENBQUMsUUFBUSxDQUFDLElBQUksS0FBSyxRQUFRLENBQUMsSUFBSSxDQUFDLENBQUM7Z0JBQzNFLElBQUksWUFBWSxDQUFDLE1BQU0sS0FBSyxDQUFDO29CQUFFLE9BQU8sU0FBUyxDQUFDO2dCQUNoRCxNQUFNLEtBQUssR0FBRyxZQUFZLENBQUMsU0FBUyxDQUFDLENBQUMsQ0FBQyxFQUFFO29CQUN4QyxNQUFNLEtBQUssR0FBYSxDQUFDLENBQUMsUUFBUSxDQUFDO29CQUNuQyxNQUFNLEdBQUcsR0FBYSxFQUFFLElBQUksRUFBRSxDQUFDLENBQUMsUUFBUSxDQUFDLElBQUksRUFBRSxTQUFTLEVBQUUsQ0FBQyxDQUFDLFFBQVEsQ0FBQyxTQUFTLEdBQUcsQ0FB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