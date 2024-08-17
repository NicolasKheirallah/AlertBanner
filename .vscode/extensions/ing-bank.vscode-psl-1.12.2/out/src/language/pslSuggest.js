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
exports.PSLCompletionItemProvider = void 0;
const path = require("path");
const vscode = require("vscode");
const config_1 = require("../parser/config");
const parser = require("../parser/parser");
const parser_1 = require("../parser/parser");
const utils = require("../parser/utilities");
const lang = require("./lang");
class PSLCompletionItemProvider {
    provideCompletionItems(document, position, cancellationToken) {
        return __awaiter(this, void 0, void 0, function* () {
            if (cancellationToken.isCancellationRequested)
                return;
            let parsedDoc = parser.parseText(document.getText());
            // get tokens on line and current token
            let tokenSearchResults = utils.searchTokens(parsedDoc.tokens, position);
            if (!tokenSearchResults)
                return;
            let { tokensOnLine, index } = tokenSearchResults;
            const workspaceDirectory = vscode.workspace.getWorkspaceFolder(document.uri);
            if (!workspaceDirectory)
                return;
            let callTokens = utils.getCallTokens(tokensOnLine, index);
            if (callTokens.length === 0)
                return;
            let paths = (0, config_1.getFinderPaths)(workspaceDirectory.uri.fsPath, document.fileName);
            let finder = new utils.ParsedDocFinder(parsedDoc, paths, lang.getWorkspaceDocumentText);
            let result = yield finder.resolveResult(callTokens.slice(0, -1));
            let resultFinder = result.member ? yield finder.newFinder(result.member.types[0].value) : yield finder.newFinder(path.basename(result.fsPath).split('.')[0]);
            let resolvedResults = yield resultFinder.findAllInDocument();
            if (resolvedResults)
                return getCompletionItems(resolvedResults, finder);
        });
    }
    resolveCompletionItem(item) {
        return __awaiter(this, void 0, void 0, function* () {
            let { code, markdown } = yield lang.getDocumentation(item.result, item.finder);
            let clean = markdown.replace(/\s*(DOC)?\s*\-+/, '').replace(/\*+\s+ENDDOC/, '').trim();
            clean = clean
                .split(/\r?\n/g).map(l => l.trim()).join('\n')
                .replace(/(@\w+)/g, '*$1*')
                .replace(/(\*(@(param|publicnew|public|throws?))\*)\s+([A-Za-z\-0-9%_\.]+)/g, '$1 `$4`');
            item.detail = code;
            item.documentation = new vscode.MarkdownString().appendMarkdown(clean);
            return item;
        });
    }
}
exports.PSLCompletionItemProvider = PSLCompletionItemProvider;
function getCompletionItems(results, finder) {
    return __awaiter(this, void 0, void 0, function* () {
        let ret = results.map((result) => __awaiter(this, void 0, void 0, function* () {
            const item = new PSLCompletionItem(result.member.id.value);
            item.kind = result.member.memberClass === parser_1.MemberClass.method ? vscode.CompletionItemKind.Method : vscode.CompletionItemKind.Property;
            item.result = result;
            item.finder = finder;
            return item;
        }));
        return Promise.all(ret);
    });
}
class PSLCompletionItem extends vscode.CompletionItem {
}
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoicHNsU3VnZ2VzdC5qcyIsInNvdXJjZVJvb3QiOiIiLCJzb3VyY2VzIjpbIi4uLy4uLy4uL3NyYy9sYW5ndWFnZS9wc2xTdWdnZXN0LnRzIl0sIm5hbWVzIjpbXSwibWFwcGluZ3MiOiI7Ozs7Ozs7Ozs7OztBQUFBLDZCQUE2QjtBQUM3QixpQ0FBaUM7QUFDakMsNkNBQStEO0FBQy9ELDJDQUEyQztBQUMzQyw2Q0FBK0M7QUFDL0MsNkNBQTZDO0FBQzdDLCtCQUErQjtBQUUvQixNQUFhLHlCQUF5QjtJQUUvQixzQkFBc0IsQ0FBQyxRQUE2QixFQUFFLFFBQXlCLEVBQUUsaUJBQTJDOztZQUNqSSxJQUFJLGlCQUFpQixDQUFDLHVCQUF1QjtnQkFBRSxPQUFPO1lBQ3RELElBQUksU0FBUyxHQUFHLE1BQU0sQ0FBQyxTQUFTLENBQUMsUUFBUSxDQUFDLE9BQU8sRUFBRSxDQUFDLENBQUM7WUFFckQsdUNBQXVDO1lBQ3ZDLElBQUksa0JBQWtCLEdBQUcsS0FBSyxDQUFDLFlBQVksQ0FBQyxTQUFTLENBQUMsTUFBTSxFQUFFLFFBQVEsQ0FBQyxDQUFDO1lBQ3hFLElBQUksQ0FBQyxrQkFBa0I7Z0JBQUUsT0FBTztZQUNoQyxJQUFJLEVBQUUsWUFBWSxFQUFFLEtBQUssRUFBRSxHQUFHLGtCQUFrQixDQUFDO1lBRWpELE1BQU0sa0JBQWtCLEdBQUcsTUFBTSxDQUFDLFNBQVMsQ0FBQyxrQkFBa0IsQ0FBQyxRQUFRLENBQUMsR0FBRyxDQUFDLENBQUM7WUFDN0UsSUFBSSxDQUFDLGtCQUFrQjtnQkFBRSxPQUFPO1lBRWhDLElBQUksVUFBVSxHQUFHLEtBQUssQ0FBQyxhQUFhLENBQUMsWUFBWSxFQUFFLEtBQUssQ0FBQyxDQUFDO1lBQzFELElBQUksVUFBVSxDQUFDLE1BQU0sS0FBSyxDQUFDO2dCQUFFLE9BQU87WUFDcEMsSUFBSSxLQUFLLEdBQWdCLElBQUEsdUJBQWMsRUFBQyxrQkFBa0IsQ0FBQyxHQUFHLENBQUMsTUFBTSxFQUFFLFFBQVEsQ0FBQyxRQUFRLENBQUMs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