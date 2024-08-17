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
exports.PSLDefinitionProvider = void 0;
const vscode = require("vscode");
const config_1 = require("../parser/config");
const parser = require("../parser/parser");
const utils = require("../parser/utilities");
const lang = require("./lang");
class PSLDefinitionProvider {
    provideDefinition(document, position, cancellationToknen) {
        return __awaiter(this, void 0, void 0, function* () {
            if (cancellationToknen.isCancellationRequested)
                return;
            let parsedDoc = parser.parseText(document.getText());
            // get tokens on line and current token
            let tokenSearchResults = utils.searchTokens(parsedDoc.tokens, position);
            if (!tokenSearchResults)
                return [];
            let { tokensOnLine, index } = tokenSearchResults;
            const workspaceDirectory = vscode.workspace.getWorkspaceFolder(document.uri);
            if (!workspaceDirectory)
                return;
            let callTokens = utils.getCallTokens(tokensOnLine, index);
            if (callTokens.length === 0)
                return;
            let paths = (0, config_1.getFinderPaths)(workspaceDirectory.uri.fsPath, document.fileName);
            let finder = new utils.ParsedDocFinder(parsedDoc, paths, lang.getWorkspaceDocumentText);
            let resolvedResult = yield finder.resolveResult(callTokens);
            if (resolvedResult)
                return getLocation(resolvedResult);
        });
    }
}
exports.PSLDefinitionProvider = PSLDefinitionProvider;
function getLocation(result) {
    if (!result.member) {
        return new vscode.Location(vscode.Uri.file(result.fsPath), new vscode.Position(0, 0));
    }
    let range = result.member.id.getRange();
    let vscodeRange = new vscode.Range(range.start.line, range.start.character, range.end.line, range.end.character);
    return new vscode.Location(vscode.Uri.file(result.fsPath), vscodeRange);
}
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoicHNsRGVmaW5pdGlvblByb3ZpZGVyLmpzIiwic291cmNlUm9vdCI6IiIsInNvdXJjZXMiOlsiLi4vLi4vLi4vc3JjL2xhbmd1YWdlL3BzbERlZmluaXRpb25Qcm92aWRlci50cyJdLCJuYW1lcyI6W10sIm1hcHBpbmdzIjoiOzs7Ozs7Ozs7Ozs7QUFBQSxpQ0FBaUM7QUFDakMsNkNBQStEO0FBQy9ELDJDQUEyQztBQUMzQyw2Q0FBNkM7QUFDN0MsK0JBQStCO0FBRS9CLE1BQWEscUJBQXFCO0lBRTNCLGlCQUFpQixDQUFDLFFBQTZCLEVBQUUsUUFBeUIsRUFBRSxrQkFBNEM7O1lBQzdILElBQUksa0JBQWtCLENBQUMsdUJBQXVCO2dCQUFFLE9BQU87WUFDdkQsSUFBSSxTQUFTLEdBQUcsTUFBTSxDQUFDLFNBQVMsQ0FBQyxRQUFRLENBQUMsT0FBTyxFQUFFLENBQUMsQ0FBQztZQUVyRCx1Q0FBdUM7WUFDdkMsSUFBSSxrQkFBa0IsR0FBRyxLQUFLLENBQUMsWUFBWSxDQUFDLFNBQVMsQ0FBQyxNQUFNLEVBQUUsUUFBUSxDQUFDLENBQUM7WUFDeEUsSUFBSSxDQUFDLGtCQUFrQjtnQkFBRSxPQUFPLEVBQUUsQ0FBQztZQUNuQyxJQUFJLEVBQUUsWUFBWSxFQUFFLEtBQUssRUFBRSxHQUFHLGtCQUFrQixDQUFDO1lBRWpELE1BQU0sa0JBQWtCLEdBQUcsTUFBTSxDQUFDLFNBQVMsQ0FBQyxrQkFBa0IsQ0FBQyxRQUFRLENBQUMsR0FBRyxDQUFDLENBQUM7WUFDN0UsSUFBSSxDQUFDLGtCQUFrQjtnQkFBRSxPQUFPO1lBRWhDLElBQUksVUFBVSxHQUFHLEtBQUssQ0FBQyxhQUFhLENBQUMsWUFBWSxFQUFFLEtBQUssQ0FBQyxDQUFDO1lBQzFELElBQUksVUFBVSxDQUFDLE1BQU0sS0FBSyxDQUFDO2dCQUFFLE9BQU87WUFDcEMsSUFBSSxLQUFLLEdBQWdCLElBQUEsdUJBQWMsRUFBQyxrQkFBa0IsQ0FBQyxHQUFHLENBQUMsTUFBTSxFQUFFLFFBQVEsQ0FBQyxRQUFRLENBQUMsQ0FBQztZQUMxRixJQUFJLE1BQU0sR0FBRyxJQUFJLEtBQUssQ0FBQyxlQUFlLENBQUMsU0FBUyxFQUFFLEtBQUssRUFBRSxJQUFJLENBQUMsd0JBQXdCLENBQUMsQ0FBQztZQUN4RixJQUFJLGNBQWMsR0FBRyxNQUFNLE1BQU0sQ0FBQyxhQUFhLENBQUMsVUFBVSxDQUFDLENBQUM7WUFDNUQsSUFBSSxjQUFjO2dCQUFFLE9BQU8sV0FBVyxDQUFDLGNBQWMsQ0FBQyxDQUFDO1FBQ3hELENBQUM7S0FBQTtDQUNEO0FBckJELHNEQXFCQztBQUVELFNBQVMsV0FBVyxDQUFDLE1BQTBCO0lBQzlDLElBQUksQ0FBQyxNQUFNLENBQUMsTUFBTSxFQUFFO1FBQ25CLE9BQU8sSUFBSSxNQUFNLENBQUMsUUFBUSxDQUFDLE1BQU0sQ0FBQyxHQUFHLENBQUMsSUFBSSxDQUFDLE1BQU0sQ0FBQyxNQUFNLENBQUMsRUFBRSxJQUFJLE1BQU0sQ0FBQyxRQUFRLENBQUMsQ0FBQyxFQUFFLENBQUMsQ0FBQyxDQUFDLENBQUM7S0FDdEY7SUFDRCxJQUFJLEtBQUssR0FBRyxNQUFNLENBQUMsTUFBTSxDQUFDLEVBQUUsQ0FBQyxRQUFRLEVBQUUsQ0FBQztJQUN4QyxJQUFJLFdBQVcsR0FBRyxJQUFJLE1BQU0sQ0FBQyxLQUFLLENBQUMsS0FBSyxDQUFDLEtBQUssQ0FBQyxJQUFJLEVBQUUsS0FBSyxDQUFDLEtBQUssQ0FBQyxTQUFTLEVBQUUsS0FBSyxDQUFDLEdBQUcsQ0FBQyxJQUFJLEVBQUUsS0FBSyxDQUFDLEdBQUcsQ0FBQyxTQUFTLENBQUMsQ0FBQztJQUNqSCxPQUFPLElBQUksTUFBTSxDQUFDLFFBQVEsQ0FBQyxNQUFNLENBQUMsR0FBRyxDQUFDLElBQUksQ0FBQyxNQUFNLENBQUMsTUFBTSxDQUFDLEVBQUUsV0FBVyxDQUFDLENBQUM7QUFDekUsQ0FBQyJ9