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
exports.activate = void 0;
const path = require("path");
const request_light_1 = require("request-light");
const vscode = require("vscode");
function activate(context) {
    context.subscriptions.push(vscode.commands.registerTextEditorCommand('psl.previewDocumentation', preparePreview));
    checkForDocumentationServer();
    vscode.workspace.onDidChangeConfiguration(event => {
        if (!event.affectsConfiguration('psl'))
            return;
        checkForDocumentationServer();
    });
}
exports.activate = activate;
function checkForDocumentationServer() {
    const documentationServer = vscode.workspace.getConfiguration('psl', null).get('documentationServer');
    if (documentationServer) {
        vscode.commands.executeCommand('setContext', 'psl.hasDocumentationServer', true);
        return documentationServer;
    }
    else {
        vscode.commands.executeCommand('setContext', 'psl.hasDocumentationServer', false);
        return '';
    }
}
function preparePreview(textEditor) {
    return __awaiter(this, void 0, void 0, function* () {
        const documentationServer = checkForDocumentationServer();
        if (!documentationServer)
            return;
        const markdown = yield getMarkdownFromApi(textEditor.document.getText(), path.basename(textEditor.document.fileName), documentationServer);
        if (!markdown)
            return;
        showPreview(markdown);
    });
}
function showPreview(markdown) {
    return __awaiter(this, void 0, void 0, function* () {
        const untitledDoc = yield vscode.workspace.openTextDocument({ language: 'markdown', content: markdown });
        vscode.commands.executeCommand('markdown.showPreview', untitledDoc.uri);
    });
}
function getMarkdownFromApi(pslText, fileName, documentationServer) {
    return __awaiter(this, void 0, void 0, function* () {
        try {
            const data = JSON.stringify({
                sourceText: pslText,
            });
            const response = yield (0, request_light_1.xhr)({
                data,
                headers: {
                    'Content-Length': `${Buffer.byteLength(data)}`,
                    'Content-Type': 'application/json',
                },
                type: 'POST',
                url: documentationServer + fileName,
            });
            return response.responseText;
        }
        catch (e) {
            vscode.window.showErrorMessage(e.responseText);
            return '';
        }
    });
}
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoicHJldmlld0RvY3VtZW50YXRpb24uanMiLCJzb3VyY2VSb290IjoiIiwic291cmNlcyI6WyIuLi8uLi8uLi9zcmMvbGFuZ3VhZ2UvcHJldmlld0RvY3VtZW50YXRpb24udHMiXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6Ijs7Ozs7Ozs7Ozs7O0FBQUEsNkJBQTZCO0FBQzdCLGlEQUFtQztBQUNuQyxpQ0FBaUM7QUFFakMsU0FBZ0IsUUFBUSxDQUFDLE9BQWdDO0lBQ3hELE9BQU8sQ0FBQyxhQUFhLENBQUMsSUFBSSxDQUN6QixNQUFNLENBQUMsUUFBUSxDQUFDLHlCQUF5QixDQUN4QywwQkFBMEIsRUFDMUIsY0FBYyxDQUNkLENBQ0QsQ0FBQztJQUVGLDJCQUEyQixFQUFFLENBQUM7SUFFOUIsTUFBTSxDQUFDLFNBQVMsQ0FBQyx3QkFBd0IsQ0FBQyxLQUFLLENBQUMsRUFBRTtRQUNqRCxJQUFJLENBQUMsS0FBSyxDQUFDLG9CQUFvQixDQUFDLEtBQUssQ0FBQztZQUFFLE9BQU87UUFDL0MsMkJBQTJCLEVBQUUsQ0FBQztJQUMvQixDQUFDLENBQUMsQ0FBQztBQUNKLENBQUM7QUFkRCw0QkFjQztBQUVELFNBQVMsMkJBQTJCO0lBQ25DLE1BQU0sbUJBQW1CLEdBQVcsTUFBTSxDQUFDLFNBQVMsQ0FBQyxnQkFBZ0IsQ0FBQyxLQUFLLEVBQUUsSUFBSSxDQUFDLENBQUMsR0FBRyxDQUFDLHFCQUFxQixDQUFDLENBQUM7SUFDOUcsSUFBSSxtQkFBbUIsRUFBRTtRQUN4QixNQUFNLENBQUMsUUFBUSxDQUFDLGNBQWMsQ0FBQyxZQUFZLEVBQUUsNEJBQTRCLEVBQUUsSUFBSSxDQUFDLENBQUM7UUFDakYsT0FBTyxtQkFBbUIsQ0FBQztLQUMzQjtTQUNJO1FBQ0osTUFBTSxDQUFDLFFBQVEsQ0FBQyxjQUFjLENBQUMsWUFBWSxFQUFFLDRCQUE0QixFQUFFLEtBQUssQ0FBQyxDQUFDO1FBQ2xGLE9BQU8sRUFBRSxDQUFDO0tBQ1Y7QUFDRixDQUFDO0FBRUQsU0FBZSxjQUFjLENBQUMsVUFBNkI7O1FBQzFELE1BQU0sbUJBQW1CLEdBQVcsMkJBQTJCLEVBQUUsQ0FBQztRQUNsRSxJQUFJLENBQUMsbUJBQW1CO1lBQUUsT0FBTztRQUVqQyxNQUFNLFFBQVEsR0FBRyxNQUFNLGtCQUFrQixDQUN4QyxVQUFVLENBQUMsUUFBUSxDQUFDLE9BQU8sRUFBRSxFQUM3QixJQUFJLENBQUMsUUFBUSxDQUFDLFVBQVUsQ0FBQyxRQUFRLENBQUMsUUFBUSxDQUFDLEVBQzNDLG1CQUFtQixDQUNuQixDQUFDO1FBQ0YsSUFBSSxDQUFDLFFBQVE7WUFBRSxPQUFPO1FBQ3RCLFdBQVcsQ0FBQyxRQUFRLENBQUMsQ0FBQztJQUN2QixDQUFDO0NBQUE7QUFFRCxTQUFlLFdBQVcsQ0FBQyxRQUFnQjs7UUFDMUMsTUFBTSxXQUFXLEdBQUcsTUFBTSxNQUFNLENBQUMsU0FBUyxDQUFDLGdCQUFnQixDQUFDLEVBQUUsUUFBUSxFQUFFLFVBQVUsRUFBRSxPQUFPLEVBQUUsUUFBUSxFQUFFLENBQUMsQ0FBQztRQUN6RyxNQUFNLENBQUMsUUFBUSxDQUFDLGNBQWMsQ0FBQyxzQkFBc0IsRUFBRSxXQUFXLENBQUMsR0FBRyxDQUFDLENBQUM7SUFDekUsQ0FBQztDQUFBO0FBRUQsU0FBZSxrQkFBa0IsQ0FBQyxPQUFlLEVBQUUsUUFBZ0IsRUFBRSxtQkFBMkI7O1FBQy9GLElBQUk7WUFDSCxNQUFNLElBQUksR0FBVyxJQUFJLENBQUMsU0FBUyxDQUFDO2dCQUNuQyxVQUFVLEVBQUUsT0FBTzthQUNuQixDQUFDLENBQUM7WUFDSCxNQUFNLFFBQVEsR0FBRyxNQUFNLElBQUEsbUJBQUcsRUFBQztnQkFDMUIsSUFBSTtnQkFDSixPQUFPLEVBQUU7b0JBQ1IsZ0JBQWdCLEVBQUUsR0FBRyxNQUFNLENBQUMsVUFBVSxDQUFDLElBQUksQ0FBQyxFQUFFO29CQUM5QyxjQUFjLEVBQUUsa0JBQWtCO2lCQUNsQztnQkFDRCxJQUFJLEVBQUUsTUFBTTtnQkFDWixHQUFHLEVBQUUsbUJBQW1CLEdBQUcsUUFBUTthQUNuQyxDQUFDLENBQUM7WUFDSCxPQUFPLFFBQVEsQ0FBQyxZQUFZLENBQUM7U0FDN0I7UUFDRCxPQUFPLENBQUMsRUFBRTtZQUNULE1BQU0sQ0FBQyxNQUFNLENBQUMsZ0JBQWdCLENBQUMsQ0FBQyxDQUFDLFlBQVksQ0FBQyxDQUFDO1lBQy9DLE9BQU8sRUFBRSxDQUFDO1NBQ1Y7SUFDRixDQUFDO0NBQUEifQ==