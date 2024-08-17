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
exports.previewEnabled = exports.activate = void 0;
const vscode = require("vscode");
const extension_1 = require("../extension");
const codeQuality = require("./codeQuality");
const dataItem_1 = require("./dataItem");
const mumps_1 = require("./mumps");
const previewDocumentation = require("./previewDocumentation");
const pslDefinitionProvider_1 = require("./pslDefinitionProvider");
const pslDocument_1 = require("./pslDocument");
const pslHoverProvider_1 = require("./pslHoverProvider");
const pslSignature_1 = require("./pslSignature");
const pslSuggest_1 = require("./pslSuggest");
const config_1 = require("../parser/config");
function activate(context) {
    return __awaiter(this, void 0, void 0, function* () {
        const PSL_MODES = [extension_1.PSL_MODE, extension_1.BATCH_MODE, extension_1.TRIG_MODE];
        const MUMPS_MODES = Object.values(mumps_1.MumpsVirtualDocument.schemes).map(scheme => ({ scheme }));
        context.subscriptions.push(
        // Data Hovers
        vscode.languages.registerHoverProvider(extension_1.DATA_MODE, new dataItem_1.DataHoverProvider()), 
        // Data Document Highlights
        vscode.languages.registerDocumentHighlightProvider(extension_1.DATA_MODE, new dataItem_1.DataDocumentHighlightProvider()));
        PSL_MODES.forEach(pslMode => {
            context.subscriptions.push(
            // Document Symbol Outline
            vscode.languages.registerDocumentSymbolProvider(pslMode, new pslDocument_1.PSLDocumentSymbolProvider()), 
            // Completion Items
            vscode.languages.registerCompletionItemProvider(pslMode, new pslSuggest_1.PSLCompletionItemProvider(), '.'), 
            // Signature Help
            vscode.languages.registerSignatureHelpProvider(pslMode, new pslSignature_1.PSLSignatureHelpProvider(), '(', ','), 
            // Go-to Definitions
            vscode.languages.registerDefinitionProvider(pslMode, new pslDefinitionProvider_1.PSLDefinitionProvider()), 
            // Hovers
            vscode.languages.registerHoverProvider(pslMode, new pslHoverProvider_1.PSLHoverProvider()));
        });
        MUMPS_MODES.forEach(mumpsMode => {
            context.subscriptions.push(
            // Content provider for virtual documents
            vscode.workspace.registerTextDocumentContentProvider(mumpsMode.scheme, new mumps_1.MumpsDocumentProvider()), 
            // Document Symbol Outline
            vscode.languages.registerDocumentSymbolProvider(mumpsMode, new pslDocument_1.MumpsDocumentSymbolProvider()));
        });
        projectActivate(context);
        codeQuality.activate(context);
        previewDocumentation.activate(context);
        // Language Configuration
        const wordPattern = /(-?\d*\.\d[a-zA-Z0-9\%\#]*)|([^\`\~\!\@\^\&\*\(\)\-\=\+\[\{\]\}\\\|\"\;\:\'\'\,\.\<\>\/\?\s_]+)/g;
        vscode.languages.setLanguageConfiguration('psl', { wordPattern });
        vscode.languages.setLanguageConfiguration('profileBatch', { wordPattern });
        vscode.languages.setLanguageConfiguration('profileTrigger', { wordPattern });
    });
}
exports.activate = activate;
function previewEnabled(uri) {
    return vscode.workspace.getConfiguration('psl', uri).get('previewFeatures');
}
exports.previewEnabled = previewEnabled;
function projectActivate(context) {
    return __awaiter(this, void 0, void 0, function* () {
        const workspaces = new Map();
        if (vscode.workspace.workspaceFolders) {
            vscode.workspace.workspaceFolders.forEach(workspace => {
                workspaces.set(workspace.name, workspace.uri.fsPath);
            });
        }
        return Promise.all(vscode.workspace.workspaceFolders
            .map(workspace => new vscode.RelativePattern(workspace, 'profile-project.json'))
            .map((pattern) => __awaiter(this, void 0, void 0, function* () {
            const watcher = vscode.workspace.createFileSystemWatcher(pattern);
            context.subscriptions.push(watcher.onDidChange(uri => {
                (0, config_1.setConfig)(uri.fsPath, workspaces);
            }), watcher.onDidCreate(uri => {
                (0, config_1.setConfig)(uri.fsPath, workspaces);
            }));
            watcher.onDidDelete(uri => {
                (0, config_1.removeConfig)(uri.fsPath);
            });
            const uris = yield vscode.workspace.findFiles(pattern);
            if (!uris.length)
                return;
            yield (0, config_1.setConfig)(uris[0].fsPath, workspaces);
        })));
    });
}
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoiYWN0aXZhdGUuanMiLCJzb3VyY2VSb290IjoiIiwic291cmNlcyI6WyIuLi8uLi8uLi9zcmMvbGFuZ3VhZ2UvYWN0aXZhdGUudHMiXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6Ijs7Ozs7Ozs7Ozs7O0FBQUEsaUNBQWlDO0FBRWpDLDRDQUEwRTtBQUUxRSw2Q0FBNkM7QUFDN0MseUNBQThFO0FBQzlFLG1DQUFzRTtBQUN0RSwrREFBK0Q7QUFDL0QsbUVBQWdFO0FBQ2hFLCtDQUF1RjtBQUN2Rix5REFBc0Q7QUFDdEQsaURBQTBEO0FBQzFELDZDQUF5RDtBQUN6RCw2Q0FBMkQ7QUFFM0QsU0FBc0IsUUFBUSxDQUFDLE9BQWdDOztRQUU5RCxNQUFNLFNBQVMsR0FBRyxDQUFDLG9CQUFRLEVBQUUsc0JBQVUsRUFBRSxxQkFBUyxDQUFDLENBQUM7UUFDcEQsTUFBTSxXQUFXLEdBQTRCLE1BQU0sQ0FBQyxNQUFNLENBQUMsNEJBQW9CLENBQUMsT0FBTyxDQUFDLENBQUMsR0FBRyxDQUFDLE1BQU0sQ0FBQyxFQUFFLENBQUMsQ0FBQyxFQUFFLE1BQU0sRUFBRSxDQUFDLENBQUMsQ0FBQztRQUVySCxPQUFPLENBQUMsYUFBYSxDQUFDLElBQUk7UUFDekIsY0FBYztRQUNkLE1BQU0sQ0FBQyxTQUFTLENBQUMscUJBQXFCLENBQ3JDLHFCQUFTLEVBQUUsSUFBSSw0QkFBaUIsRUFBRSxDQUNsQztRQUVELDJCQUEyQjtRQUMzQixNQUFNLENBQUMsU0FBUyxDQUFDLGlDQUFpQyxDQUNqRCxxQkFBUyxFQUFFLElBQUksd0NBQTZCLEVBQUUsQ0FDOUMsQ0FDRCxDQUFDO1FBRUYsU0FBUyxDQUFDLE9BQU8sQ0FBQyxPQUFPLENBQUMsRUFBRTtZQUMzQixPQUFPLENBQUMsYUFBYSxDQUFDLElBQUk7WUFDekIsMEJBQTBCO1lBQzFCLE1BQU0sQ0FBQyxTQUFTLENBQUMsOEJBQThCLENBQzlDLE9BQU8sRUFBRSxJQUFJLHVDQUF5QixFQUFFLENBQ3hDO1lBRUQsbUJBQW1CO1lBQ25CLE1B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