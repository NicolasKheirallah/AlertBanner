"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const vscode_1 = require("vscode");
const preview_manager_1 = require("./preview/preview-manager");
const checkFsPath_1 = require("./security/checkFsPath");
let disposables = [];
// https://code.visualstudio.com/docs/extensionAPI/vscode-api#_workspace
// https://code.visualstudio.com/docs/extensionAPI/vscode-api#TextDocument
function initWorkspaceHandlers(context) {
    vscode_1.workspace.onDidOpenTextDocument(event => { });
    vscode_1.workspace.onDidSaveTextDocument(event => {
        if (preview_manager_1.currentPreview) {
            preview_manager_1.currentPreview.handleDidSaveTextDocument(event.uri.fsPath);
        }
    }, null, disposables);
    vscode_1.workspace.onDidChangeTextDocument(event => {
        if (preview_manager_1.currentPreview) {
            preview_manager_1.currentPreview.handleDidChangeTextDocument(event.document.uri.fsPath, event.document);
        }
    }, null, disposables);
    vscode_1.workspace.onDidChangeConfiguration(event => {
        if (preview_manager_1.currentPreview) {
            preview_manager_1.currentPreview.updateConfiguration();
        }
    });
    vscode_1.workspace.onDidChangeWorkspaceFolders(() => {
        checkFsPath_1.handleDidChangeWorkspaceFolders();
    });
    // workspace.createFileSystemWatcher
    // workspace.findFiles
}
exports.initWorkspaceHandlers = initWorkspaceHandlers;
//# sourceMappingURL=workspace-manager.js.map