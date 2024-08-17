'use strict';
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (Object.hasOwnProperty.call(mod, k)) result[k] = mod[k];
    result["default"] = mod;
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
const vscode = __importStar(require("vscode"));
const preview_manager_1 = require("./preview/preview-manager");
const security_1 = require("./security/security");
const webview_manager_1 = require("./preview/webview-manager");
const workspace_manager_1 = require("./workspace-manager");
// this method is called when your extension is activated
// your extension is activated the very first time the command is executed
function activate(context) {
    webview_manager_1.initWebviewAppHTMLResources(context);
    workspace_manager_1.initWorkspaceHandlers(context);
    let openPreviewCommand = vscode.commands.registerCommand('mdx-preview.commands.openPreview', () => {
        preview_manager_1.openPreview();
    });
    let refreshPreviewCommand = vscode.commands.registerCommand('mdx-preview.commands.refreshPreview', () => {
        preview_manager_1.refreshPreview();
    });
    let toggleUseVscodeMarkdownStylesCommand = vscode.commands.registerCommand('mdx-preview.commands.toggleUseVscodeMarkdownStyles', () => {
        const extensionConfig = vscode.workspace.getConfiguration('mdx-preview');
        const useVscodeMarkdownStyles = extensionConfig.get('preview.useVscodeMarkdownStyles', false);
        extensionConfig.update('preview.useVscodeMarkdownStyles', !useVscodeMarkdownStyles);
    });
    let toggleUseWhiteBackgroundCommand = vscode.commands.registerCommand('mdx-preview.commands.toggleUseWhiteBackground', () => {
        const extensionConfig = vscode.workspace.getConfiguration('mdx-preview');
        const useWhiteBackground = extensionConfig.get('preview.useWhiteBackground', false);
        extensionConfig.update('preview.useWhiteBackground', !useWhiteBackground);
    });
    let toggleChangeSecuritySettings = vscode.commands.registerCommand('mdx-preview.commands.changeSecuritySettings', () => {
        security_1.selectSecurityPolicy();
    });
    context.subscriptions.push(openPreviewCommand, refreshPreviewCommand, toggleUseVscodeMarkdownStylesCommand, toggleUseWhiteBackgroundCommand, toggleChangeSecuritySettings);
}
exports.activate = activate;
// this method is called when your extension is deactivated
function deactivate() {
}
exports.deactivate = deactivate;
//# sourceMappingURL=extension.js.map