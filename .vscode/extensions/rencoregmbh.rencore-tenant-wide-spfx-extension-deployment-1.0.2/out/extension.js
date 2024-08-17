'use strict';
Object.defineProperty(exports, "__esModule", { value: true });
const vscode = require("vscode");
const addDeploymentInfo_1 = require("./command/addDeploymentInfo");
function activate(context) {
    const disposable = vscode.commands.registerCommand('rencoreSpfxGlobalExtension.addDeploymentInfo', (fileUri) => {
        addDeploymentInfo_1.addDeploymentInfo(fileUri);
    });
    context.subscriptions.push(disposable);
}
exports.activate = activate;
function deactivate() {
}
exports.deactivate = deactivate;
//# sourceMappingURL=extension.js.map