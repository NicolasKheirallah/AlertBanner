'use strict';
Object.defineProperty(exports, "__esModule", { value: true });
const vscode = require("vscode");
const command_1 = require("./command");
const _1 = require(".");
function activate(context) {
    const statusBarItem = new _1.StatusBarItem();
    const output = new _1.Output();
    const auth = new _1.Auth(output, statusBarItem);
    const deployTenantAppCatalogCommand = vscode.commands.registerCommand('rencoreSpfxDeploy.deployTenantAppCatalog', (fileUri) => {
        command_1.deployTenant(fileUri, auth, output);
    });
    const deployTenantAppCatalogGlobalCommand = vscode.commands.registerCommand('rencoreSpfxDeploy.deployTenantAppCatalogGlobal', (fileUri) => {
        command_1.deployTenantGlobal(fileUri, auth, output);
    });
    const deploySiteCollectionAppCatalogCommand = vscode.commands.registerCommand('rencoreSpfxDeploy.deploySiteCollectionAppCatalog', (fileUri) => {
        command_1.deploySiteCollection(fileUri, auth, output);
    });
    const deploySiteCollctionAppCatalogGlobalCommand = vscode.commands.registerCommand('rencoreSpfxDeploy.deploySiteCollectionAppCatalogGlobal', (fileUri) => {
        command_1.deploySiteCollectionGlobal(fileUri, auth, output);
    });
    const statusCommand = vscode.commands.registerCommand('rencoreSpfxDeploy.status', (fileUri) => {
        command_1.status(auth, statusBarItem);
    });
    vscode.commands.executeCommand('setContext', 'hasSppkg', true);
    context.subscriptions.push(deployTenantAppCatalogCommand, deployTenantAppCatalogGlobalCommand, deploySiteCollectionAppCatalogCommand, deploySiteCollctionAppCatalogGlobalCommand, statusCommand, auth, statusBarItem, output);
}
exports.activate = activate;
// this method is called when your extension is deactivated
function deactivate() {
}
exports.deactivate = deactivate;
//# sourceMappingURL=extension.js.map