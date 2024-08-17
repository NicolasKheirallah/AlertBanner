"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const vscode = require("vscode");
const Removal_1 = require("./commands/Removal");
exports.EXTENSION_LOG_NAME = "[SPFx remove]";
function activate(context) {
    let disposable = vscode.commands.registerCommand('estruyf.removespfxcomponent', Removal_1.removeComponent);
    context.subscriptions.push(disposable);
    console.log(`${exports.EXTENSION_LOG_NAME}: is now active!`);
}
exports.activate = activate;
function deactivate() { }
exports.deactivate = deactivate;
//# sourceMappingURL=extension.js.map