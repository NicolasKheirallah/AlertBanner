"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const vscode = require("vscode");
class StatusBarItem extends vscode.Disposable {
    set connected(c) {
        this.statusBarItem.text = c ? '$(globe) Connected' : '$(circle-slash) Disconnected';
        this.statusBarItem.tooltip = `Rencore Deploy SPFx Package extension ${(c ? '' : 'not ')}connected to SharePoint`;
    }
    constructor() {
        super(() => {
            this.statusBarItem.dispose();
        });
        this.statusBarItem = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Left, 10);
        this.statusBarItem.command = 'rencoreSpfxDeploy.status';
        this.connected = false;
        this.statusBarItem.show();
    }
}
exports.StatusBarItem = StatusBarItem;
//# sourceMappingURL=StatusBarItem.js.map