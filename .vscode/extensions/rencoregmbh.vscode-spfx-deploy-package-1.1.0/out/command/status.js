"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const vscode = require("vscode");
const __1 = require("../");
function status(auth, statusBarItem) {
    if (auth.isConnected) {
        auth
            .getAccessToken()
            .then((accessToken) => {
            const userName = __1.Utils.getUserNameFromAccessToken(accessToken);
            vscode.window
                .showInformationMessage(`Connected to ${auth.sharePointUrl} as ${userName}`, 'OK', 'Disconnect')
                .then((action) => {
                if (action === 'Disconnect') {
                    auth.disconnect();
                }
            });
        });
    }
    else {
        vscode.window.showInformationMessage('Not connected to SharePoint', 'OK');
    }
}
exports.status = status;
//# sourceMappingURL=status.js.map