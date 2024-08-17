"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const vscode = require("vscode");
const adal_node_1 = require("adal-node");
class Auth extends vscode.Disposable {
    constructor(output, statusBarItem) {
        super(() => {
            this.ctx = undefined;
        });
        this.output = output;
        this.statusBarItem = statusBarItem;
        this.appId = 'be278e09-27a8-47ac-91fe-27f05350b7d8';
        this.ctx = new adal_node_1.AuthenticationContext('https://login.microsoftonline.com/common');
    }
    get isConnected() {
        return !!this.accessToken;
    }
    getAccessToken() {
        return new Promise((resolve, reject) => {
            if (this.accessToken && this.accessTokenExpiresOn) {
                const expiresOn = new Date(this.accessTokenExpiresOn);
                if (expiresOn > new Date()) {
                    resolve(this.accessToken);
                    return;
                }
                else {
                    this.ctx.acquireTokenWithRefreshToken(this.refreshToken, this.appId, this.sharePointUrl, (error, response) => {
                        if (error) {
                            reject((response && response.error_description) || error.message);
                            return;
                        }
                        const token = response;
                        this.accessToken = token.accessToken;
                        this.accessTokenExpiresOn = token.expiresOn;
                        this.refreshToken = token.refreshToken;
                        resolve(this.accessToken);
                    });
                }
            }
            (() => {
                let sharePointUrl = vscode.workspace.getConfiguration('rencoreSpfxDeploy').get('sharePointUrl');
                if (sharePointUrl) {
                    return Promise.resolve(sharePointUrl);
                }
                return vscode.window.showInputBox({
                    ignoreFocusOut: true,
                    prompt: 'URL of your SharePoint tenant',
                    placeHolder: 'https://contoso.sharepoint.com'
                });
            })()
                .then((sharePointUrl) => {
                if (!sharePointUrl) {
                    reject('Please specify SharePoint URL');
                }
                this.sharePointUrl = sharePointUrl;
                this.ctx.acquireUserCode(this.sharePointUrl, this.appId, 'en-us', (error, response) => {
                    if (error) {
                        reject((response && response.error_description) || error.message);
                        return;
                    }
                    this.output.write(`- ${response.message}`);
                    vscode.window.showInformationMessage(response.message);
                    this.ctx.acquireTokenWithDeviceCode(this.sharePointUrl, this.appId, response, (error, response) => {
                        if (error) {
                            reject((response && response.error_description) || error.message);
                            return;
                        }
                        const token = response;
                        this.accessToken = token.accessToken;
                        this.accessTokenExpiresOn = token.expiresOn;
                        this.refreshToken = token.refreshToken;
                        this.statusBarItem.connected = true;
                        resolve(this.accessToken);
                    });
                });
            }, (error) => {
                this.statusBarItem.connected = false;
                reject(error);
            });
        });
    }
    disconnect() {
        this.sharePointUrl = undefined;
        this.accessToken = undefined;
        this.accessTokenExpiresOn = undefined;
        this.refreshToken = undefined;
        this.statusBarItem.connected = false;
    }
}
exports.Auth = Auth;
//# sourceMappingURL=Auth.js.map