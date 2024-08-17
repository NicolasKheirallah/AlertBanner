"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const vscode = require("vscode");
const path = require("path");
const fs = require("fs");
const request = require("request-promise-native");
class Utils {
    static warnIfNotSppkg(fileUri) {
        if (path.extname(fileUri.fsPath) !== '.sppkg') {
            vscode.window.showErrorMessage(`File '${path.basename(fileUri.path)}' is not a SharePoint Framework solution package`);
            return false;
        }
        return true;
    }
    static getRequestDigestForSite(siteUrl, auth) {
        return new Promise((resolve, reject) => {
            auth
                .getAccessToken()
                .then((accessToken) => {
                const requestOptions = {
                    url: `${siteUrl}/_api/contextinfo`,
                    headers: {
                        authorization: `Bearer ${accessToken}`,
                        accept: 'application/json;odata=nometadata',
                    },
                    json: true
                };
                console.log(requestOptions);
                return request.post(requestOptions);
            })
                .then((res) => {
                console.log(res);
                resolve(res.FormDigestValue);
            }, (error) => {
                reject(error);
            });
        });
    }
    static getTenantAppCatalogUrl(auth) {
        return new Promise((resolve, reject) => {
            auth
                .getAccessToken()
                .then((accessToken) => {
                const requestOptions = {
                    url: `${auth.sharePointUrl}/_api/SP_TenantSettings_Current`,
                    headers: {
                        authorization: `Bearer ${accessToken}`,
                        accept: 'application/json;odata=nometadata',
                    },
                    json: true
                };
                console.log(requestOptions);
                return request.get(requestOptions);
            })
                .then((res) => {
                console.log(res);
                if (res.CorporateCatalogUrl) {
                    resolve(res.CorporateCatalogUrl);
                }
                else {
                    reject(`Couldn't locate tenant app catalog`);
                }
            }, (error) => {
                reject(error);
            });
        });
    }
    static addSolutionToCatalog(fileUri, appCatalogUrl, tenantAppCatalog, auth, output) {
        return new Promise((resolve, reject) => {
            let accessToken = '';
            auth
                .getAccessToken()
                .then((at) => {
                output.write(`- Retrieving request digest for ${appCatalogUrl}...`);
                accessToken = at;
                return Utils.getRequestDigestForSite(appCatalogUrl, auth);
            })
                .then((requestDigest) => {
                const solutionFileName = path.basename(fileUri.fsPath).toLowerCase();
                output.write(`- Adding solution ${solutionFileName} to the app catalog ${appCatalogUrl}...`);
                const requestOptions = {
                    url: `${appCatalogUrl}/_api/web/${(tenantAppCatalog ? 'tenantappcatalog' : 'sitecollectionappcatalog')}/Add(overwrite=true, url='${solutionFileName}')`,
                    headers: {
                        authorization: `Bearer ${accessToken}`,
                        accept: 'application/json;odata=nometadata',
                        'X-RequestDigest': requestDigest,
                        binaryStringRequestBody: 'true'
                    },
                    body: fs.readFileSync(fileUri.fsPath)
                };
                console.log(requestOptions);
                return request.post(requestOptions);
            })
                .then((res) => {
                console.log(res);
                const solution = JSON.parse(res);
                resolve(solution.UniqueId);
            }, (error) => {
                reject(error);
            });
        });
    }
    static deploySolution(fileUri, tenantAppCatalog, skipFeatureDeployment, auth, output) {
        output.show();
        output.write('Deploying solution package...');
        let accessToken = '';
        let appCatalogUrl = '';
        let solutionId = '';
        auth
            .getAccessToken()
            .then((at) => {
            accessToken = at;
            if (tenantAppCatalog) {
                return Utils.getTenantAppCatalogUrl(auth);
            }
            else {
                return vscode.window.showInputBox({
                    ignoreFocusOut: true,
                    prompt: 'URL of your SharePoint site collection app catalog',
                    placeHolder: 'https://contoso.sharepoint.com/site/marketing'
                });
            }
        })
            .then((catalogUrl) => {
            console.log(catalogUrl);
            if (!catalogUrl) {
                if (tenantAppCatalog) {
                    throw 'Unable to determine tenant app catalog URL';
                }
                else {
                    throw 'Please specify the URL of the site collection app catalog';
                }
            }
            appCatalogUrl = catalogUrl;
            return Utils.addSolutionToCatalog(fileUri, appCatalogUrl, tenantAppCatalog, auth, output);
        })
            .then((sid) => {
            solutionId = sid;
            output.write(`- Retrieving request digest for ${appCatalogUrl}...`);
            return Utils.getRequestDigestForSite(appCatalogUrl, auth);
        })
            .then((requestDigest) => {
            output.write(`- Deploying solution to ${skipFeatureDeployment ? 'all sites' : 'the app catalog'}...`);
            const requestOptions = {
                url: `${appCatalogUrl}/_api/web/${(tenantAppCatalog ? 'tenantappcatalog' : 'sitecollectionappcatalog')}/AvailableApps/GetById('${solutionId}')/deploy`,
                headers: {
                    authorization: `Bearer ${accessToken}`,
                    accept: 'application/json;odata=nometadata',
                    'content-type': 'application/json;odata=nometadata;charset=utf-8',
                    'X-RequestDigest': requestDigest
                },
                body: { 'skipFeatureDeployment': skipFeatureDeployment },
                json: true
            };
            console.log(requestOptions);
            return request.post(requestOptions);
        })
            .then((res) => {
            console.log(res);
            output.write('DONE');
            output.write('');
        }, (error) => {
            console.log(error);
            let message = error;
            if (typeof error === 'string') {
                error = JSON.parse(error);
            }
            if (typeof error.error === 'string') {
                error = JSON.parse(error.error);
                if (error['odata.error']) {
                    message = error['odata.error'].message.value;
                }
            }
            else {
                if (error.error &&
                    error.error['odata.error']) {
                    message = error.error['odata.error'].message.value;
                }
            }
            output.write(`Error: ${message}`);
            output.write('');
            vscode.window.showErrorMessage(`The following error has occurred while deploying the solution package to the app catalog: ${message}`);
        });
    }
    static getUserNameFromAccessToken(accessToken) {
        let userName = '';
        if (!accessToken || accessToken.length === 0) {
            return userName;
        }
        const chunks = accessToken.split('.');
        if (chunks.length !== 3) {
            return userName;
        }
        const tokenString = Buffer.from(chunks[1], 'base64').toString();
        try {
            const token = JSON.parse(tokenString);
            userName = token.upn;
        }
        catch (_a) {
        }
        return userName;
    }
}
exports.Utils = Utils;
//# sourceMappingURL=Utils.js.map