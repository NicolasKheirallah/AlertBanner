"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const NO_CSP = '';
const DEFAULT_CSP = "default-src 'none'; connect-src vscode-resource: https:; img-src vscode-resource: https:; script-src vscode-resource: 'unsafe-inline';style-src vscode-resource: 'unsafe-inline' http: https: data:;";
function getCSP(securityConfiguration) {
    const { securityPolicy } = securityConfiguration;
    if (securityPolicy === "strict" /* Strict */) {
        return DEFAULT_CSP;
    }
    if (securityPolicy === "disabled" /* Disabled */) {
        return NO_CSP;
    }
    return DEFAULT_CSP;
}
exports.getCSP = getCSP;
//# sourceMappingURL=CSP.js.map