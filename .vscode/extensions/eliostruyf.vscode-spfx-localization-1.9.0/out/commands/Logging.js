"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const vscode = require("vscode");
const EXTENSION_NAME = "SPFx Localization";
class Logging {
    /**
     * Show an information message
     *
     * @param msg
     */
    static info(msg) {
        vscode.window.showInformationMessage(`${EXTENSION_NAME}: ${msg}`);
    }
    /**
     * Show an error message
     *
     * @param msg
     */
    static error(msg) {
        vscode.window.showErrorMessage(`${EXTENSION_NAME}: ${msg}`);
    }
    /**
     * Show an error message
     *
     * @param msg
     */
    static warning(msg) {
        vscode.window.showWarningMessage(`${EXTENSION_NAME}: ${msg}`);
    }
}
exports.default = Logging;
//# sourceMappingURL=Logging.js.map