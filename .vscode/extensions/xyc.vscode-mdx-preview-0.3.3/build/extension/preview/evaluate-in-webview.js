"use strict";
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (Object.hasOwnProperty.call(mod, k)) result[k] = mod[k];
    result["default"] = mod;
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
const fs = __importStar(require("fs"));
const transform_1 = require("../module-fetcher/transform");
const precinct = require("precinct");
const { performance } = require('perf_hooks');
/**
 * @param text text to preview
 * @param fsPath fsPath of current document
 */
async function evaluateInWebview(preview, text, fsPath) {
    const { webviewHandle } = preview;
    try {
        performance.mark('preview/start');
        const code = await transform_1.transformEntry(text, fsPath, preview);
        const entryFilePath = fs.realpathSync(fsPath);
        const entryFileDependencies = precinct(code);
        console.log(code);
        console.log(entryFilePath);
        console.log(entryFileDependencies);
        await preview.webviewHandshakePromise;
        if (webviewHandle && webviewHandle.updatePreview) {
            webviewHandle.updatePreview(code, entryFilePath, entryFileDependencies);
        }
    }
    catch (error) {
        console.error(error);
        if (webviewHandle) {
            webviewHandle.showPreviewError({
                message: error.message
            });
        }
    }
}
exports.default = evaluateInWebview;
//# sourceMappingURL=evaluate-in-webview.js.map