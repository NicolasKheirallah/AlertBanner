"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.activate = void 0;
// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
const path = require("path");
const os = require("os");
const fs = require("fs");
const vscode_1 = require("vscode");
const json_to_ts_1 = require("json-to-ts");
// this method is called when your extension is activated
// your extension is activated the very first time the command is executed
function activate(context) {
    context.subscriptions.push(vscode_1.commands.registerCommand("jsonToTs.fromClipboard", transformFromClipboard));
    context.subscriptions.push(vscode_1.commands.registerCommand("jsonToTs.fromSelection", transformFromSelection));
}
exports.activate = activate;
function transformFromSelection() {
    const tmpFilePath = path.join(os.tmpdir(), "json-to-ts.ts");
    const tmpFileUri = vscode_1.Uri.file(tmpFilePath);
    getSelectedText()
        .then(validateLength)
        .then(parseJson)
        .then((json) => {
        return (0, json_to_ts_1.default)(json).reduce((a, b) => `${a}\n\n${b}`);
    })
        .then((interfaces) => {
        fs.writeFileSync(tmpFilePath, interfaces);
    })
        .then(() => {
        vscode_1.commands.executeCommand("vscode.open", tmpFileUri, getViewColumn());
    })
        .catch(handleError);
}
async function transformFromClipboard() {
    const text = await vscode_1.env.clipboard.readText();
    Promise.resolve(text)
        .then(validateLength)
        .then(parseJson)
        .then((json) => (0, json_to_ts_1.default)(json).reduce((a, b) => `${a}\n\n${b}`))
        .then((interfaces) => {
        pasteToMarker(interfaces);
    })
        .catch(handleError);
}
function handleError(error) {
    vscode_1.window.showErrorMessage(error.message);
}
function parseJson(json) {
    const tryEval = (str) => eval(`const a = ${str}; a`);
    try {
        return Promise.resolve(JSON.parse(json));
        // eslint-disable-next-line no-empty
    }
    catch (ignored) { }
    try {
        return Promise.resolve(tryEval(json));
    }
    catch (error) {
        return Promise.reject(new Error("Selected string is not a valid JSON"));
    }
}
function getViewColumn() {
    const activeEditor = vscode_1.window.activeTextEditor;
    if (!activeEditor) {
        return vscode_1.ViewColumn.One;
    }
    switch (activeEditor.viewColumn) {
        case vscode_1.ViewColumn.One:
            return vscode_1.ViewColumn.Two;
        case vscode_1.ViewColumn.Two:
            return vscode_1.ViewColumn.Three;
    }
    return activeEditor.viewColumn;
}
function pasteToMarker(content) {
    const { activeTextEditor } = vscode_1.window;
    return activeTextEditor?.edit((editBuilder) => {
        editBuilder.replace(activeTextEditor.selection, content);
    });
}
function getSelectedText() {
    const { selection, document } = vscode_1.window.activeTextEditor;
    return Promise.resolve(document.getText(selection).trim());
}
function validateLength(text) {
    if (text.length === 0) {
        return Promise.reject(new Error("Nothing selected"));
    }
    else {
        return Promise.resolve(text);
    }
}
//# sourceMappingURL=extension.js.map