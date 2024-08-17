"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const vscode = require("vscode");
const stripJsonComments = require("strip-json-comments");
class FileHelper {
    /**
     * Retrieve the JSON file contents
     *
     * @param fileUri
     */
    static getJsonContents(fileUri) {
        return __awaiter(this, void 0, void 0, function* () {
            const file = yield vscode.workspace.openTextDocument(fileUri.path);
            if (file) {
                const contents = file.getText();
                if (contents) {
                    const configJson = JSON.parse(stripJsonComments(contents));
                    return configJson;
                }
            }
            return null;
        });
    }
}
exports.FileHelper = FileHelper;
//# sourceMappingURL=FileHelper.js.map