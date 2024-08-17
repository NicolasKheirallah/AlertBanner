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
const fs = require("fs");
const cdnQuality_1 = require("./cdnQuality");
const scriptChecker_1 = require("./scriptChecker");
const externalLibrary_1 = require("./externalLibrary");
/**
 * Questions
 */
const scriptUrlOption = {
    ignoreFocusOut: true,
    placeHolder: "Enter the URL of the external library you want to check and add to your config.",
    prompt: "Example: https://code.jquery.com/jquery-2.2.4.min.js"
};
const moduleNameOption = {
    ignoreFocusOut: true,
    placeHolder: "Enter the name of your module.",
    prompt: "Enter the name of your module. Example: jquery, angular, ..."
};
const scriptPluginOption = {
    ignoreFocusOut: true,
    placeHolder: "Is this module a plugin?",
    prompt: `Is this module a plugin? Enter: "true" OR "false"`,
    validateInput: (val) => (val === 'true' || val === 'false' ? "" : `Please enter: "true" OR "false".`)
};
const globalNameOption = {
    ignoreFocusOut: true,
    placeHolder: "Enter the global module name.",
    prompt: "Enter the global module name. Example: jquery, angular, ...",
};
const dependencyOptions = {
    ignoreFocusOut: true,
    placeHolder: "Enter the module dependencies (comma-separated if multiple).",
    prompt: "Enter the module dependencies (comma-separated if multiple). Example: jquery, angular, ...",
};
/**
 * Visual Studio Code Activate Extension
 * @param context
 */
function activate(context) {
    // register VSCode command "SPFx Script Check"
    const disposable = vscode.commands.registerCommand('spfx.scriptcheck', () => {
        // Show the script URL option
        const url = vscode.window.showInputBox(scriptUrlOption).then((url) => __awaiter(this, void 0, void 0, function* () {
            if (!url) {
                return;
            }
            const isSharePointUrl = url.indexOf('.sharepoint.com') > -1 && url.indexOf('publicdn') < 0;
            // Create script data for calling the API
            const scriptData = { url };
            // Check the CDN quality
            vscode.window.withProgress({
                location: vscode.ProgressLocation.Window,
                title: 'Detecting CDN quality...'
            }, () => {
                return cdnQuality_1.default.test(scriptData, isSharePointUrl);
            }).then((quality) => {
                vscode.window.showInformationMessage(quality);
            });
            // Check the script type
            if (isSharePointUrl) {
                vscode.window.showWarningMessage(`Can't analyze script ${url} from SharePoint because it's not available to anonymous users`);
                return;
            }
            vscode.window.withProgress({
                location: vscode.ProgressLocation.Window,
                title: 'Detecting script type...'
            }, () => {
                return scriptChecker_1.default.check(scriptData);
            }).then((scriptType) => __awaiter(this, void 0, void 0, function* () {
                if (typeof scriptType === "string") {
                    vscode.window.showErrorMessage(`${scriptType}`);
                    return;
                }
                else if (scriptType === null) {
                    vscode.window.showErrorMessage('Unable to detect script type');
                    return;
                }
                else {
                    // Ask the module name
                    const moduleName = yield vscode.window.showInputBox(moduleNameOption);
                    if (moduleName) {
                        // Check to see if other questions need to be asked when script is not a module
                        let scriptPlugin = null;
                        let globalName = null;
                        let scriptDependencies = null;
                        // Check if it was a non-module
                        if (scriptType === scriptChecker_1.ScriptType.nonModule) {
                            // Set the default value for the plugin script to "false"
                            scriptPluginOption.value = "false";
                            const scriptPluginTxt = yield vscode.window.showInputBox(scriptPluginOption);
                            if (!scriptPluginTxt) {
                                vscode.window.showErrorMessage(`You entered an incorrect value.`);
                                return;
                            }
                            scriptPlugin = scriptPluginTxt === "true";
                        }
                        // Check if file is a plugin
                        if (scriptPlugin) {
                            globalName = yield vscode.window.showInputBox(globalNameOption);
                            scriptDependencies = yield vscode.window.showInputBox(dependencyOptions);
                        }
                        // Retrieve the SharePoint Framework config file
                        const filesUri = yield vscode.workspace.findFiles('**/config/config.json');
                        if (!filesUri || filesUri.length === 0) {
                            return;
                        }
                        // Get the first config file
                        const configFileUri = filesUri[0];
                        // Open and show the config file
                        vscode.window.showTextDocument(configFileUri);
                        const configFile = yield vscode.workspace.openTextDocument(configFileUri);
                        const configContent = configFile.getText();
                        // Check if file content is retrieved
                        if (!configContent) {
                            vscode.window.showErrorMessage(`Failed retrieving the config.json file.`);
                            return;
                        }
                        // Parse the JSON content
                        const configJson = JSON.parse(configContent);
                        if (!configJson.externals) {
                            // If the config file does not contain an external section, we stop,
                            // it might be a wrong file or something is wrong with the file
                            vscode.window.showErrorMessage(`Your config.json file does not have the "externals" section.`);
                            return;
                        }
                        // Update the config file based on the information from the questions
                        const updatedConfig = externalLibrary_1.default.update(configJson, scriptType, moduleName, url, globalName, scriptDependencies);
                        // Write the updated content back to the config file
                        fs.writeFileSync(configFileUri.path, JSON.stringify(updatedConfig, null, 2), 'utf8');
                        vscode.window.showInformationMessage("Script reference successfully added to config.json");
                    }
                }
            }));
        }));
    });
    context.subscriptions.push(disposable);
}
exports.activate = activate;
// this method is called when your extension is deactivated
function deactivate() { }
exports.deactivate = deactivate;
//# sourceMappingURL=extension.js.map