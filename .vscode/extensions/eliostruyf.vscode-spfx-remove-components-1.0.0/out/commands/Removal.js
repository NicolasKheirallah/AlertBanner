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
const path = require("path");
const rimraf = require("rimraf");
const FileHelper_1 = require("../helpers/FileHelper");
const extension_1 = require("../extension");
/**
 * TODO: Test folder path on windows
 */
const MANIFESTS_PATH = "**/src/**/*.manifest.json";
const CONFIG_PATH = "**/config/**/config.json";
function removeComponent() {
    return __awaiter(this, void 0, void 0, function* () {
        try {
            const wsFolder = vscode.workspace.rootPath;
            if (wsFolder) {
                // Check if git is enabled
                const isGitEnabled = fs.existsSync(path.join(wsFolder, '.git'));
                let proceed = false;
                if (!isGitEnabled) {
                    const confirmValue = yield vscode.window.showQuickPick(["no", "yes"], {
                        placeHolder: "You are not using GIT. Are you sure you want to remove the component?",
                        canPickMany: false
                    });
                    proceed = confirmValue === "yes";
                }
                else {
                    proceed = true;
                }
                // Check if code needs to proceed
                if (!proceed) {
                    return;
                }
                // Find all manifest files in the current project
                const manifests = yield vscode.workspace.findFiles(MANIFESTS_PATH, "node_modules");
                if (manifests && manifests.length > 0) {
                    let entries = [];
                    for (const manifest of manifests) {
                        const manifestJson = yield FileHelper_1.FileHelper.getJsonContents(manifest);
                        if (manifestJson && manifestJson.alias) {
                            entries.push({
                                path: manifest.path.replace(wsFolder, "."),
                                name: manifestJson.alias
                            });
                        }
                    }
                    if (entries.length === 0) {
                        vscode.window.showWarningMessage(`${extension_1.EXTENSION_LOG_NAME}: No component(s) found.`);
                        return;
                    }
                    // Show a prompt to ask which component to remove
                    const removeName = yield vscode.window.showQuickPick(entries.map(e => e.name), {
                        placeHolder: "Which component do you want to remove?",
                        canPickMany: false
                    });
                    // Start the file / folder deletion
                    if (!removeName) {
                        return;
                    }
                    // Get the entry to remove
                    const entryToremove = entries.find(e => e.name === removeName);
                    if (!entryToremove) {
                        return;
                    }
                    // Retrieve the global configuration file
                    const configFiles = yield vscode.workspace.findFiles(CONFIG_PATH, "node_modules", 1);
                    if (configFiles && configFiles.length > 0) {
                        const config = configFiles[0];
                        const configJson = yield FileHelper_1.FileHelper.getJsonContents(config);
                        if (!configJson) {
                            return;
                        }
                        // Which entry to update
                        const keysToDelete = [];
                        for (const bundleName in configJson.bundles) {
                            let bundle = configJson.bundles[bundleName];
                            // Filter out the components to remove
                            bundle.components = bundle.components.filter(c => c.manifest.toLowerCase() !== entryToremove.path.toLowerCase());
                            if (bundle.components.length === 0) {
                                keysToDelete.push(bundleName);
                            }
                        }
                        // Check if there are bundle entries which can be removed
                        for (const keyToDelete of keysToDelete) {
                            delete configJson.bundles[keyToDelete];
                        }
                        // Retrieve the component its path
                        const folderPath = entryToremove.path.substring(0, entryToremove.path.lastIndexOf("/"));
                        // Check if component had a manifest file
                        let libPath = folderPath.replace("src/", "lib/");
                        if (libPath.startsWith(".")) {
                            libPath = libPath.substring(1);
                        }
                        if (libPath.startsWith("/")) {
                            libPath = libPath.substring(1);
                        }
                        let localesToDelete = [];
                        for (const locale in configJson.localizedResources) {
                            const resourcePath = configJson.localizedResources[locale];
                            if (resourcePath.includes(libPath)) {
                                localesToDelete.push(locale);
                            }
                        }
                        for (const localeToDelete of localesToDelete) {
                            delete configJson.localizedResources[localeToDelete];
                        }
                        // Store the updated config
                        fs.writeFileSync(config.path, JSON.stringify(configJson, null, 2));
                        // Remove the component folder
                        const absPath = path.join(wsFolder, folderPath);
                        if (fs.existsSync(absPath)) {
                            rimraf.sync(absPath);
                            vscode.window.showInformationMessage(`${extension_1.EXTENSION_LOG_NAME}: Component ${removeName} was successfully removed.`);
                        }
                    }
                }
                else {
                    vscode.window.showWarningMessage(`${extension_1.EXTENSION_LOG_NAME}: No component manifest files found.`);
                    return;
                }
            }
        }
        catch (e) {
            vscode.window.showErrorMessage(`${extension_1.EXTENSION_LOG_NAME}: Sorry, something went wrong.`);
            console.error(e);
        }
    });
}
exports.removeComponent = removeComponent;
//# sourceMappingURL=Removal.js.map