"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const vscode = require("vscode");
const path = require("path");
const ProjectFileHelper_1 = require("../helpers/ProjectFileHelper");
const ResourceHelper_1 = require("../helpers/ResourceHelper");
class LanguageHover {
    static async onHover(document, position, token) {
        // Retrieve the word the user is currently hovering
        const wordRange = await document.getWordRangeAtPosition(position);
        const word = document.getText(wordRange);
        // Check if a word has been found
        if (word) {
            // Get the project config file
            const config = await ProjectFileHelper_1.default.getConfig();
            if (config) {
                // Get only the project resources
                const resx = ResourceHelper_1.default.excludeResourcePaths(config);
                if (resx && resx.length > 0) {
                    const hoverValues = [];
                    // Loop over the resource files
                    for (const resource of resx) {
                        let crntResxAdded = false;
                        // Get the path
                        let resourcePath = resource.value.substring(0, resource.value.lastIndexOf('/'));
                        // Use the src directory
                        if (resourcePath.startsWith("lib/")) {
                            resourcePath = resourcePath.replace("lib/", "src/");
                        }
                        // Get all files from the localization folder
                        const jsFiles = await vscode.workspace.findFiles(`${resourcePath}/*.js`);
                        // Loop over the files to see the 
                        for (const jsFile of jsFiles) {
                            if (jsFile) {
                                const fileData = await vscode.workspace.openTextDocument(jsFile);
                                const fileContents = fileData.getText();
                                const fileName = path.basename(fileData.fileName);
                                const localeName = fileName.split('.').slice(0, -1).join('.');
                                // Process the file with the hovered word
                                const value = ResourceHelper_1.default.getResourceValue(fileContents, word);
                                if (value) {
                                    if (!crntResxAdded) {
                                        // Add extra line break if it is not the first line
                                        if (hoverValues.length !== 0) {
                                            hoverValues.push(`\n`);
                                        }
                                        hoverValues.push(`**${resource.key}**\n`);
                                        crntResxAdded = true;
                                    }
                                    hoverValues.push(`- **${localeName}**: "${value}"`);
                                }
                            }
                        }
                    }
                    // Check if something needs to be added for the hover panel
                    if (hoverValues && hoverValues.length > 0) {
                        return new vscode.Hover(hoverValues.join("\n"));
                    }
                }
            }
        }
    }
}
exports.default = LanguageHover;
//# sourceMappingURL=LanguageHover.js.map