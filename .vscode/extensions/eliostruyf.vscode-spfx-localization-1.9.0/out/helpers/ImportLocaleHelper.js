"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const vscode = require("vscode");
const fs = require("fs");
const path = require("path");
const Logging_1 = require("../commands/Logging");
const ProjectFileHelper_1 = require("./ProjectFileHelper");
const ExtensionSettings_1 = require("./ExtensionSettings");
const TextHelper_1 = require("./TextHelper");
class ImportLocaleHelper {
    /**
     * Create the locale files
     *
     * @param csvData
     * @param resx
     */
    static async createLocaleFiles(resx, localeData) {
        if (!resx || !localeData) {
            return;
        }
        let fileExtension = vscode.workspace.getConfiguration(ExtensionSettings_1.CONFIG_KEY).get(ExtensionSettings_1.CONFIG_FILE_EXTENSION);
        if (!fileExtension) {
            fileExtension = "js";
        }
        // Create the key in the localized resource file
        let resourcePath = ProjectFileHelper_1.default.getResourcePath(resx);
        // Start creating the files
        for (const key in localeData) {
            const localLabels = localeData[key];
            if (key && localLabels && localLabels.length > 0) {
                const resourceKeys = localeData[key].filter(l => l.resx === resx.key);
                if (resourceKeys && resourceKeys.length > 0) {
                    await this.ensureTypescriptKeysDefined(resx, localLabels);
                    // Create the file content
                    let fileContents = fileExtension === "ts" ? `declare var define: any;
       
define([], () => {
` : `define([], function() {
`;
                    fileContents += `  return {
    ${resourceKeys.map(k => `${k.key}: "${k.label}"`).join(`,\n    `)}
  };
});`;
                    // Start creating the file
                    const fileLocation = path.join(vscode.workspace.rootPath || __dirname, resourcePath, `${key}.${fileExtension}`);
                    fs.writeFileSync(fileLocation, fileContents, { encoding: "utf8" });
                    Logging_1.default.info(`Localization labels have been imported.`);
                }
            }
        }
    }
    /**
     * Ensure all lables are inserted in the definition (.d.ts) file when importing csv
     *
     * @param resx
     * @param localLabels
    */
    static async ensureTypescriptKeysDefined(resx, localLabels) {
        // Create the key in the localized resource file
        let resourcePath = ProjectFileHelper_1.default.getResourcePath(resx);
        // Get all files from the localization folder
        const definitionFiles = await vscode.workspace.findFiles(`${resourcePath}/*.d.ts`);
        // nothing to update
        if (definitionFiles.length === 0) {
            return;
        }
        if (definitionFiles.length > 1) {
            Logging_1.default.warning(`There is more than one typescript definition file (.d.ts), the update skipped.`);
            return;
        }
        const fileData = await vscode.workspace.openTextDocument(definitionFiles[0]);
        const fileName = fileData.fileName;
        const fileContents = fileData.getText();
        const fileLines = fileContents.split("\n");
        // Create workspace edit
        const edit = new vscode.WorkspaceEdit();
        const startPos = fileLines.findIndex(line => {
            const matches = line.trim().match(/(^declare interface|{$)/gi);
            return matches !== null && matches.length >= 2;
        });
        // the file is non-standard
        if (startPos === -1) {
            Logging_1.default.warning(`The file ${fileName} does not start with 'declare interface'. File updated skipped.`);
            return;
        }
        let applyEdit = false;
        for (let localLabel of localLabels) {
            const localeKey = localLabel.key;
            // Check if the line was found, add the key and save the file
            if (!fileContents.includes(`${localeKey}: string;`)) {
                applyEdit = true;
                const getLine = TextHelper_1.default.findInsertPosition(fileLines, localeKey, TextHelper_1.default.findPositionTs);
                const getLinePos = fileLines[getLine + 1].search(/\S|$/);
                // Create the data to insert in the file
                const newLineData = `${localeKey}: string;\r\n${' '.repeat(getLinePos)}`;
                edit.insert(fileData.uri, new vscode.Position(getLine + 1, getLinePos), newLineData);
            }
        }
        if (applyEdit) {
            try {
                const result = await vscode.workspace.applyEdit(edit).then(success => success);
                if (!result) {
                    Logging_1.default.warning(`Couldn't update the typescript definition file: ${fileName}.`);
                }
                else {
                    await fileData.save();
                }
            }
            catch (e) {
                Logging_1.default.warning(`Something went wrong when updating the typescript definition file: ${fileName}.`);
            }
        }
    }
}
exports.default = ImportLocaleHelper;
//# sourceMappingURL=ImportLocaleHelper.js.map