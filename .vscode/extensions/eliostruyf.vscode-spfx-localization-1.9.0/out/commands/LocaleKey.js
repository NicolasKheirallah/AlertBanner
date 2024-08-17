"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.LocaleKey = void 0;
const ExtensionSettings_1 = require("./../helpers/ExtensionSettings");
const vscode = require("vscode");
const ActionType_1 = require("./ActionType");
const ProjectFileHelper_1 = require("../helpers/ProjectFileHelper");
const ResourceHelper_1 = require("../helpers/ResourceHelper");
const TextHelper_1 = require("../helpers/TextHelper");
const Logging_1 = require("./Logging");
const ExtensionSettings_2 = require("../helpers/ExtensionSettings");
const CsvCommands_1 = require("./CsvCommands");
class LocaleKey {
    /**
     * Create a new localization key for a SharePoint Framework solution
     */
    static async insert() {
        // The code you place here will be executed every time your command is executed
        let editor = vscode.window.activeTextEditor;
        if (!editor) {
            Logging_1.default.error(`You aren't editing a file at the moment.`);
            return; // No open text editor
        }
        // Get the current text selection
        let selection = editor.selection;
        let text = editor.document.getText(selection);
        if (!text) {
            Logging_1.default.error(`You didn't select a string to replace with the locale key.`);
            return;
        }
        // Check if the text start and ends width quotes
        text = TextHelper_1.default.stripQuotes(text);
        // Create the localization information
        this.createLocalization(editor, text, ActionType_1.ActionType.insert);
    }
    /**
     * Create a new key and insert it in the current document. Same process as creation, but without text selection.
     */
    static async create() {
        // The code you place here will be executed every time your command is executed
        let editor = vscode.window.activeTextEditor;
        if (!editor) {
            Logging_1.default.error(`You aren't editing a file at the moment.`);
            return; // No open text editor
        }
        // Create the localization information
        this.createLocalization(editor, "", ActionType_1.ActionType.create);
    }
    /**
     * Import a localization dependency in the local file of a SharePoint Framework solution
     */
    static async import() {
        let editor = vscode.window.activeTextEditor;
        if (!editor) {
            return; // No open text editor
        }
        // Get the current text of the document
        const crntFile = editor.document.getText();
        if (crntFile && crntFile.includes("import * as strings")) {
            Logging_1.default.warning(`Current file already contains a localized resources strings import.`);
            return;
        }
        // Fetch the project config
        const configInfo = await ProjectFileHelper_1.default.getConfig();
        if (configInfo && configInfo.localizedResources) {
            const resx = ResourceHelper_1.default.excludeResourcePaths(configInfo);
            // Check if resources were retrieved
            if (resx && resx.length > 0) {
                // Take the default one to import
                let defaultResx = resx[0].key;
                if (resx.length > 1) {
                    defaultResx = await vscode.window.showQuickPick(resx.map(r => r.key), {
                        placeHolder: "Specify which localized resource you want to insert in your file.",
                        canPickMany: false
                    });
                }
                if (defaultResx) {
                    editor.edit(builder => builder.insert(new vscode.Position(0, 0), `import * as strings from '${defaultResx}';\r\n`));
                }
                else {
                    Logging_1.default.error("You didn't select a localized resource to insert.");
                }
            }
        }
    }
    /**
     * Creates the localization keys and values in the right files
     *
     * @param editor
     * @param text
     * @param action
     */
    static async createLocalization(editor, text, action) {
        const localeKey = await vscode.window.showInputBox({
            ignoreFocusOut: true,
            placeHolder: "Specify the key to create",
            prompt: "Example: InputTitleLabel, TitleFieldLabel, ..."
        });
        if (!localeKey) {
            Logging_1.default.error(`You didn't specify a locale key to create.`);
            return;
        }
        // Check if text is empty.
        if (!text && action === ActionType_1.ActionType.create) {
            const localeValue = await vscode.window.showInputBox({
                ignoreFocusOut: true,
                placeHolder: "Specify the default localization value",
                prompt: "Example: Loading profile information..."
            });
            if (!localeValue) {
                Logging_1.default.error(`You didn't specify the default localization value.`);
                return;
            }
            else {
                text = localeValue;
            }
        }
        // Check if the user wants to surround the key with curly brackets. Only during the insert process.
        let useBrackets = "no";
        const bracketsResult = await vscode.window.showQuickPick(["no", "yes"], {
            placeHolder: "Do you want to surround the localized key with curly brackets `{}`?",
            canPickMany: false
        });
        if (bracketsResult) {
            useBrackets = bracketsResult;
        }
        // Fetch the project config
        const configInfo = await ProjectFileHelper_1.default.getConfig();
        if (configInfo) {
            if (!configInfo.localizedResources) {
                Logging_1.default.error(`No localizedResources were defined in the config.`);
                return;
            }
            // Convert to array and filter out the none project related resource files
            const resx = ResourceHelper_1.default.excludeResourcePaths(configInfo);
            let defaultResx = null;
            if (resx && resx.length > 0) {
                // Fetch the default locale resource
                defaultResx = resx[0];
                // Check if there were more localized resources defined
                if (resx.length > 1) {
                    // Show the quick pick control with the available options
                    const resxKey = await vscode.window.showQuickPick(resx.map(r => r.key), {
                        placeHolder: "Specify which localized resource file to use.",
                        canPickMany: false
                    });
                    // Check if option was selected
                    if (resxKey) {
                        const selected = resx.filter(r => r.key === resxKey);
                        defaultResx = selected && selected.length > 0 ? selected[0] : defaultResx;
                    }
                }
                // Create the key in the localized resource file
                let resourcePath = ProjectFileHelper_1.default.getResourcePath(defaultResx);
                // Get all files from the localization folder
                const localeFiles = await vscode.workspace.findFiles(`${resourcePath}/*`);
                // Loop over all the files
                for (const filePath of localeFiles) {
                    await this.addKeyToFile(filePath, localeKey, text);
                }
                // Insert the newly created key on the insert action
                if (action === ActionType_1.ActionType.insert) {
                    // Update the current selected text to the used resouce key
                    await editor.edit(builder => {
                        builder.replace(editor.selection, useBrackets === "yes" ? `{strings.${localeKey}}` : `strings.${localeKey}`);
                    });
                }
                else {
                    if (editor.selection.active) {
                        // Update the current selected text to the used resouce key
                        await editor.edit(builder => {
                            builder.replace(editor.selection.active, useBrackets === "yes" ? `{strings.${localeKey}}` : `strings.${localeKey}`);
                        });
                    }
                }
            }
            // Display a message box to the user
            // vscode.window.showInformationMessage(`${EXTENSION_NAME}: "${localeKey}" key has been added.`);
            // Check if auto CSV export needs to start
            const autoExport = vscode.workspace.getConfiguration(ExtensionSettings_2.CONFIG_KEY).get(ExtensionSettings_1.CONFIG_AUTO_EXPORT);
            if (autoExport && defaultResx) {
                // Start the export to the CSV file
                CsvCommands_1.default.export(defaultResx);
            }
        }
    }
    /**
     * Adds the locale key to the found file
     *
     * @param fileName
     * @param localeKey
     * @param localeValue
     */
    static async addKeyToFile(fileName, localeKey, localeValue) {
        const fileData = await vscode.workspace.openTextDocument(fileName);
        const fileContents = fileData.getText();
        const fileLines = fileContents.split("\n");
        // Create workspace edit
        const edit = new vscode.WorkspaceEdit();
        let applyEdit = false;
        // Check if the key is already in place
        if (fileContents.includes(localeKey)) {
            Logging_1.default.warning(`The key (${localeKey}) was already defined in the following file: ${fileData.fileName}.`);
            return;
        }
        let idx = -1;
        // Check if "d.ts" file
        if (fileData.fileName.endsWith(".d.ts")) {
            idx = TextHelper_1.default.findInsertPosition(fileLines, localeKey, TextHelper_1.default.findPositionTs);
        }
        // Check if "js" file
        if (fileData.fileName.endsWith(".js") || (fileData.fileName.endsWith(".ts") && !fileData.fileName.endsWith(".d.ts"))) {
            // Check if line starts with "return" and ends with "{"
            idx = TextHelper_1.default.findInsertPosition(fileLines, localeKey, TextHelper_1.default.findPositionJs);
        }
        // Check if the line was found, add the key and save the file
        if (idx !== -1) {
            applyEdit = true;
            const getLinePos = fileLines[idx + 1].search(/\S|$/);
            // Create the data to insert in the file
            let newLineData = null;
            if (fileData.fileName.endsWith(".d.ts")) {
                newLineData = `${localeKey}: string;\r\n${' '.repeat(getLinePos)}`;
            }
            else if (fileData.fileName.endsWith(".js") || (fileData.fileName.endsWith(".ts") && !fileData.fileName.endsWith(".d.ts"))) {
                newLineData = `${localeKey}: "${localeValue.replace(/"/g, `\\"`)}",\r\n${' '.repeat(getLinePos)}`;
            }
            // Check if there is data to insert
            if (newLineData) {
                edit.insert(fileData.uri, new vscode.Position((idx + 1), getLinePos), newLineData);
            }
        }
        if (applyEdit) {
            try {
                const result = await vscode.workspace.applyEdit(edit).then(success => success);
                if (!result) {
                    Logging_1.default.error(`Couldn't add the key to the file: ${fileName}.`);
                }
                else {
                    await fileData.save();
                }
            }
            catch (e) {
                Logging_1.default.error(`Something went wrong adding the locale key to the file: ${fileName}.`);
            }
        }
        return;
    }
}
exports.LocaleKey = LocaleKey;
//# sourceMappingURL=LocaleKey.js.map