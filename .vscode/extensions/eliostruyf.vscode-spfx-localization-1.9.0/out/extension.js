'use strict';
Object.defineProperty(exports, "__esModule", { value: true });
exports.deactivate = exports.activate = void 0;
// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
const vscode = require("vscode");
const LocaleKey_1 = require("./commands/LocaleKey");
// import LanguageHover from './hover/LanguageHover';
const CsvCommands_1 = require("./commands/CsvCommands");
// this method is called when your extension is activated
// your extension is activated the very first time the command is executed
function activate(context) {
    // Register the localization command
    const creating = vscode.commands.registerCommand('extension.spfxLocalizationCreateKey', () => {
        LocaleKey_1.LocaleKey.create();
    });
    const inserting = vscode.commands.registerCommand('extension.spfxLocalizationInsertKey', () => {
        LocaleKey_1.LocaleKey.insert();
    });
    // Register the localization importer
    const importing = vscode.commands.registerCommand('extension.spfxLocalizationImport', () => {
        LocaleKey_1.LocaleKey.import();
    });
    // Register the localization importer
    const csvImport = vscode.commands.registerCommand('extension.spfxCsvImport', () => {
        CsvCommands_1.default.import();
    });
    // Register the localization importer
    const csvExport = vscode.commands.registerCommand('extension.spfxCsvExport', () => {
        CsvCommands_1.default.export();
    });
    // Register hover providers
    // vscode.languages.registerHoverProvider({ scheme: 'file', language: 'typescript' }, { provideHover: LanguageHover.onHover });
    // vscode.languages.registerHoverProvider({ scheme: 'file', language: 'typescriptreact' }, { provideHover: LanguageHover.onHover });
    context.subscriptions.push(creating);
    context.subscriptions.push(inserting);
    context.subscriptions.push(importing);
    context.subscriptions.push(csvImport);
    context.subscriptions.push(csvExport);
    // Show the actions in the context menu
    vscode.commands.executeCommand('setContext', 'spfxProjectCheck', true);
    console.log('SPFx localization is now active!');
}
exports.activate = activate;
// this method is called when your extension is deactivated
function deactivate() { }
exports.deactivate = deactivate;
//# sourceMappingURL=extension.js.map