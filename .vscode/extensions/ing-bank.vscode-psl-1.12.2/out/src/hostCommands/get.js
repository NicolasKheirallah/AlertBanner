"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.getCompiledCodeHandler = exports.getTableHandler = exports.getElementHandler = void 0;
const vscode = require("vscode");
const utils = require("./hostCommandUtils");
const path = require("path");
const fs = require("fs-extra");
const environment = require("../common/environment");
const mumps_1 = require("../language/mumps");
const icon = "\u21E9" /* utils.icons.GET */;
function getElementHandler(context) {
    return __awaiter(this, void 0, void 0, function* () {
        let c = utils.getFullContext(context);
        if (c.mode === 2 /* utils.ContextMode.DIRECTORY */) {
            let input = yield promptUserForComponent();
            if (input)
                return getElement(path.join(c.fsPath, input)).catch(() => { });
        }
        else if (c.mode === 1 /* utils.ContextMode.FILE */) {
            let workspace = vscode.workspace.getWorkspaceFolder(vscode.Uri.file(c.fsPath));
            if (!workspace) {
                // skeptical of this approach
                return;
            }
            let input = yield promptUserForComponent();
            if (!input)
                return;
            let extension = path.extname(input).replace('.', '');
            let description = utils.extensionToDescription[extension];
            let filters = {};
            filters[description] = [extension];
            let currentFolder = vscode.workspace.getWorkspaceFolder(vscode.Uri.file(c.fsPath));
            if (!currentFolder)
                return;
            let target;
            let defaultDir = DIR_MAPPINGS[extension];
            if (defaultDir) {
                target = { fsPath: path.join(currentFolder.uri.fsPath, defaultDir, input) };
            }
            else {
                let defaultUri = vscode.Uri.file(path.join(currentFolder.uri.fsPath, input));
                target = yield vscode.window.showSaveDialog({ defaultUri, filters: filters });
            }
            if (!target)
                return;
            return getElement(target.fsPath).catch(() => { });
        }
        else {
            let quickPick = yield environment.workspaceQuickPick();
            if (!quickPick)
                return;
            let chosenEnv = quickPick;
            let input = yield promptUserForComponent();
            if (!input)
                return;
            let extension = path.extname(input).replace('.', '');
            let description = utils.extensionToDescription[extension];
            let filters = {};
            filters[description] = [extension];
            let target;
            let defaultDir = DIR_MAPPINGS[extension];
            if (defaultDir) {
                target = { fsPath: path.join(chosenEnv.fsPath, defaultDir, input) };
            }
            else {
                let defaultUri = vscode.Uri.file(path.join(chosenEnv.fsPath, input));
                target = yield vscode.window.showSaveDialog({ defaultUri, filters: filters });
            }
            if (!target)
                return;
            return getElement(target.fsPath).catch(() => { });
        }
        return;
    });
}
exports.getElementHandler = getElementHandler;
function getTableHandler(context) {
    return __awaiter(this, void 0, void 0, function* () {
        let c = utils.getFullContext(context);
        if (c.mode === 2 /* utils.ContextMode.DIRECTORY */) {
            let input = yield promptUserForTable();
            if (input)
                return getTable(input, c.fsPath, c.fsPath).catch(() => { });
        }
        else if (c.mode === 1 /* utils.ContextMode.FILE */) {
            let workspace = vscode.workspace.getWorkspaceFolder(vscode.Uri.file(c.fsPath));
            if (!workspace) {
                // skeptical of this approach
                return;
            }
            let tableName = yield promptUserForTable();
            if (!tableName)
                return;
            let tableDir = DIR_MAPPINGS['TABLE'];
            let target;
            if (tableDir) {
                target = [{ fsPath: path.join(workspace.uri.fsPath, tableDir) }];
            }
            else {
                target = yield vscode.window.showOpenDialog({ defaultUri: workspace.uri, canSelectFiles: false, canSelectFolders: true, canSelectMany: false, filters: { 'Table Directory': [] } });
            }
            if (!target)
                return;
            return getTable(tableName, target[0].fsPath, workspace.uri.fsPath).catch(() => { });
        }
        else {
            let quickPick = yield environment.workspaceQuickPick();
            if (!quickPick)
                return;
            let chosenEnv = quickPick;
            let tableName = yield promptUserForTable();
            if (!tableName)
                return;
            let tableDir = DIR_MAPPINGS['TABLE'];
            let target;
            if (tableDir) {
                target = [{ fsPath: path.join(chosenEnv.description, tableDir) }];
            }
            else {
                target = yield vscode.window.showOpenDialog({ defaultUri: vscode.Uri.file(chosenEnv.description), canSelectFiles: false, canSelectFolders: true, canSelectMany: false, filters: { 'Table Directory': [] } });
            }
            if (!target)
                return;
            return getTable(tableName, target[0].fsPath, chosenEnv.description).catch(() => { });
        }
        return;
    });
}
exports.getTableHandler = getTableHandler;
function getCompiledCodeHandler(context) {
    return __awaiter(this, void 0, void 0, function* () {
        let c = utils.getFullContext(context);
        if (c.mode === 1 /* utils.ContextMode.FILE */) {
            return getCompiledCode(c.fsPath).catch(() => { });
        }
        else if (c.mode === 2 /* utils.ContextMode.DIRECTORY */) {
            let files = yield vscode.window.showOpenDialog({ defaultUri: vscode.Uri.file(c.fsPath), canSelectMany: true, openLabel: 'Refresh' });
            if (!files)
                return;
            for (let fsPath of files.map(file => file.fsPath)) {
                yield getCompiledCode(fsPath).catch(() => { });
            }
        }
        else {
            let quickPick = yield environment.workspaceQuickPick();
            if (!quickPick)
                return;
            let chosenEnv = quickPick;
            let files = yield vscode.window.showOpenDialog({ defaultUri: vscode.Uri.file(chosenEnv.fsPath), canSelectMany: true, openLabel: 'Refresh' });
            if (!files)
                return;
            for (let fsPath of files.map(file => file.fsPath)) {
                yield getCompiledCode(fsPath).catch(() => { });
            }
        }
        return;
    });
}
exports.getCompiledCodeHandler = getCompiledCodeHandler;
function getCompiledCode(fsPath) {
    return __awaiter(this, void 0, void 0, function* () {
        if (!fs.statSync(fsPath).isFile())
            return;
        let env;
        const routineName = `${path.basename(fsPath).split('.')[0]}.m`;
        return utils.executeWithProgress(`${icon} ${path.basename(fsPath)} GET`, () => __awaiter(this, void 0, void 0, function* () {
            let envs;
            try {
                envs = yield utils.getEnvironment(fsPath);
            }
            catch (e) {
                utils.logger.error(`${"\u274C" /* utils.icons.ERROR */} ${icon} Invalid environment configuration.`);
                return;
            }
            if (envs.length === 0) {
                utils.logger.error(`${"\u274C" /* utils.icons.ERROR */} ${icon} No environments selected.`);
                return;
            }
            let choice = yield utils.getCommandenvConfigQuickPick(envs);
            if (!choice)
                return;
            env = choice;
            utils.logger.info(`${"\u2026" /* utils.icons.WAIT */} ${icon} ${routineName} GET COMPILED from ${env.name}`);
            let doc = yield vscode.workspace.openTextDocument(fsPath);
            yield doc.save();
            let connection = yield utils.getConnection(env);
            let output = yield connection.get(routineName);
            const uri = vscode.Uri.parse(`${mumps_1.MumpsVirtualDocument.schemes.compiled}:/${env.name}/${routineName}`);
            const virtualDocument = new mumps_1.MumpsVirtualDocument(routineName, output, uri);
            utils.logger.info(`${"\u2714" /* utils.icons.SUCCESS */} ${icon} ${routineName} GET COMPILED from ${env.name} succeeded`);
            connection.close();
            vscode.window.showTextDocument(virtualDocument.uri, { preview: false });
        })).catch((e) => {
            if (env && env.name) {
                utils.logger.error(`${"\u274C" /* utils.icons.ERROR */} ${icon} error in ${env.name} ${e.message}`);
            }
            else {
                utils.logger.error(`${"\u274C" /* utils.icons.ERROR */} ${icon} ${e.message}`);
            }
        });
    });
}
function getElement(fsPath) {
    return __awaiter(this, void 0, void 0, function* () {
        let env;
        yield utils.executeWithProgress(`${icon} ${path.basename(fsPath)} GET`, () => __awaiter(this, void 0, void 0, function* () {
            let envs;
            try {
                envs = yield utils.getEnvironment(fsPath);
            }
            catch (e) {
                utils.logger.error(`${"\u274C" /* utils.icons.ERROR */} ${icon} Invalid environment configuration.`);
                return;
            }
            if (envs.length === 0) {
                utils.logger.error(`${"\u274C" /* utils.icons.ERROR */} ${icon} No environments selected.`);
                return;
            }
            let choice = yield utils.getCommandenvConfigQuickPick(envs);
            if (!choice)
                return;
            env = choice;
            utils.logger.info(`${"\u2026" /* utils.icons.WAIT */} ${icon} ${path.basename(fsPath)} GET from ${env.name}`);
            let connection = yield utils.getConnection(env);
            let output = yield connection.get(fsPath);
            yield fs.ensureDir(path.dirname(fsPath));
            yield utils.writeFileWithSettings(fsPath, output);
            utils.logger.info(`${"\u2714" /* utils.icons.SUCCESS */} ${icon} ${path.basename(fsPath)} GET from ${env.name} succeeded`);
            connection.close();
            yield vscode.workspace.openTextDocument(fsPath).then(vscode.window.showTextDocument);
        })).catch((e) => {
            if (env && env.name) {
                utils.logger.error(`${"\u274C" /* utils.icons.ERROR */} ${icon} error in ${env.name} ${e.message}`);
            }
            else {
                utils.logger.error(`${"\u274C" /* utils.icons.ERROR */} ${icon} ${e.message}`);
            }
        });
        return;
    });
}
function getTable(tableName, targetDirectory, workpacePath) {
    return __awaiter(this, void 0, void 0, function* () {
        let env;
        yield utils.executeWithProgress(`${icon} ${tableName} TABLE GET`, () => __awaiter(this, void 0, void 0, function* () {
            let envs;
            try {
                envs = yield utils.getEnvironment(workpacePath);
            }
            catch (e) {
                utils.logger.error(`${"\u274C" /* utils.icons.ERROR */} ${icon} Invalid environment configuration.`);
                return;
            }
            if (envs.length === 0) {
                utils.logger.error(`${"\u274C" /* utils.icons.ERROR */} ${icon} No environments selected.`);
                return;
            }
            let choice = yield utils.getCommandenvConfigQuickPick(envs);
            if (!choice)
                return;
            env = choice;
            utils.logger.info(`${"\u2026" /* utils.icons.WAIT */} ${icon} ${tableName} TABLE GET from ${env.name}`);
            let connection = yield utils.getConnection(env);
            let output = yield connection.getTable(tableName.toUpperCase() + '.TBL');
            yield fs.ensureDir(path.join(targetDirectory, tableName.toLowerCase()));
            let tableFiles = (yield fs.readdir(targetDirectory)).filter(f => f.startsWith(tableName));
            for (let file of tableFiles) {
                yield fs.remove(file);
            }
            const promises = output.split(String.fromCharCode(0)).map(content => {
                const contentArray = content.split(String.fromCharCode(1));
                const fileName = contentArray[0];
                const fileContent = contentArray[1];
                return utils.writeFileWithSettings(path.join(targetDirectory, tableName.toLowerCase(), fileName), fileContent);
            });
            yield Promise.all(promises);
            utils.logger.info(`${"\u2714" /* utils.icons.SUCCESS */} ${icon} ${tableName} TABLE GET from ${env.name} succeeded`);
            connection.close();
        })).catch((e) => {
            if (env && env.name) {
                utils.logger.error(`${"\u274C" /* utils.icons.ERROR */} ${icon} error in ${env.name} ${e.message}`);
            }
            else {
                utils.logger.error(`${"\u274C" /* utils.icons.ERROR */} ${icon} ${e.message}`);
            }
        });
        return;
    });
}
function promptUserForComponent() {
    return __awaiter(this, void 0, void 0, function* () {
        let inputOptions = {
            prompt: 'Name of Component (with extension)', validateInput: (input) => {
                if (!input)
                    return;
                let extension = path.extname(input) ? path.extname(input).replace('.', '') : 'No extension';
                if (extension in utils.extensionToDescription)
                    return '';
                return `Invalid extension (${extension})`;
            }
        };
        return vscode.window.showInputBox(inputOptions);
    });
}
function promptUserForTable() {
    return __awaiter(this, void 0, void 0, function* () {
        let inputOptions = {
            prompt: 'Name of Table (no extension)',
            validateInput: (value) => {
                if (!value)
                    return;
                if (value.includes('.'))
                    return 'Do not include the extension';
            }
        };
        return vscode.window.showInputBox(inputOptions);
    });
}
const DIR_MAPPINGS = {
    'BATCH': 'dataqwik/batch',
    'COL': '',
    'DAT': 'data',
    'FKY': 'dataqwik/foreign_key',
    // 'G': 'Global',
    'IDX': 'dataqwik/index',
    'JFD': 'dataqwik/journal',
    'm': 'routine',
    'PPL': '',
    'PROC': 'dataqwik/procedure',
    'properties': 'property',
    'PSL': '',
    'psl': '',
    'pslx': '',
    'pslxtra': '',
    'psql': '',
    'QRY': 'dataqwik/query',
    'RPT': 'dataqwik/report',
    'SCR': 'dataqwik/screen',
    // TABLE not supported
    'TABLE': 'dataqwik/table',
    'TBL': '',
    'TRIG': 'dataqwik/trigger',
};
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoiZ2V0LmpzIiwic291cmNlUm9vdCI6IiIsInNvdXJjZXMiOlsiLi4vLi4vLi4vc3JjL2hvc3RDb21tYW5kcy9nZXQudHMiXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6Ijs7Ozs7Ozs7Ozs7O0FBQUEsaUNBQWlDO0FBQ2pDLDRDQUE0QztBQUM1Qyw2QkFBNkI7QUFDN0IsK0JBQStCO0FBQy9CLHFEQUFxRDtBQUNyRCw2Q0FBeUQ7QUFFekQsTUFBTSxJQUFJLGlDQUFrQixDQUFDO0FBRTdCLFNBQXNCLGlCQUFpQixDQUFDLE9BQXNDOztRQUM3RSxJQUFJLENBQUMsR0FBRyxLQUFLLENBQUMsY0FBYyxDQUFDLE9BQU8sQ0FBQyxDQUFDO1FBQ3RDLElBQUksQ0FBQyxDQUFDLElBQUksd0NBQWdDLEVBQUU7WUFDM0MsSUFBSSxLQUFLLEdBQUcsTUFBTSxzQkFBc0IsRUFBRSxDQUFDO1lBQzNDLElBQUksS0FBSztnQkFBRSxPQUFPLFVBQVUsQ0FBQyxJQUFJLENBQUMsSUFBSSxDQUFDLENBQUMsQ0FBQyxNQUFNLEVBQUUsS0FBSyxDQUFDLENBQUMsQ0FBQyxLQUFLLENBQUMsR0FBRyxFQUFFLEdBQUcsQ0FBQyxDQUFDLENBQUM7U0FDMUU7YUFDSSxJQUFJLENBQUMsQ0FBQyxJQUFJLG1DQUEyQixFQUFFO1lBQzNDLElBQUksU0FBUyxHQUFHLE1BQU0sQ0FBQyxTQUFTLENBQUMsa0JBQWtCLENBQUMsTUFBTSxDQUFDLEdBQUcsQ0FBQyxJQUFJLENBQUMsQ0FBQyxDQUFDLE1BQU0sQ0FBQyxDQUFDLENBQUM7WUFDL0UsSUFBSSxDQUFDLFNBQVMsRUFBRTtnQkFDZiw2QkFBNkI7Z0JBQzdCLE9BQU87YUFDUDtZQUNELElBQUksS0FBSyxHQUFHLE1BQU0sc0JBQXNCLEVBQUUsQ0FBQztZQUMzQyxJQUFJLENBQUMsS0FBSztnQkFBRSxPQUFPO1lBQ25CLElBQUksU0FBUyxHQUFHLElBQUksQ0FBQyxPQUFPLENBQUMsS0FBSyxDQUFDLENBQUMsT0FBTyxDQUFDLEdBQUcsRUFBRSxF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