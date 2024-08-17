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
exports.GlobalFile = exports.WorkspaceFile = exports.workspaceQuickPick = exports.activate = void 0;
const vscode = require("vscode");
const fs = require("fs-extra");
const path = require("path");
const jsonc = require("jsonc-parser");
const os = require("os");
const configEnvCommand = 'psl.configureEnvironment';
const LOCAL_ENV_DIR = path.join('.vscode', 'environment.json');
const statusBar = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Right, 900);
statusBar.command = configEnvCommand;
function activate(context) {
    context.subscriptions.push(vscode.commands.registerCommand(configEnvCommand, configureEnvironmentHandler));
    context.subscriptions.push(statusBar);
    context.subscriptions.push(vscode.window.onDidChangeActiveTextEditor((e) => changeTextEditorHandler(e)));
    changeTextEditorHandler(vscode.window.activeTextEditor);
}
exports.activate = activate;
function workspaceQuickPick() {
    return __awaiter(this, void 0, void 0, function* () {
        try {
            yield GlobalFile.read();
        }
        catch (e) {
            let defaultConfig = { environments: [{ name: '', host: '', port: 0, user: '', password: '', sshLogin: '', serverType: 'SCA$IBS', encoding: 'utf8' }] };
            yield GlobalFile.write(defaultConfig);
            yield GlobalFile.show();
            return;
        }
        if (!vscode.workspace.workspaceFolders)
            return;
        let workspaceFolders = vscode.workspace.workspaceFolders;
        let items = yield Promise.all(workspaceFolders.map((folder) => __awaiter(this, void 0, void 0, function* () {
            let name;
            try {
                let envObjects = yield new WorkspaceFile(folder.uri.fsPath).environmentObjects;
                if (envObjects.length === 1) {
                    name = '\u00a0 \u00a0 $(server) ' + envObjects[0].name;
                }
                else if (envObjects.length > 1) {
                    name = '\u00a0 \u00a0 $(server) ' + envObjects.map(e => e.name).join(', ');
                }
                else {
                    name = '\u00a0 \u00a0 Not configured';
                }
            }
            catch (e) {
                name = '\u00a0 \u00a0 Not configured';
            }
            let item = { label: '$(file-directory) ' + folder.name, description: folder.uri.fsPath, detail: name, fsPath: folder.uri.fsPath };
            return item;
        })));
        if (items.length === 1)
            return items[0];
        let configureEnvironments = '\u270E Edit Environments...';
        items.push({ label: configureEnvironments, description: '', fsPath: '' });
        let choice = yield vscode.window.showQuickPick(items, { placeHolder: 'Select a Workspace.' });
        if (!choice)
            return;
        if (choice.label === configureEnvironments) {
            yield GlobalFile.show();
            return;
        }
        return choice;
    });
}
exports.workspaceQuickPick = workspaceQuickPick;
function configureEnvironmentHandler() {
    return __awaiter(this, void 0, void 0, function* () {
        let workspace = yield workspaceQuickPick();
        if (!workspace)
            return;
        environmentQuickPick(new WorkspaceFile(workspace.fsPath));
    });
}
function environmentQuickPick(workspaceFile) {
    return __awaiter(this, void 0, void 0, function* () {
        let choice = undefined;
        let workspaceEnvironments;
        let globalConfig;
        let names;
        try {
            globalConfig = yield GlobalFile.read();
        }
        catch (e) {
            if (e === GlobalFile.INVALID_CONFIG) {
                yield GlobalFile.show();
            }
            else {
            }
            let defaultConfig = { environments: [{ name: '', host: '', port: 0, user: '', password: '', sshLogin: '' }] };
            yield GlobalFile.write(defaultConfig);
            yield GlobalFile.show();
            return;
        }
        try {
            workspaceEnvironments = yield workspaceFile.environment;
            names = workspaceEnvironments.names;
        }
        catch (e) {
            yield workspaceFile.writeLocalEnv({ 'names': [] });
            workspaceEnvironments = yield workspaceFile.environment;
            names = workspaceEnvironments.names;
        }
        do {
            let items = globalConfig.environments.map(env => {
                if (names.indexOf(env.name) > -1) {
                    return { label: `${env.name}`, description: '✔' };
                }
                return { label: `${env.name}`, description: '' };
            });
            let configureEnvironments = '\u270E Edit Environments...';
            let back = '\u21a9 Back to Workspaces';
            items.push({ label: configureEnvironments, description: '' });
            if (vscode.workspace.workspaceFolders.length > 1) {
                items.push({ label: back, description: '' });
            }
            choice = yield vscode.window.showQuickPick(items, { placeHolder: `Enable environments for ${workspaceFile.workspaceFolder.name}` });
            if (choice) {
                if (choice.label === configureEnvironments) {
                    GlobalFile.show();
                    break;
                }
                if (choice.label === back) {
                    configureEnvironmentHandler();
                    break;
                }
                let index = names.indexOf(choice.label);
                if (index > -1) {
                    names.splice(index, 1);
                }
                else
                    names.push(choice.label);
                workspaceFile.writeLocalEnv(workspaceEnvironments);
            }
        } while (choice);
        yield changeTextEditorHandler(vscode.window.activeTextEditor);
    });
}
function changeTextEditorHandler(textEditor) {
    return __awaiter(this, void 0, void 0, function* () {
        let configureEnvironmentText = '$(server) Configure Environments';
        try {
            let workspaceFile = new WorkspaceFile(textEditor.document.fileName);
            let workspaceEnvironments = yield workspaceFile.environment;
            if (workspaceEnvironments.names.length === 0) {
                statusBar.text = configureEnvironmentText;
            }
            else if (workspaceEnvironments.names.length === 1) {
                statusBar.text = '$(server) ' + workspaceEnvironments.names[0];
            }
            else {
                statusBar.text = '$(server) ' + workspaceEnvironments.names.length + ' environments';
            }
        }
        catch (e) {
            statusBar.text = configureEnvironmentText;
        }
        statusBar.show();
    });
}
class WorkspaceFile {
    /**
     * @param {string} fsPath The file system path of the file.
     */
    constructor(fsPath) {
        /**
         * The file system path of the file.
         */
        this.workspaceFolder = undefined;
        /**
         * Contents of local environment.json
         */
        this._enviornment = undefined;
        /**
         * Environment configurations from global environments.json
         * corresponding to names in local environment.json
         */
        this._environmentObjects = undefined;
        this.fsPath = fsPath;
        if (!fsPath) {
            this.environmentPath = '';
            return;
        }
        this.workspaceFolder = vscode.workspace.getWorkspaceFolder(vscode.Uri.file(fsPath));
        if (!this.workspaceFolder) {
            this.environmentPath = '';
        }
        else {
            this.environmentPath = path.join(this.workspaceFolder.uri.fsPath, LOCAL_ENV_DIR);
        }
    }
    /**
     * Environment configurations from global environments.json
     * corresponding to names in local environment.json
     */
    get environmentObjects() {
        if (this._environmentObjects)
            return Promise.resolve(this._environmentObjects);
        return this.getEnvironmentObjects();
    }
    getEnvironmentObjects() {
        return __awaiter(this, void 0, void 0, function* () {
            let environment = yield this.environment;
            let globalEnv = yield this.getEnvironmentFromGlobalConfig(environment.names);
            this._environmentObjects = globalEnv;
            return this._environmentObjects;
        });
    }
    /**
     *
     * @param nameArray An array of names to match the names of configurations in the GlobalConfig.
     */
    getEnvironmentFromGlobalConfig(nameArray) {
        return __awaiter(this, void 0, void 0, function* () {
            let allEnvs = (yield GlobalFile.read()).environments;
            let ret = [];
            for (let name of nameArray) {
                for (let env of allEnvs) {
                    if (env.name === name) {
                        ret.push(env);
                    }
                }
            }
            return ret;
        });
    }
    /**
     * Contents of local environment.json
     */
    get environment() {
        if (this._enviornment)
            return Promise.resolve(this._enviornment);
        return fs.readFile(this.environmentPath).then((file) => __awaiter(this, void 0, void 0, function* () {
            let localEnvironment = jsonc.parse(file.toString());
            if (!localEnvironment.names || !Array.isArray(localEnvironment.names)) {
                throw new Error('Local environment.json is not properly configured.');
            }
            this._enviornment = localEnvironment;
            return localEnvironment;
        }));
    }
    writeLocalEnv(newLocalEnv) {
        return __awaiter(this, void 0, void 0, function* () {
            // TODO prune names
            yield fs.ensureFile(this.environmentPath);
            yield fs.writeFile(this.environmentPath, JSON.stringify(newLocalEnv, null, '\t'));
        });
    }
}
exports.WorkspaceFile = WorkspaceFile;
class GlobalFile {
    /**
     * Reads and returns the contents of the file.
     *
     * @throws An error if parsing fails or if improperly formatted.
     */
    static read() {
        return __awaiter(this, void 0, void 0, function* () {
            let globalConfig = jsonc.parse((yield fs.readFile(this.path)).toString());
            if (!globalConfig.environments)
                throw this.INVALID_CONFIG;
            return globalConfig;
        });
    }
    /**
     * Writes the new configuration to the file.
     *
     * @param newGlobalConfig The new configuration.
     */
    static write(newGlobalConfig) {
        return __awaiter(this, void 0, void 0, function* () {
            yield fs.ensureFile(this.path);
            yield fs.writeFile(this.path, JSON.stringify(newGlobalConfig, null, '\t'));
        });
    }
    /**
     * Shows the configuration file in the editor window.
     */
    static show() {
        return __awaiter(this, void 0, void 0, function* () {
            yield vscode.window.showTextDocument(vscode.Uri.file(this.path));
        });
    }
}
exports.GlobalFile = GlobalFile;
/**
 * Path to the global config file
 */
GlobalFile.path = (() => {
    const envFileName = 'environments.json';
    const appdata = process.env.APPDATA || (process.platform === 'darwin' ? process.env.HOME + '/Library/Application Support' : '/var/local');
    let channelPath;
    if (vscode.env.appName.indexOf('Insiders') > 0) {
        channelPath = 'Code - Insiders';
    }
    else {
        channelPath = 'Code';
    }
    let envPath = path.join(appdata, channelPath, 'User', envFileName);
    // in linux, it may not work with /var/local, then try to use /home/myuser/.config
    if ((process.platform === 'linux') && (!fs.existsSync(envPath))) {
        envPath = path.join(os.homedir(), '.config/', channelPath, 'User', envFileName);
    }
    return envPath;
})();
GlobalFile.INVALID_CONFIG = new Error('Missing environments in global config.');
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoiZW52aXJvbm1lbnQuanMiLCJzb3VyY2VSb290IjoiIiwic291cmNlcyI6WyIuLi8uLi8uLi9zcmMvY29tbW9uL2Vudmlyb25tZW50LnRzIl0sIm5hbWVzIjpbXSwibWFwcGluZ3MiOiI7Ozs7Ozs7Ozs7OztBQUFBLGlDQUFpQztBQUNqQywrQkFBK0I7QUFDL0IsNkJBQTZCO0FBQzdCLHNDQUFzQztBQUN0Qyx5QkFBeUI7QUFFekIsTUFBTSxnQkFBZ0IsR0FBRywwQkFBMEIsQ0FBQztBQUVwRCxNQUFNLGFBQWEsR0FBRyxJQUFJLENBQUMsSUFBSSxDQUFDLFNBQVMsRUFBRSxrQkFBa0IsQ0FBQyxDQUFDO0FBRS9ELE1BQU0sU0FBUyxHQUFHLE1BQU0sQ0FBQyxNQUFNLENBQUMsbUJBQW1CLENBQUMsTUFBTSxDQUFDLGtCQUFrQixDQUFDLEtBQUssRUFBRSxHQUFHLENBQUMsQ0FBQztBQUMxRixTQUFTLENBQUMsT0FBTyxHQUFHLGdCQUFnQixDQUFDO0FBRXJDLFNBQWdCLFFBQVEsQ0FBQyxPQUFnQztJQUV4RCxPQUFPLENBQUMsYUFBYSxDQUFDLElBQUksQ0FDekIsTUFBTSxDQUFDLFFBQVEsQ0FBQyxlQUFlLENBQzlCLGdCQUFnQixFQUFFLDJCQUEyQixDQUM3QyxDQUNELENBQUM7SUFFRixPQUFPLENBQUMsYUFBYSxDQUFDLElBQUksQ0FBQyxTQUFTLENBQUMsQ0FBQztJQUV0QyxPQUFPLENBQUMsYUFBYSxDQUFDLElBQUksQ0FBQyxNQUFNLENBQUMsTUFBTSxDQUFDLDJCQUEyQixDQUFDLENBQUMsQ0FBQyxFQUFFLEVBQUUsQ0FBQyx1QkFBdUIsQ0FBQyxDQUFDLENBQUMsQ0FBQyxDQUFDLENBQUM7SUFDekcsdUJBQXVCLENBQUMsTUFBTSxDQUFDLE1BQU0sQ0FBQyxnQkFBZ0IsQ0FBQyxDQUFBO0FBRXhELENBQUM7QUFiRCw0QkFhQztBQXlCRCxTQUFzQixrQkFBa0I7O1FBQ3ZDLElBQUk7WUFDSCxNQUFNLFVBQVUsQ0FBQyxJQUFJLEVBQUUsQ0FBQztTQUN4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