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
exports.writeFileWithSettings = exports.getCommandenvConfigQuickPick = exports.getEnvironment = exports.getConnection = exports.executeWithProgress = exports.getFullContext = exports.logger = exports.extensionToDescription = void 0;
const fs = require("fs-extra");
const vscode = require("vscode");
var utils_1 = require("../mtm/utils");
Object.defineProperty(exports, "extensionToDescription", { enumerable: true, get: function () { return utils_1.extensionToDescription; } });
const environment = require("../common/environment");
const mtm_1 = require("../mtm/mtm");
const outputChannel = vscode.window.createOutputChannel('Profile Host');
exports.logger = {
    error: (message) => {
        outputChannel.show(true);
        outputChannel.appendLine(`[ERR!][${new Date().toTimeString().split(' ')[0]}]    ${message.trim()}\n`);
    },
    info: (message, hide) => {
        if (!hide)
            outputChannel.show(true);
        outputChannel.appendLine(`[INFO][${new Date().toTimeString().split(' ')[0]}]    ${message.trim()}\n`);
    },
};
function getFullContext(context) {
    let fsPath = '';
    let mode;
    const activeTextEditor = vscode.window.activeTextEditor;
    if (context && context.dialog) {
        mode = 3 /* ContextMode.EMPTY */;
        return { fsPath, mode };
    }
    if ((!context || !context.fsPath) && activeTextEditor) {
        fsPath = activeTextEditor.document.fileName;
        mode = 1 /* ContextMode.FILE */;
        return { fsPath, mode };
    }
    else if (!context) {
        mode = 3 /* ContextMode.EMPTY */;
        return { fsPath, mode };
    }
    else {
        fsPath = context.fsPath;
        mode = fs.lstatSync(fsPath).isFile() ? 1 /* ContextMode.FILE */ : 2 /* ContextMode.DIRECTORY */;
        return { fsPath, mode };
    }
}
exports.getFullContext = getFullContext;
function executeWithProgress(message, task) {
    return __awaiter(this, void 0, void 0, function* () {
        return vscode.window.withProgress({ location: vscode.ProgressLocation.Window, title: message }, () => __awaiter(this, void 0, void 0, function* () {
            yield task();
            return;
        }));
    });
}
exports.executeWithProgress = executeWithProgress;
function getConnection(env) {
    return __awaiter(this, void 0, void 0, function* () {
        const connection = new mtm_1.MtmConnection(env.serverType, env.encoding);
        yield connection.open(env.host, env.port, env.user, env.password);
        return connection;
    });
}
exports.getConnection = getConnection;
function getEnvironment(fsPath) {
    return __awaiter(this, void 0, void 0, function* () {
        const workspaceFile = new environment.WorkspaceFile(fsPath);
        try {
            const envs = yield workspaceFile.environmentObjects;
            envs.forEach(env => {
                if (!env.host || !env.port || !env.user || !env.password) {
                    throw new Error();
                }
            });
            return envs;
        }
        catch (e) {
            const workspaceFolder = workspaceFile.workspaceFolder;
            if (workspaceFolder) {
                throw new Error(`Invalid configuration for Workspace Folder ${workspaceFolder.name}`);
            }
            throw new Error(`File ${fsPath} is not a member of the Workspace.`);
        }
    });
}
exports.getEnvironment = getEnvironment;
function getCommandenvConfigQuickPick(envs) {
    return __awaiter(this, void 0, void 0, function* () {
        const items = envs.map(env => {
            return { label: env.name, description: '', env };
        });
        if (items.length === 1)
            return items[0].env;
        const choice = yield vscode.window.showQuickPick(items, { placeHolder: 'Select environment to get from.' });
        if (!choice)
            return undefined;
        return choice.env;
    });
}
exports.getCommandenvConfigQuickPick = getCommandenvConfigQuickPick;
function writeFileWithSettings(fsPath, output) {
    const trailingNewline = vscode.workspace.getConfiguration('psl', vscode.Uri.file(fsPath)).get('trailingNewline');
    switch (trailingNewline) {
        case "always" /* NEWLINE_SETTING.ALWAYS */:
            if (!output.endsWith('\n'))
                output += detectNewline(output);
            break;
        case "never" /* NEWLINE_SETTING.NEVER */:
            output = output.replace(/(\r?\n)+$/, '');
            break;
        default:
            break;
    }
    return fs.writeFile(fsPath, output);
}
exports.writeFileWithSettings = writeFileWithSettings;
/**
 * https://github.com/sindresorhus/detect-newline
 */
function detectNewline(output) {
    const newlines = (output.match(/(?:\r?\n)/g) || []);
    if (newlines.length === 0) {
        return '\n';
    }
    const crlfCount = newlines.filter(el => {
        return el === '\r\n';
    }).length;
    const lfCount = newlines.length - crlfCount;
    return crlfCount > lfCount ? '\r\n' : '\n';
}
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoiaG9zdENvbW1hbmRVdGlscy5qcyIsInNvdXJjZVJvb3QiOiIiLCJzb3VyY2VzIjpbIi4uLy4uLy4uL3NyYy9ob3N0Q29tbWFuZHMvaG9zdENvbW1hbmRVdGlscy50cyJdLCJuYW1lcyI6W10sIm1hcHBpbmdzIjoiOzs7Ozs7Ozs7Ozs7QUFBQSwrQkFBK0I7QUFDL0IsaUNBQWlDO0FBQ2pDLHNDQUFzRDtBQUE3QywrR0FBQSxzQkFBc0IsT0FBQTtBQUMvQixxREFBcUQ7QUFDckQsb0NBQTJDO0FBRTNDLE1BQU0sYUFBYSxHQUFHLE1BQU0sQ0FBQyxNQUFNLENBQUMsbUJBQW1CLENBQUMsY0FBYyxDQUFDLENBQUM7QUFFM0QsUUFBQSxNQUFNLEdBQUc7SUFDckIsS0FBSyxFQUFFLENBQUMsT0FBZSxFQUFFLEVBQUU7UUFDMUIsYUFBYSxDQUFDLElBQUksQ0FBQyxJQUFJLENBQUMsQ0FBQztRQUN6QixhQUFhLENBQUMsVUFBVSxDQUFDLFVBQVUsSUFBSSxJQUFJLEVBQUUsQ0FBQyxZQUFZLEVBQUUsQ0FBQyxLQUFLLENBQUMsR0FBRyxDQUFDLENBQUMsQ0FBQyxDQUFDLFFBQVEsT0FBTyxDQUFDLElBQUksRUFBRSxJQUFJLENBQUMsQ0FBQztJQUN2RyxDQUFDO0lBQ0QsSUFBSSxFQUFFLENBQUMsT0FBZSxFQUFFLElBQWMsRUFBRSxFQUFFO1FBQ3pDLElBQUksQ0FBQyxJQUFJO1lBQUUsYUFBYSxDQUFDLElBQUksQ0FBQyxJQUFJLENBQUMsQ0FBQztRQUNwQyxhQUFhLENBQUMsVUFBVSxDQUFDLFVBQVUsSUFBSSxJQUFJLEVBQUUsQ0FBQyxZQUFZLEVBQUUsQ0FBQyxLQUFLLENBQUMsR0FBRyxDQUFDLENBQUMsQ0FBQyxDQUFDLFFBQVEsT0FBTyxDQUFDLElBQUksRUFBRSxJQUFJLENBQUMsQ0FBQztJQUN2RyxDQUFDO0NBRUQsQ0FBQztBQW9DRixTQUFnQixjQUFjLENBQUMsT0FBNEM7SUFDMUUsSUFBSSxNQUFNLEdBQVcsRUFBRSxDQUFDO0lBQ3hCLElBQUksSUFB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