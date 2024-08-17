'use strict';
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
const TaskRunner_1 = require("./TaskRunner");
const _1 = require(".");
let taskProvider;
function activate(context) {
    return __awaiter(this, void 0, void 0, function* () {
        let spfxFncs = undefined;
        // Retrieve the gulp path and register the task provider
        const gulpPaths = yield _1.SPFxTaskProvider.gulpPath();
        // Register the SPFx task provider
        taskProvider = vscode.workspace.registerTaskProvider(_1.TASKRUNNER_TYPE, {
            provideTasks: () => {
                if (!spfxFncs) {
                    spfxFncs = _1.SPFxTaskProvider.get(gulpPaths);
                }
                return spfxFncs;
            },
            resolveTask(task) {
                return task;
            }
        });
        // List all available gulp tasks
        let taskList = vscode.commands.registerCommand('spfxTaskRunner.list', () => __awaiter(this, void 0, void 0, function* () {
            yield TaskRunner_1.TaskRunner.list();
        }));
        // Create the debug package
        let pkgDebug = vscode.commands.registerCommand('spfxTaskRunner.pkgDebug', () => __awaiter(this, void 0, void 0, function* () {
            yield TaskRunner_1.TaskRunner.packaging();
        }));
        // Create the release package
        let pkgRelease = vscode.commands.registerCommand('spfxTaskRunner.pkgRelease', () => __awaiter(this, void 0, void 0, function* () {
            yield TaskRunner_1.TaskRunner.packaging(true);
        }));
        // Start serving the local server
        let serve = vscode.commands.registerCommand('spfxTaskRunner.serve', () => __awaiter(this, void 0, void 0, function* () {
            yield TaskRunner_1.TaskRunner.serve();
        }));
        // Start serving the local server
        let pickTask = vscode.commands.registerCommand('spfxTaskRunner.pickTask', () => __awaiter(this, void 0, void 0, function* () {
            yield TaskRunner_1.TaskRunner.showOptions();
        }));
        // Fix to only show the menu actions when the extension activation conditions were met
        vscode.commands.executeCommand('setContext', 'spfxProjectCheck', true);
        // Register all actions
        context.subscriptions.push(taskProvider, taskList, pkgDebug, pkgRelease, serve, pickTask);
        // Log that the extension is active
        console.log('SPFx Task Runner is now active!');
    });
}
exports.activate = activate;
// this method is called when your extension is deactivated
function deactivate() {
    if (taskProvider) {
        taskProvider.dispose();
    }
}
exports.deactivate = deactivate;
//# sourceMappingURL=extension.js.map