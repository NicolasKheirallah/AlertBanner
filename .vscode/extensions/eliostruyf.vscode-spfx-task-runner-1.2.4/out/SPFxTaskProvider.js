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
exports.TASKRUNNER_TYPE = "SPFx";
exports.TASKRUNNER_DEBUG = "debug";
exports.TASKRUNNER_RELEASE = "release";
class SPFxTaskProvider {
    /**
     * Retrieves the gulp path
     */
    static gulpPath() {
        return __awaiter(this, void 0, void 0, function* () {
            // Retrieve the gulp command from the existing commands
            const tasks = yield vscode.tasks.fetchTasks({ type: "gulp" });
            if (tasks && tasks.length > 0) {
                const firstTask = tasks[0];
                if (firstTask && firstTask.execution && firstTask.execution["commandLine"]) {
                    // Get the gulp path from the task command line property
                    let cmdLine = firstTask.execution["commandLine"];
                    let folderPath = "";
                    // Check if extension is loaded in a workspace
                    if (firstTask.execution.options && firstTask.execution.options["cwd"]) {
                        folderPath = firstTask.execution.options["cwd"];
                    }
                    // Return the gulp execution path
                    if (cmdLine) {
                        return {
                            executePath: cmdLine.split(" ")[0],
                            folderPath
                        };
                    }
                }
            }
            return null;
        });
    }
    /**
     * Returns the task provider registration
     */
    static get(gulpCmd) {
        let gulpCommand = gulpCmd && gulpCmd.executePath ? gulpCmd.executePath : "gulp";
        let rootFolder = gulpCmd && gulpCmd.folderPath ? gulpCmd.folderPath : vscode.workspace.rootPath;
        // Escape the command path on Windows
        if (process && process.platform && process.platform.toLowerCase() === "win32") {
            gulpCommand = gulpCommand && gulpCommand.includes("(") ? gulpCommand.replace(/\(/g, "`(") : gulpCommand;
            gulpCommand = gulpCommand && gulpCommand.includes(")") ? gulpCommand.replace(/\)/g, "`)") : gulpCommand;
            rootFolder = rootFolder && rootFolder.includes("(") ? rootFolder.replace(/\(/g, "`(") : rootFolder;
            rootFolder = rootFolder && rootFolder.includes(")") ? rootFolder.replace(/\)/g, "`)") : rootFolder;
        }
        // Retrieve all the workspace folders, and match based on the retrieved command path folder
        const folders = vscode.workspace.workspaceFolders;
        let taskScope = vscode.TaskScope.Workspace;
        if (folders && folders.length > 0) {
            const crntFolder = folders.find(f => f.uri.path === rootFolder);
            if (crntFolder) {
                taskScope = crntFolder;
            }
        }
        // Register the tasks
        return [
            new vscode.Task({
                type: exports.TASKRUNNER_TYPE,
                task: `clean`
            }, taskScope, `clean`, exports.TASKRUNNER_TYPE, new vscode.ShellExecution(`${gulpCommand} clean`)),
            new vscode.Task({
                type: exports.TASKRUNNER_TYPE,
                task: `${exports.TASKRUNNER_DEBUG} bundle`
            }, taskScope, `${exports.TASKRUNNER_DEBUG} bundle`, exports.TASKRUNNER_TYPE, new vscode.ShellExecution(`${gulpCommand} bundle`)),
            new vscode.Task({
                type: exports.TASKRUNNER_TYPE,
                task: `${exports.TASKRUNNER_DEBUG} packaging`
            }, taskScope, `${exports.TASKRUNNER_DEBUG} packaging`, exports.TASKRUNNER_TYPE, new vscode.ShellExecution(`${gulpCommand} package-solution`)),
            new vscode.Task({
                type: exports.TASKRUNNER_TYPE,
                task: `${exports.TASKRUNNER_RELEASE} bundle`
            }, taskScope, `${exports.TASKRUNNER_RELEASE} bundle`, exports.TASKRUNNER_TYPE, new vscode.ShellExecution(`${gulpCommand} bundle --ship`)),
            new vscode.Task({
                type: exports.TASKRUNNER_TYPE,
                task: `${exports.TASKRUNNER_RELEASE} packaging`
            }, taskScope, `${exports.TASKRUNNER_RELEASE} packaging`, exports.TASKRUNNER_TYPE, new vscode.ShellExecution(`${gulpCommand} package-solution --ship`)),
            new vscode.Task({
                type: exports.TASKRUNNER_TYPE,
                task: "serve"
            }, taskScope, "serve", exports.TASKRUNNER_TYPE, new vscode.ShellExecution(`${gulpCommand} serve --nobrowser`))
        ];
    }
}
exports.SPFxTaskProvider = SPFxTaskProvider;
//# sourceMappingURL=SPFxTaskProvider.js.map