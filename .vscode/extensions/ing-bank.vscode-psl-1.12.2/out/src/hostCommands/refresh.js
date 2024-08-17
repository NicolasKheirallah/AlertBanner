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
exports.refreshTableHandler = exports.refreshElementHandler = void 0;
const vscode = require("vscode");
const utils = require("./hostCommandUtils");
const path = require("path");
const fs = require("fs-extra");
const environment = require("../common/environment");
const icon = "\uD83D\uDD03" /* utils.icons.REFRESH */;
function refreshElementHandler(context) {
    return __awaiter(this, void 0, void 0, function* () {
        let c = utils.getFullContext(context);
        if (c.mode === 1 /* utils.ContextMode.FILE */) {
            return refreshElement(c.fsPath).catch(() => { });
        }
        else if (c.mode === 2 /* utils.ContextMode.DIRECTORY */) {
            let files = yield vscode.window.showOpenDialog({ defaultUri: vscode.Uri.file(c.fsPath), canSelectMany: true, openLabel: 'Refresh' });
            if (!files)
                return;
            for (let fsPath of files.map(file => file.fsPath)) {
                yield refreshElement(fsPath).catch(() => { });
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
                yield refreshElement(fsPath).catch(() => { });
            }
        }
        return;
    });
}
exports.refreshElementHandler = refreshElementHandler;
function refreshElement(fsPath) {
    return __awaiter(this, void 0, void 0, function* () {
        if (!fs.statSync(fsPath).isFile())
            return;
        let env;
        return utils.executeWithProgress(`${icon} ${path.basename(fsPath)} REFRESH`, () => __awaiter(this, void 0, void 0, function* () {
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
            utils.logger.info(`${"\u2026" /* utils.icons.WAIT */} ${icon} ${path.basename(fsPath)} REFRESH from ${env.name}`);
            let doc = yield vscode.workspace.openTextDocument(fsPath);
            yield doc.save();
            let connection = yield utils.getConnection(env);
            let output = yield connection.get(fsPath);
            yield utils.writeFileWithSettings(fsPath, output);
            utils.logger.info(`${"\u2714" /* utils.icons.SUCCESS */} ${icon} ${path.basename(fsPath)} REFRESH from ${env.name} succeeded`);
            connection.close();
            yield vscode.window.showTextDocument(doc);
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
function refreshTableHandler(context) {
    return __awaiter(this, void 0, void 0, function* () {
        let c = utils.getFullContext(context);
        if (c.mode === 1 /* utils.ContextMode.FILE */) {
            let tableName;
            if (path.extname(c.fsPath) === '.TBL') {
                tableName = path.basename(c.fsPath).split('.TBL')[0];
            }
            else if (path.extname(c.fsPath) === '.COL') {
                tableName = path.basename(c.fsPath).split('.COL')[0].split('-')[0];
            }
            else {
                return;
            }
            let targetDir = path.dirname(c.fsPath);
            return refreshTable(tableName, targetDir).catch(() => { });
        }
    });
}
exports.refreshTableHandler = refreshTableHandler;
function refreshTable(tableName, targetDirectory) {
    return __awaiter(this, void 0, void 0, function* () {
        let env;
        yield utils.executeWithProgress(`${icon} ${tableName} TABLE REFRESH`, () => __awaiter(this, void 0, void 0, function* () {
            let envs = yield utils.getEnvironment(targetDirectory);
            let choice = yield utils.getCommandenvConfigQuickPick(envs);
            if (!choice)
                return;
            env = choice;
            utils.logger.info(`${"\u2026" /* utils.icons.WAIT */} ${icon} ${tableName} TABLE REFRESH from ${env.name}`);
            let connection = yield utils.getConnection(env);
            let output = yield connection.getTable(tableName.toUpperCase() + '.TBL');
            let tableFiles = (yield fs.readdir(targetDirectory)).filter(f => f.startsWith(tableName));
            for (let file of tableFiles) {
                yield fs.remove(file);
            }
            const promises = output.split(String.fromCharCode(0)).map(content => {
                const contentArray = content.split(String.fromCharCode(1));
                const fileName = contentArray[0];
                const fileContent = contentArray[1];
                return utils.writeFileWithSettings(path.join(targetDirectory, fileName), fileContent);
            });
            yield Promise.all(promises);
            utils.logger.info(`${"\u2714" /* utils.icons.SUCCESS */} ${icon} ${tableName} TABLE REFRESH from ${env.name} succeeded`);
            connection.close();
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
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoicmVmcmVzaC5qcyIsInNvdXJjZVJvb3QiOiIiLCJzb3VyY2VzIjpbIi4uLy4uLy4uL3NyYy9ob3N0Q29tbWFuZHMvcmVmcmVzaC50cyJdLCJuYW1lcyI6W10sIm1hcHBpbmdzIjoiOzs7Ozs7Ozs7Ozs7QUFBQSxpQ0FBaUM7QUFDakMsNENBQTRDO0FBQzVDLDZCQUE2QjtBQUM3QiwrQkFBK0I7QUFDL0IscURBQXFEO0FBRXJELE1BQU0sSUFBSSwyQ0FBc0IsQ0FBQztBQUVqQyxTQUFzQixxQkFBcUIsQ0FBQyxPQUFzQzs7UUFDakYsSUFBSSxDQUFDLEdBQUcsS0FBSyxDQUFDLGNBQWMsQ0FBQyxPQUFPLENBQUMsQ0FBQztRQUN0QyxJQUFJLENBQUMsQ0FBQyxJQUFJLG1DQUEyQixFQUFFO1lBQ3RDLE9BQU8sY0FBYyxDQUFDLENBQUMsQ0FBQyxNQUFNLENBQUMsQ0FBQyxLQUFLLENBQUMsR0FBRyxFQUFFLEdBQUUsQ0FBQyxDQUFDLENBQUM7U0FDaEQ7YUFDSSxJQUFJLENBQUMsQ0FBQyxJQUFJLHdDQUFnQyxFQUFFO1lBQ2hELElBQUksS0FBSyxHQUFHLE1BQU0sTUFBTSxDQUFDLE1BQU0sQ0FBQyxjQUFjLENBQUMsRUFBQyxVQUFVLEVBQUUsTUFBTSxDQUFDLEdBQUcsQ0FBQyxJQUFJLENBQUMsQ0FBQyxDQUFDLE1BQU0sQ0FBQyxFQUFFLGFBQWEsRUFBRSxJQUFJLEVBQUUsU0FBUyxFQUFFLFNBQVMsRUFBQyxDQUFDLENBQUE7WUFDbEksSUFBSSxDQUFDLEtBQUs7Z0JBQUUsT0FBTztZQUNuQixLQUFLLElBQUksTUFBTSxJQUFJLEtBQUssQ0FBQyxHQUFHLENBQUMsSUFBSSxDQUFDLEVBQUUsQ0FBQyxJQUFJLENBQUMsTUFBTSxDQUFDLEVBQUU7Z0JBQ2xELE1BQU0sY0FBYyxDQUFDLE1BQU0sQ0FBQyxDQUFDLEtBQUssQ0FBQyxHQUFHLEVBQUUsR0FBRSxDQUFDLENBQUMsQ0FBQzthQUM3QztTQUNEO2FBQ0k7WUFDSixJQUFJLFNBQVMsR0FBRyxNQUFNLFdB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