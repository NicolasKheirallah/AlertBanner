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
exports.runPSLHandler = void 0;
const fs = require("fs-extra");
const path = require("path");
const vscode = require("vscode");
const environment = require("../common/environment");
const utils = require("./hostCommandUtils");
const icon = "\u25B6" /* utils.icons.RUN */;
function runPSLHandler(context) {
    return __awaiter(this, void 0, void 0, function* () {
        handle(context);
    });
}
exports.runPSLHandler = runPSLHandler;
function handle(context) {
    return __awaiter(this, void 0, void 0, function* () {
        const c = utils.getFullContext(context);
        if (c.mode === 1 /* utils.ContextMode.FILE */) {
            return runPSL(c.fsPath).catch(() => { });
        }
        else if (c.mode === 2 /* utils.ContextMode.DIRECTORY */) {
            const files = yield vscode.window.showOpenDialog({ defaultUri: vscode.Uri.file(c.fsPath), canSelectMany: true, openLabel: 'Run PSL' });
            if (!files)
                return;
            for (const fsPath of files.map(file => file.fsPath)) {
                yield runPSL(fsPath).catch(() => { });
            }
        }
        else {
            const quickPick = yield environment.workspaceQuickPick();
            if (!quickPick)
                return;
            const chosenEnv = quickPick;
            const files = yield vscode.window.showOpenDialog({ defaultUri: vscode.Uri.file(chosenEnv.fsPath), canSelectMany: true, openLabel: 'Run PSL' });
            if (!files)
                return;
            for (const fsPath of files.map(file => file.fsPath)) {
                yield runPSL(fsPath).catch(() => { });
            }
        }
        return;
    });
}
function runPSL(fsPath) {
    return __awaiter(this, void 0, void 0, function* () {
        if (!fs.statSync(fsPath).isFile())
            return;
        const doc = yield vscode.workspace.openTextDocument(fsPath);
        yield doc.save();
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
        const promises = [];
        for (const env of envs) {
            promises.push(utils.executeWithProgress(`${icon} ${path.basename(fsPath)} RUN`, () => __awaiter(this, void 0, void 0, function* () {
                utils.logger.info(`${"\u2026" /* utils.icons.WAIT */} ${icon} ${path.basename(fsPath)} RUN in ${env.name}`);
                const connection = yield utils.getConnection(env);
                const output = yield connection.runPsl(fsPath);
                connection.close();
                utils.logger.info(output.trim());
            })).catch((e) => {
                utils.logger.error(`${"\u274C" /* utils.icons.ERROR */} ${icon} error in ${env.name} ${e.message}`);
            }));
        }
        yield Promise.all(promises);
    });
}
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoicnVuLmpzIiwic291cmNlUm9vdCI6IiIsInNvdXJjZXMiOlsiLi4vLi4vLi4vc3JjL2hvc3RDb21tYW5kcy9ydW4udHMiXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6Ijs7Ozs7Ozs7Ozs7O0FBQUEsK0JBQStCO0FBQy9CLDZCQUE2QjtBQUM3QixpQ0FBaUM7QUFDakMscURBQXFEO0FBQ3JELDRDQUE0QztBQUU1QyxNQUFNLElBQUksaUNBQWtCLENBQUM7QUFFN0IsU0FBc0IsYUFBYSxDQUFDLE9BQXNDOztRQUN6RSxNQUFNLENBQUMsT0FBTyxDQUFDLENBQUM7SUFDakIsQ0FBQztDQUFBO0FBRkQsc0NBRUM7QUFFRCxTQUFlLE1BQU0sQ0FBQyxPQUFzQzs7UUFDM0QsTUFBTSxDQUFDLEdBQUcsS0FBSyxDQUFDLGNBQWMsQ0FBQyxPQUFPLENBQUMsQ0FBQztRQUN4QyxJQUFJLENBQUMsQ0FBQyxJQUFJLG1DQUEyQixFQUFFO1lBQ3RDLE9BQU8sTUFBTSxDQUFDLENBQUMsQ0FBQyxNQUFNLENBQUMsQ0FBQyxLQUFLLENBQUMsR0FBRyxFQUFFLEdBQUcsQ0FBQyxDQUFDLENBQUM7U0FDekM7YUFDSSxJQUFJLENBQUMsQ0FBQyxJQUFJLHdDQUFnQyxFQUFFO1lBQ2hELE1BQU0sS0FBSyxHQUFHLE1BQU0sTUFBTSxDQUFDLE1BQU0sQ0FBQyxjQUFjLENBQUMsRUFBRSxVQUFVLEVBQUUsTUFBTSxDQUFDLEdBQUcsQ0FBQyxJQUFJLENBQUMsQ0FBQyxDQUFDLE1BQU0sQ0FBQyxFQUFFLGFBQWEsRUFBRSxJQUFJLEVBQUUsU0FBUyxFQUFFLFNBQVMsRUFBRSxDQUFDLENBQUM7WUFDdkksSUFBSSxDQUFDLEtBQUs7Z0JBQUUsT0FBTztZQUNuQixLQUFLLE1BQU0sTUFBTSxJQUFJLEtBQUssQ0FBQyxHQUFHLENBQUMsSUFBSSxDQUFDLEVBQUUsQ0FBQyxJQUFJLENBQUMsTUFBTSxDQUFDLEVBQUU7Z0JBQ3BELE1BQU0sTUFBTSxDQUFDLE1BQU0sQ0FBQyxDQUFDLEtBQUssQ0FB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