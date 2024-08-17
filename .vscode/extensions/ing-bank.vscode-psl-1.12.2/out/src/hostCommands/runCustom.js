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
exports.setCustomRunContext = exports.registerCustomRunContext = exports.runCoverageHandler = exports.runTestHandler = exports.coverageContext = exports.testContext = void 0;
const fs = require("fs-extra");
const path = require("path");
const vscode = require("vscode");
const environment = require("../common/environment");
const utils = require("./hostCommandUtils");
const pslUnitTest_1 = require("./pslUnitTest");
const icon = "\u25B6" /* utils.icons.RUN */;
exports.testContext = {
    command: 'runTest',
    contextKey: 'psl.runTestContext',
};
exports.coverageContext = {
    command: 'runCoverage',
    contextKey: 'psl.runCoverageContext',
};
const customRunContexts = [exports.testContext, exports.coverageContext];
function runTestHandler(context) {
    return __awaiter(this, void 0, void 0, function* () {
        handle(context, exports.testContext);
    });
}
exports.runTestHandler = runTestHandler;
function runCoverageHandler(context) {
    return __awaiter(this, void 0, void 0, function* () {
        handle(context, exports.coverageContext);
    });
}
exports.runCoverageHandler = runCoverageHandler;
function handle(context, runContext) {
    return __awaiter(this, void 0, void 0, function* () {
        const c = utils.getFullContext(context);
        if (c.mode === 1 /* utils.ContextMode.FILE */) {
            return runPSL(c.fsPath, runContext).catch(() => { });
        }
        else if (c.mode === 2 /* utils.ContextMode.DIRECTORY */) {
            const files = yield vscode.window.showOpenDialog({ defaultUri: vscode.Uri.file(c.fsPath), canSelectMany: true, openLabel: 'Run PSL' });
            if (!files)
                return;
            for (const fsPath of files.map(file => file.fsPath)) {
                yield runPSL(fsPath, runContext).catch(() => { });
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
                yield runPSL(fsPath, runContext).catch(() => { });
            }
        }
        return;
    });
}
function runPSL(fsPath, runContext) {
    return __awaiter(this, void 0, void 0, function* () {
        if (!fs.statSync(fsPath).isFile())
            return;
        const doc = yield vscode.workspace.openTextDocument(fsPath);
        const config = getFromConfiguration(doc.uri, runContext);
        if (!config)
            throw new Error(`Invalid configuration for ${runContext.command}`);
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
                const output = yield runCustom(connection, fsPath, config, env);
                connection.close();
                utils.logger.info(output.trim());
            })).catch((e) => {
                utils.logger.error(`${"\u274C" /* utils.icons.ERROR */} ${icon} error in ${env.name} ${e.message}`);
            }));
        }
        yield Promise.all(promises);
    });
}
function getFromConfiguration(uri, runContext) {
    const configs = vscode.workspace.getConfiguration('psl', uri).get('customTasks');
    const config = configs.find(c => c.command === runContext.command);
    if (!config || !config.mrpcID || !config.request) {
        return undefined;
    }
    return config;
}
function registerCustomRunContext() {
    if (vscode.window.activeTextEditor)
        setCustomRunContext(vscode.window.activeTextEditor);
    vscode.window.onDidChangeActiveTextEditor(setCustomRunContext);
}
exports.registerCustomRunContext = registerCustomRunContext;
function setCustomRunContext(textEditor) {
    for (const context of customRunContexts) {
        let showCommand = false;
        if (textEditor) {
            if (getFromConfiguration(textEditor.document.uri, context))
                showCommand = true;
        }
        vscode.commands.executeCommand('setContext', context.contextKey, showCommand);
    }
}
exports.setCustomRunContext = setCustomRunContext;
function runCustom(connection, fsPath, config, env) {
    return __awaiter(this, void 0, void 0, function* () {
        const output = yield connection.runCustom(fsPath, config.mrpcID, config.request);
        if (config.command !== exports.coverageContext.command)
            return output;
        const parsedOutput = (0, pslUnitTest_1.parseCoverageOutput)(output);
        if (parsedOutput.documents.length) {
            const items = parsedOutput.documents.map(documentCoverage => {
                return {
                    description: documentCoverage.coverage,
                    documentCoverage,
                    label: documentCoverage.name,
                };
            });
            vscode.window.showQuickPick(items, { canPickMany: true, placeHolder: 'Show coverage', ignoreFocusOut: true })
                .then(choices => {
                if (!choices || !choices.length)
                    return;
                (0, pslUnitTest_1.displayCoverage)(choices.map(x => x.documentCoverage), env, path.basename(fsPath));
            });
        }
        return parsedOutput.output;
    });
}
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoicnVuQ3VzdG9tLmpzIiwic291cmNlUm9vdCI6IiIsInNvdXJjZXMiOlsiLi4vLi4vLi4vc3JjL2hvc3RDb21tYW5kcy9ydW5DdXN0b20udHMiXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6Ijs7Ozs7Ozs7Ozs7O0FBQUEsK0JBQStCO0FBQy9CLDZCQUE2QjtBQUM3QixpQ0FBaUM7QUFDakMscURBQXFEO0FBRXJELDRDQUE0QztBQUM1QywrQ0FBc0Y7QUFFdEYsTUFBTSxJQUFJLGlDQUFrQixDQUFDO0FBT2hCLFFBQUEsV0FBVyxHQUFxQjtJQUM1QyxPQUFPLEVBQUUsU0FBUztJQUNsQixVQUFVLEVBQUUsb0JBQW9CO0NBQ2hDLENBQUM7QUFFVyxRQUFBLGVBQWUsR0FBcUI7SUFDaEQsT0FBTyxFQUFFLGFBQWE7SUFDdEIsVUFBVSxFQUFFLHdCQUF3QjtDQUNwQyxDQUFDO0FBRUYsTUFBTSxpQkFBaUIsR0FBRyxDQUFDLG1CQUFXLEVBQUUsdUJBQWUsQ0FBQyxDQUFDO0FBUXpELFNBQXNCLGNBQWMsQ0FBQyxPQUFzQzs7UUFDMUUsTUFBTSxDQUFDLE9BQU8sRUFBRSxtQkFBVyxDQUFDLENBQUM7SUFDOUIsQ0FBQztDQUFBO0FBRkQsd0NBRUM7QUFFRCxTQUFzQixrQkFBa0IsQ0FBQyxPQUFzQzs7UUFDOUUsTUFBTSxDQUFDLE9BQU8sRUFBRSx1QkFBZSxDQUFDLENBQUM7SUFDbEMsQ0FBQztDQUFBO0FBRkQsZ0RBRUM7QUFFRCxTQUFlLE1BQU0sQ0FBQyxPQUFzQyxFQUFFLFVBQTZCOztRQUMxRixNQUFNLENBQUMsR0FBRyxLQUFLLENBQUMsY0FBYyxDQUFDLE9BQU8sQ0FBQyxDQUFDO1FBQ3hDLElBQUksQ0FBQyxDQUFDLElBQUksbUNBQTJCLEVBQUU7WUFDdEMsT0FBTyxNQUFNLENBQUMsQ0FBQyxDQUFDLE1BQU0sRUFBRSxVQUFVLENBQUMsQ0FBQyxLQUFLLENBQUMsR0FBRyxFQUFFLEdBQUcsQ0FBQyxDQUFDLENBQUM7U0FDckQ7YUFDSSxJQUFJLENBQUMsQ0FB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