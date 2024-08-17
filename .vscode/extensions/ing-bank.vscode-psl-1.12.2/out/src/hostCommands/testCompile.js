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
exports.testCompile = exports.testCompileHandler = void 0;
const vscode = require("vscode");
const utils = require("./hostCommandUtils");
const path = require("path");
const fs = require("fs-extra");
const diagnostics_1 = require("../common/diagnostics");
const extension = require("../extension");
const environment = require("../common/environment");
const icon = "\u2699" /* utils.icons.TEST */;
function testCompileHandler(context) {
    return __awaiter(this, void 0, void 0, function* () {
        let c = utils.getFullContext(context);
        let diagnostics = [];
        if (c.mode === 1 /* utils.ContextMode.FILE */) {
            yield testCompile(c.fsPath).catch(() => { });
        }
        else if (c.mode === 2 /* utils.ContextMode.DIRECTORY */) {
            let files = yield vscode.window.showOpenDialog({ defaultUri: vscode.Uri.file(c.fsPath), canSelectMany: true, openLabel: 'Test Compile' });
            if (!files)
                return;
            for (let fsPath of files.map(file => file.fsPath)) {
                let result = yield testCompile(fsPath).catch(() => { });
                if (result)
                    diagnostics = diagnostics.concat(result);
            }
        }
        else {
            let quickPick = yield environment.workspaceQuickPick();
            if (!quickPick)
                return;
            let chosenEnv = quickPick;
            let files = yield vscode.window.showOpenDialog({ defaultUri: vscode.Uri.file(chosenEnv.fsPath), canSelectMany: true, openLabel: 'Test Compile' });
            if (!files)
                return;
            for (let fsPath of files.map(file => file.fsPath)) {
                let result = yield testCompile(fsPath);
                if (result)
                    diagnostics = diagnostics.concat(result);
            }
        }
    });
}
exports.testCompileHandler = testCompileHandler;
function testCompile(fsPath) {
    return __awaiter(this, void 0, void 0, function* () {
        const fileStats = yield fs.stat(fsPath);
        if (!fileStats.isFile()) {
            utils.logger.error(`${"\u274C" /* utils.icons.ERROR */} ${icon} ${fsPath} is not a file.`);
            return true;
        }
        let textDocument = yield vscode.workspace.openTextDocument(fsPath);
        if (!canTestCompileFile(textDocument, fsPath)) {
            // The error message for the specific error was already added in 'canTestCompileFile'
            return true;
        }
        let testCompileSucceeded = false;
        let envs;
        try {
            envs = yield utils.getEnvironment(fsPath);
        }
        catch (e) {
            utils.logger.error(`${"\u274C" /* utils.icons.ERROR */} ${icon} Invalid environment configuration.`);
            return true;
        }
        if (envs.length === 0) {
            utils.logger.error(`${"\u274C" /* utils.icons.ERROR */} ${icon} No environments selected.`);
            return true;
        }
        let testCompiles = [];
        for (let env of envs) {
            testCompiles.push(utils.executeWithProgress(`${icon} ${path.basename(fsPath)} TEST COMPILE`, () => __awaiter(this, void 0, void 0, function* () {
                yield textDocument.save();
                utils.logger.info(`${"\u2026" /* utils.icons.WAIT */} ${icon} ${path.basename(fsPath)} TEST COMPILE in ${env.name}`);
                let connection = yield utils.getConnection(env);
                let output = yield connection.testCompile(fsPath);
                connection.close();
                let pslDiagnostics = parseCompilerOutput(output, textDocument);
                testCompileSucceeded = pslDiagnostics.filter(d => d.severity === vscode.DiagnosticSeverity.Error).length === 0;
                let testCompileWarning = pslDiagnostics.filter(d => d.severity === vscode.DiagnosticSeverity.Warning).length > 0;
                if (!testCompileSucceeded) {
                    output = `${"\u274C" /* utils.icons.ERROR */} ${icon} ${path.basename(fsPath)} TEST COMPILE in ${env.name} failed\n` + output;
                }
                else if (testCompileWarning) {
                    output = `${"\u26A0" /* utils.icons.WARN */} ${icon} ${path.basename(fsPath)} TEST COMPILE in ${env.name} succeeded with warning\n` + output;
                }
                else {
                    output = `${"\u2714" /* utils.icons.SUCCESS */} ${icon} ${path.basename(fsPath)} TEST COMPILE in ${env.name} succeeded\n` + output;
                }
                utils.logger.info(output.split('\n').join('\n' + ' '.repeat(20)));
                diagnostics_1.PSLDiagnostic.setDiagnostics(pslDiagnostics, env.name, fsPath);
            })).catch((e) => {
                utils.logger.error(`${"\u274C" /* utils.icons.ERROR */} ${icon} error in ${env.name} ${e.message}`);
            }));
        }
        return false;
    });
}
exports.testCompile = testCompile;
function parseCompilerOutput(compilerOutput, document) {
    /*
    ZFeatureToggleUtilities.PROC compiled at 15:31 on 29-05-17
    Source: ZFeatureToggleUtilities.PROC

    %PSL-E-SYNTAX: Missing #PROPERTYDEF
    In module: ZFeatureToggleUtilities

    Source: ZFeatureToggleUtilities.PROC
        #PROPEYDEF dummy class = String private node = "dummy"
    %PSL-E-SYNTAX: Unexpected compiler command: PROPEYDEF
    At source code line: 25 in subroutine:

    Source: ZFeatureToggleUtilities.PROC

    %PSL-I-LIST: 2 errors, 0 warnings, 0 informational messages ** failed **
    In module: ZFeatureToggleUtilities
    */
    let outputArrays = splitCompilerOutput(compilerOutput);
    let pslDiagnostics = [];
    outputArrays.slice(0, outputArrays.length - 1).forEach(pslCompilerMessage => {
        let lineNumber = pslCompilerMessage.getLineNumber();
        if (lineNumber - 1 > document.lineCount || lineNumber <= 0)
            return;
        let codeLine = document.lineAt(lineNumber - 1).text;
        let startIndex = codeLine.search(/\S/); // returns the index of the first non-whitespace character
        if (startIndex === -1)
            startIndex = 0; // codeLine is only whitespace characters
        let range = new vscode.Range(lineNumber - 1, startIndex, lineNumber - 1, codeLine.length);
        let severity = pslCompilerMessage.getSeverity();
        if (severity >= 0) {
            pslDiagnostics.push(new diagnostics_1.PSLDiagnostic(`${pslCompilerMessage.message}`, severity, document.fileName, range));
        }
    });
    return pslDiagnostics;
}
function canTestCompileFile(document, fsPath) {
    let compilable = false;
    if (vscode.languages.match(extension.PSL_MODE, document)) {
        compilable = true;
    }
    else {
        let fileTypeDescription = "";
        if (vscode.languages.match(extension.BATCH_MODE, document)) {
            fileTypeDescription = "Batch";
        }
        else if (vscode.languages.match(extension.COL_MODE, document)) {
            fileTypeDescription = "Column Definition";
        }
        else if (vscode.languages.match(extension.DATA_MODE, document)) {
            fileTypeDescription = "Data File";
        }
        else if (vscode.languages.match(extension.SERIAL_MODE, document)) {
            fileTypeDescription = "Serialized Data";
        }
        else if (vscode.languages.match(extension.TBL_MODE, document)) {
            fileTypeDescription = "Table Definition";
        }
        else if (vscode.languages.match(extension.TRIG_MODE, document)) {
            fileTypeDescription = "Trigger";
        }
        if (fileTypeDescription != "") {
            utils.logger.error(`${"\u274C" /* utils.icons.ERROR */} ${icon} ${fileTypeDescription} ${path.basename(fsPath)} cannot be test compiled.`);
        }
        else {
            utils.logger.error(`${"\u274C" /* utils.icons.ERROR */} ${icon} ${path.basename(fsPath)} is not a PSL file.`);
        }
    }
    return compilable;
}
class PSLCompilerMessage {
    isFilled() {
        return (this.source && this.message && this.location) !== '';
    }
    getLineNumber() {
        if (this.location.startsWith('In module:'))
            return -1;
        return parseInt(this.location.replace('At source code line: ', '').split(' ')[0]);
    }
    getSeverity() {
        if (this.message.startsWith('%PSL-W-')) {
            return vscode.DiagnosticSeverity.Warning;
        }
        else if (this.message.startsWith('%PSL-E-')) {
            return vscode.DiagnosticSeverity.Error;
        }
        else if (this.message.startsWith('%PSL-I-')) {
            return vscode.DiagnosticSeverity.Information;
        }
        return -1;
    }
}
function splitCompilerOutput(compilerOutput) {
    /**
     * breaks apart the psl compiler output string into an arrays of compiler messages
     */
    let outputArrays = [];
    let compilerMessage;
    let splitCompilerOutput = compilerOutput.replace(/\r/g, '').trim().split('\n');
    for (let i = 1; i < splitCompilerOutput.length; i++) {
        compilerMessage = new PSLCompilerMessage();
        compilerMessage.source = splitCompilerOutput[i];
        compilerMessage.code = splitCompilerOutput[i + 1];
        compilerMessage.message = splitCompilerOutput[i + 2];
        compilerMessage.location = splitCompilerOutput[i + 3];
        if (compilerMessage.isFilled())
            outputArrays.push(compilerMessage);
        i = i + 4;
    }
    return outputArrays;
}
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoidGVzdENvbXBpbGUuanMiLCJzb3VyY2VSb290IjoiIiwic291cmNlcyI6WyIuLi8uLi8uLi9zcmMvaG9zdENvbW1hbmRzL3Rlc3RDb21waWxlLnRzIl0sIm5hbWVzIjpbXSwibWFwcGluZ3MiOiI7Ozs7Ozs7Ozs7OztBQUFBLGlDQUFpQztBQUNqQyw0Q0FBNEM7QUFDNUMsNkJBQTZCO0FBQzdCLCtCQUErQjtBQUMvQix1REFBc0Q7QUFDdEQsMENBQTBDO0FBQzFDLHFEQUFxRDtBQUVyRCxNQUFNLElBQUksa0NBQW1CLENBQUM7QUFFOUIsU0FBc0Isa0JBQWtCLENBQUMsT0FBc0M7O1FBQzlFLElBQUksQ0FBQyxHQUFHLEtBQUssQ0FBQyxjQUFjLENBQUMsT0FBTyxDQUFDLENBQUM7UUFDdEMsSUFBSSxXQUFXLEdBQUcsRUFBRSxDQUFBO1FBQ3BCLElBQUksQ0FBQyxDQUFDLElBQUksbUNBQTJCLEVBQUU7WUFDdEMsTUFBTSxXQUFXLENBQUMsQ0FBQyxDQUFDLE1BQU0sQ0FBQyxDQUFDLEtBQUssQ0FBQyxHQUFHLEVBQUUsR0FBRyxDQUFDLENBQUMsQ0FBQztTQUM3QzthQUNJLElBQUksQ0FBQyxDQUFDLElBQUksd0NBQWdDLEVBQUU7WUFDaEQsSUFBSSxLQUFLLEdBQUcsTUFBTSxNQUFNLENBQUMsTUFBTSxDQUFDLGNBQWMsQ0FBQyxFQUFFLFVBQVUsRUFBRSxNQUFNLENBQUMsR0FBRyxDQUFDLElBQUksQ0FBQyxDQUFDLENBQUMsTUFBTSxDQUFDLEVBQUUsYUFBYSxFQUFFLElBQUksRUFBRSxTQUFTLEVBQUUsY0FBYyxFQUFFLENBQUMsQ0FBQTtZQUN6SSxJQUFJLENBQUMsS0FBSztnQkFBRSxPQUFPO1lBQ25CLEtBQUssSUFBSSxNQUFNLElBQUksS0FBSyxDQUFDLEdBQUcsQ0FBQyxJQUFJLENBQUMsRUFBRSxDQUFDLElBQUksQ0FBQyxNQUFNLENBQUMsRUFBRTtnQkFDbEQsSUFBSSxNQUFNLEdBQUcsTUFBTSxXQUFXLENBQUMsTUFBTSxDQUFDLENBQUMsS0FB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