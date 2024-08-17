#!/usr/bin/env node
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
exports.readPath = void 0;
const commander_1 = require("commander");
const crypto = require("crypto");
const fs = require("fs-extra");
const path = require("path");
const process = require("process");
const parser_1 = require("../../parser/parser");
const activate_1 = require("../activate");
const api_1 = require("../api");
const config_1 = require("../config");
const diagnosticStore = new Map();
let useConfig;
function getMessage(storedDiagnostic) {
    const { diagnostic, fsPath } = storedDiagnostic;
    const range = `${diagnostic.range.start.line + 1},${diagnostic.range.start.character + 1}`;
    const severity = `${api_1.DiagnosticSeverity[diagnostic.severity].substr(0, 4).toUpperCase()}`;
    return `${fsPath}(${range}) [${severity}][${diagnostic.source}][${diagnostic.ruleName}] ${diagnostic.message}`;
}
function readFile(filename) {
    return __awaiter(this, void 0, void 0, function* () {
        let errorCount = 0;
        const fsPath = path.relative(process.cwd(), filename);
        if (!api_1.ProfileComponent.isProfileComponent(fsPath)) {
            return errorCount;
        }
        const textDocument = (yield fs.readFile(fsPath)).toString();
        const parsedDocument = api_1.ProfileComponent.isPsl(fsPath) ? (0, parser_1.parseText)(textDocument) : undefined;
        const profileComponent = new api_1.ProfileComponent(fsPath, textDocument);
        const diagnostics = (0, activate_1.getDiagnostics)(profileComponent, parsedDocument, useConfig);
        diagnostics.forEach(diagnostic => {
            if (diagnostic.severity === api_1.DiagnosticSeverity.Warning || diagnostic.severity === api_1.DiagnosticSeverity.Error) {
                errorCount += 1;
            }
            const mapDiagnostics = diagnosticStore.get(diagnostic.source);
            if (!mapDiagnostics)
                diagnosticStore.set(diagnostic.source, [{ diagnostic, fsPath }]);
            else
                mapDiagnostics.push({ diagnostic, fsPath });
        });
        return errorCount;
    });
}
function readPath(fileString) {
    return __awaiter(this, void 0, void 0, function* () {
        const files = fileString.split(';').filter(x => x);
        const promises = [];
        let exitCode = 0;
        for (const filePath of files) {
            const absolutePath = path.resolve(filePath);
            if (!absolutePath)
                continue;
            const stat = yield fs.lstat(absolutePath);
            if (stat.isDirectory()) {
                const fileNames = yield fs.readdir(absolutePath);
                for (const fileName of fileNames) {
                    const absolutePathInDir = path.resolve(path.join(absolutePath, fileName));
                    yield readPath(absolutePathInDir);
                }
            }
            else if (stat.isFile()) {
                const promise = readFile(absolutePath).then(errorCount => {
                    exitCode += errorCount;
                }).catch((e) => {
                    if (e.message)
                        console.error(absolutePath, e.message, e.stack);
                    else
                        console.error(absolutePath, e);
                });
                promises.push(promise);
            }
        }
        yield Promise.all(promises);
        return exitCode;
    });
}
exports.readPath = readPath;
function processConfig() {
    return __awaiter(this, void 0, void 0, function* () {
        const configPath = path.join(process.cwd(), 'psl-lint.json');
        yield fs.lstat(configPath).then(() => __awaiter(this, void 0, void 0, function* () {
            yield (0, config_1.setConfig)(configPath);
            useConfig = true;
        })).catch(() => {
            useConfig = false;
        });
    });
}
function outputResults(reportFileName) {
    return __awaiter(this, void 0, void 0, function* () {
        if (reportFileName) {
            yield generateCodeQualityReport(reportFileName);
            console.log('Finished report.');
        }
        else {
            printOutputToConsole();
            console.log('Finished lint.');
        }
    });
}
function printOutputToConsole() {
    for (const source of diagnosticStore.keys()) {
        const diagnostics = diagnosticStore.get(source);
        const word = diagnosticStore.get(source).length === 1 ? 'diagnostic' : 'diagnostics';
        console.log(`[${source}] ${diagnostics.length} ${word}:`);
        diagnostics.forEach(diagnostic => {
            console.log(getMessage(diagnostic));
        });
    }
}
function generateCodeQualityReport(reportFileName) {
    return __awaiter(this, void 0, void 0, function* () {
        const counts = {};
        const issues = [];
        for (const ruleDiagnostics of diagnosticStore.values()) {
            for (const storedDiagnostic of ruleDiagnostics) {
                const { diagnostic, fsPath } = storedDiagnostic;
                const count = counts[diagnostic.ruleName];
                if (!count) {
                    counts[diagnostic.ruleName] = 1;
                }
                else {
                    counts[diagnostic.ruleName] = counts[diagnostic.ruleName] + 1;
                }
                if (diagnostic.ruleName === 'MemberCamelCase')
                    continue;
                const issue = {
                    check_name: diagnostic.ruleName,
                    description: `[${diagnostic.ruleName}] ${diagnostic.message.trim().replace(/\.$/, '')}`,
                    fingerprint: hashObject(diagnostic),
                    location: {
                        lines: {
                            begin: diagnostic.range.start.line + 1,
                            end: diagnostic.range.end.line + 1,
                        },
                        path: fsPath,
                    },
                };
                issues.push(issue);
            }
        }
        console.log('Diagnostics found in repository:');
        console.table(counts);
        yield fs.writeFile(reportFileName, JSON.stringify(issues));
    });
}
function hashObject(object) {
    const hash = crypto.createHash('md5')
        .update(JSON.stringify(object, (key, value) => {
        if (key[0] === '_')
            return undefined; // remove api stuff
        else if (typeof value === 'function') { // consider functions
            return value.toString();
        }
        else
            return value;
    }))
        .digest('hex');
    return hash;
}
function getCliArgs() {
    const command = new commander_1.Command('psl-lint');
    command
        .argument('<fileList>')
        .name('psl-lint')
        .usage('<fileString>')
        .option('-o, --output <output>', 'Name of output file')
        .description('fileString    a ; delimited string of file paths')
        .parse(process.argv);
    return { fileString: command.args[0], reportFileName: command.getOptionValue('output') };
}
(function main() {
    return __awaiter(this, void 0, void 0, function* () {
        if (require.main !== module) {
            return;
        }
        const { fileString, reportFileName } = getCliArgs();
        if (fileString) {
            yield processConfig();
            if (reportFileName)
                console.log('Starting report.');
            else
                console.log('Starting lint.');
            const exitCode = yield readPath(fileString);
            yield outputResults(reportFileName);
            process.exit(exitCode);
        }
        else {
            console.log('Nothing to lint.');
        }
    });
})();
// psl-lint $(git diff master...${CI_BUILD_REF_NAME} --name-only | tr "\n" ";")
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoiY2xpLmpzIiwic291cmNlUm9vdCI6IiIsInNvdXJjZXMiOlsiLi4vLi4vLi4vLi4vc3JjL3BzbExpbnQvY2xpL2NsaS50cyJdLCJuYW1lcyI6W10sIm1hcHBpbmdzIjoiOzs7Ozs7Ozs7Ozs7O0FBQ0EseUNBQW9DO0FBQ3BDLGlDQUFpQztBQUNqQywrQkFBK0I7QUFDL0IsNkJBQTZCO0FBQzdCLG1DQUFtQztBQUNuQyxnREFBZ0Q7QUFDaEQsMENBQTZDO0FBQzdDLGdDQUEwRTtBQUMxRSxzQ0FBc0M7QUF5QnRDLE1BQU0sZUFBZSxHQUFvQyxJQUFJLEdBQUcsRUFBRSxDQUFDO0FBQ25FLElBQUksU0FBa0IsQ0FBQztBQUV2QixTQUFTLFVBQVUsQ0FBQyxnQkFBa0M7SUFDckQsTUFBTSxFQUFFLFVBQVUsRUFBRSxNQUFNLEVBQUUsR0FBRyxnQkFBZ0IsQ0FBQztJQUNoRCxNQUFNLEtBQUssR0FBRyxHQUFHLFVBQVUsQ0FBQyxLQUFLLENBQUMsS0FBSyxDQUFDLElBQUksR0FBRyxDQUFDLElBQUksVUFBVSxDQUFDLEtBQUssQ0FBQyxLQUFLLENBQUMsU0FBUyxHQUFHLENBQUMsRUFBRSxDQUFDO0lBQzNGLE1BQU0sUUFBUSxHQUFHLEdBQUcsd0JBQWtCLENBQUMsVUFBVSxDQUFDLFFBQVEsQ0FBQyxDQUFDLE1BQU0sQ0FBQyxDQUFDLEVBQUUsQ0FBQyxDQUFDLENBQUMsV0FBVyxFQUFFLEVBQUUsQ0FBQztJQUN6RixPQUFPLEdBQUcsTUFBTSxJQUFJLEtBQUssTUFBTSxRQUFRLEtBQUssVUFBVSxDQUFDLE1BQU0sS0FBSyxVQUFVLENBQUMsUUFBUSxLQUFLLFVBQVUsQ0FBQyxPQUFPLEVBQUUsQ0FBQztBQUNoSCxDQUFDO0FBRUQsU0FBZSxRQUFRLENBQUMsUUFBZ0I7O1FBQ3ZDLElBQUksVUFBVSxHQUFHLENBQUMsQ0FBQztRQUNuQixNQUFNLE1BQU0sR0FBRyxJQUFJLENBQUMsUUFBUSxDQUFDLE9BQU8sQ0FBQyxHQUFHLEVBQUUsRUFBRSxRQUFRLENBQUMsQ0FB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