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
exports.parseCoverageOutput = exports.displayCoverage = void 0;
const vscode = require("vscode");
const mumps_1 = require("../language/mumps");
const hostCommandUtils_1 = require("./hostCommandUtils");
var CoverageIndicator;
(function (CoverageIndicator) {
    CoverageIndicator[CoverageIndicator["NOT_COVERED"] = 0] = "NOT_COVERED";
    CoverageIndicator[CoverageIndicator["COVERED"] = 1] = "COVERED";
    CoverageIndicator[CoverageIndicator["COMMENT"] = 2] = "COMMENT";
})(CoverageIndicator || (CoverageIndicator = {}));
const diagnosticCollection = vscode.languages.createDiagnosticCollection('psl-test');
const coverageScheme = mumps_1.MumpsVirtualDocument.schemes.coverage;
function createDecoration(backgroundKey, rulerKey) {
    return vscode.window.createTextEditorDecorationType({
        backgroundColor: new vscode.ThemeColor(backgroundKey),
        isWholeLine: true,
        overviewRulerColor: new vscode.ThemeColor(rulerKey),
        overviewRulerLane: vscode.OverviewRulerLane.Full,
        rangeBehavior: vscode.DecorationRangeBehavior.ClosedOpen,
    });
}
const notCovered = createDecoration('diffEditor.removedTextBackground', 'editorOverviewRuler.errorForeground');
const covered = createDecoration('diffEditor.insertedTextBackground', 'diffEditor.insertedTextBackground');
(0, mumps_1.onDidDeleteVirtualMumps)(uri => {
    if (uri.scheme === coverageScheme) {
        diagnosticCollection.delete(uri);
    }
});
vscode.window.onDidChangeActiveTextEditor(textEditor => {
    if (textEditor && textEditor.document.uri.scheme === coverageScheme) {
        setCoverageDecorations(textEditor);
    }
});
function displayCoverage(documents, env, testName) {
    return __awaiter(this, void 0, void 0, function* () {
        const baseUri = vscode.Uri.parse(`${coverageScheme}:`);
        const connection = yield (0, hostCommandUtils_1.getConnection)(env);
        for (const documentCoverage of documents) {
            yield connection.get(`${documentCoverage.name}.m`).then(mCode => {
                const sourceCode = mCode;
                const uri = baseUri.with({
                    path: `/${env.name}/${testName}/${documentCoverage.name}.m`,
                    query: JSON.stringify(documentCoverage),
                });
                const virtualMumps = new mumps_1.MumpsVirtualDocument(documentCoverage.name, sourceCode, uri);
                setCoverageDiagnostics(virtualMumps);
                vscode.window.showTextDocument(virtualMumps.uri, { preview: false });
            });
        }
    });
}
exports.displayCoverage = displayCoverage;
function getRoutineCoverage(uri) {
    return JSON.parse(uri.query);
}
function setCoverageDiagnostics(virtualMumps) {
    let allDiagnostics = [];
    getRoutineCoverage(virtualMumps.uri).methods.forEach(coverageMethod => {
        const documentMethod = virtualMumps.parsedDocument.methods.find(method => {
            return method.id.value === coverageMethod.name;
        });
        if (!documentMethod)
            return;
        const methodRanges = collectMethodRanges(coverageMethod);
        const diagnostics = methodRanges.map(methodRange => {
            const vscodeRange = new vscode.Range(documentMethod.line + methodRange.start, 0, documentMethod.line + methodRange.end, Number.MAX_VALUE);
            return new vscode.Diagnostic(vscodeRange, `Missing coverage in method "${coverageMethod.name}"`, vscode.DiagnosticSeverity.Error);
        });
        allDiagnostics = [...allDiagnostics, ...diagnostics];
    });
    diagnosticCollection.set(virtualMumps.uri, allDiagnostics);
}
function collectMethodRanges(methodCoverage) {
    const ranges = [];
    let previousIndicator;
    const last = methodCoverage.coverageSequence.reduce((range, lineCoverage, index) => {
        let indicator;
        if (indicator === CoverageIndicator.COMMENT)
            indicator = previousIndicator;
        else
            indicator = lineCoverage.indicator;
        if (indicator === CoverageIndicator.NOT_COVERED) {
            if (!range) {
                previousIndicator = indicator;
                return { start: index, end: index };
            }
            else {
                previousIndicator = indicator;
                range.end = index;
                return range;
            }
        }
        if (indicator === CoverageIndicator.COVERED && range) {
            previousIndicator = indicator;
            ranges.push(range);
        }
    }, undefined);
    if (last)
        ranges.push(last);
    return ranges;
}
/**
 * Called every time the document becomes active (`onDidChangeActiveTextEditor`)
 * for the mumps coverage  uri scheme.
 */
function setCoverageDecorations(textEditor) {
    const notCoveredLines = [];
    const coveredLines = [];
    const virtualMumps = (0, mumps_1.getVirtualDocument)(textEditor.document.uri);
    getRoutineCoverage(virtualMumps.uri).methods.forEach(coverageMethod => {
        const documentMethod = virtualMumps.parsedDocument.methods.find(method => {
            return method.id.value === coverageMethod.name;
        });
        if (!documentMethod)
            return;
        let lastIndicator;
        for (let lineNumber = 0; lineNumber < coverageMethod.coverageSequence.length; lineNumber++) {
            const indicator = coverageMethod.coverageSequence[lineNumber].indicator;
            if (!indicator || (indicator === CoverageIndicator.COMMENT && !lastIndicator)) {
                notCoveredLines.push(documentMethod.line + lineNumber);
                lastIndicator = 0;
            }
            else {
                coveredLines.push(documentMethod.line + lineNumber);
                lastIndicator = 1;
            }
        }
    });
    textEditor.setDecorations(notCovered, notCoveredLines.map(line => new vscode.Range(line, 0, line, Number.MAX_VALUE)));
    textEditor.setDecorations(covered, coveredLines.map(line => new vscode.Range(line, 0, line, Number.MAX_VALUE)));
}
/**
 * Parses the RPC output of a coverage run. Returns sanitized output and parsed coverage report.
 */
function parseCoverageOutput(input) {
    const parsed = {
        documents: [],
        output: input,
    };
    const begin = '#BeginCoverageInfo';
    const end = '#EndCoverageInfo';
    if (!input.includes(begin) && !input.includes(end)) {
        return parsed;
    }
    const split1 = input.split(begin);
    const split2 = split1[1].split(end);
    const output = split1[0] + split2[split2.length - 1];
    parsed.output = output;
    const routinesToPercentages = new Map();
    const match = output.match(/\d+\.\d+% - \w+/g);
    if (!match)
        return parsed;
    match.forEach(l => routinesToPercentages.set((l.split(' - ')[1]), l.split(' - ')[0]));
    parsed.documents = extractDocumentCoverage(split2[0], routinesToPercentages);
    return parsed;
}
exports.parseCoverageOutput = parseCoverageOutput;
function extractDocumentCoverage(codeOutput, routinesToPercentages) {
    const splitOutput = codeOutput.split(/\r?\n/).filter(x => x).map(x => x.trim());
    const documents = [];
    let documentCoverage = { coverage: '', methods: [], name: '' };
    const initialize = (routineName) => {
        documentCoverage = { name: routineName, methods: [], coverage: routinesToPercentages.get(routineName) || '' };
        documents.push(documentCoverage);
    };
    for (const line of splitOutput) {
        if (line.match(/^9\|.*/)) {
            initialize(line.split('|')[1]);
        }
        else if (line.match(/^1/)) {
            documentCoverage.methods.push({ name: line.split('|')[1], coverageSequence: line.split('|')[2].split('').map(s => ({ indicator: Number(s) })) });
        }
    }
    return documents;
}
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoicHNsVW5pdFRlc3QuanMiLCJzb3VyY2VSb290IjoiIiwic291cmNlcyI6WyIuLi8uLi8uLi9zcmMvaG9zdENvbW1hbmRzL3BzbFVuaXRUZXN0LnRzIl0sIm5hbWVzIjpbXSwibWFwcGluZ3MiOiI7Ozs7Ozs7Ozs7OztBQUFBLGlDQUFpQztBQUVqQyw2Q0FBc0c7QUFDdEcseURBQW1EO0FBaUJuRCxJQUFLLGlCQUlKO0FBSkQsV0FBSyxpQkFBaUI7SUFDckIsdUVBQWUsQ0FBQTtJQUNmLCtEQUFXLENBQUE7SUFDWCwrREFBVyxDQUFBO0FBQ1osQ0FBQyxFQUpJLGlCQUFpQixLQUFqQixpQkFBaUIsUUFJckI7QUFjRCxNQUFNLG9CQUFvQixHQUFHLE1BQU0sQ0FBQyxTQUFTLENBQUMsMEJBQTBCLENBQUMsVUFBVSxDQUFDLENBQUM7QUFFckYsTUFBTSxjQUFjLEdBQUcsNEJBQW9CLENBQUMsT0FBTyxDQUFDLFFBQVEsQ0FBQztBQUU3RCxTQUFTLGdCQUFnQixDQUFDLGFBQXFCLEVBQUUsUUFBZ0I7SUFDaEUsT0FBTyxNQUFNLENBQUMsTUFBTSxDQUFDLDhCQUE4QixDQUFDO1FBQ25ELGVBQWUsRUFBRSxJQUFJLE1BQU0sQ0FBQyxVQUFVLENBQUMsYUFBYSxDQUFDO1FBQ3JELFdBQVcsRUFBRSxJQUFJO1FBQ2pCLGtCQUFrQixFQUFFLElBQUksTUFBTSxDQUFDLFVBQVUsQ0FBQyxRQUFRLENBQUM7UUFDbkQsaUJBQWlCLEVBQUUsTUFBTSxDQUFDLGlCQUFpQixDQUFDLElBQUk7UUFDaEQsYUFBYSxFQUFFLE1BQU0sQ0FBQyx1QkFBdUIsQ0FBQyxVQUFVO0tBQ3hELENBQUMsQ0FBQztBQUNKLENBQUM7QUFDRCxNQUFNLFVBQVUsR0FBRyxnQkFBZ0IsQ0FBQyxrQ0FBa0MsRUFBRSxxQ0FBcUMsQ0FBQyxDQUFDO0FBQy9HLE1BQU0sT0FBTyxHQUFHLGdCQUFnQixDQUFDLG1DQUFtQyxFQUFFLG1DQUFtQyxDQUFDLENBQUM7QUFFM0csSUFBQSwrQkFBdUIsRUFBQyxHQUFH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