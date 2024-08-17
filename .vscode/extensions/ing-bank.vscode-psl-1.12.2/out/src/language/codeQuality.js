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
exports.MemberDiagnostic = exports.activate = void 0;
const vscode = require("vscode");
const extension_1 = require("../extension");
const parser = require("../parser/parser");
const activate_1 = require("../pslLint/activate");
const api = require("../pslLint/api");
const config_1 = require("../pslLint/config");
const codeAction_1 = require("./codeAction");
function activate(context) {
    return __awaiter(this, void 0, void 0, function* () {
        yield pslLintConfigurationWatchers(context);
        const lintDiagnostics = vscode.languages.createDiagnosticCollection('psl-lint');
        context.subscriptions.push(lintDiagnostics);
        // initial token
        let tokenSource = new vscode.CancellationTokenSource();
        if (vscode.window.activeTextEditor) {
            prepareRules(vscode.window.activeTextEditor.document, lintDiagnostics, tokenSource.token);
        }
        vscode.window.onDidChangeActiveTextEditor(e => {
            if (!e)
                return;
            prepareRules(e.document, lintDiagnostics, tokenSource.token);
        });
        vscode.workspace.onDidChangeTextDocument(e => {
            if (!e)
                return;
            tokenSource.cancel();
            tokenSource = new vscode.CancellationTokenSource();
            prepareRules(e.document, lintDiagnostics, tokenSource.token);
        });
        vscode.workspace.onDidCloseTextDocument(closedDocument => {
            lintDiagnostics.delete(closedDocument.uri);
        });
        const actionProvider = new codeAction_1.PSLActionProvider();
        for (const mode of [extension_1.PSL_MODE, extension_1.BATCH_MODE, extension_1.TRIG_MODE]) {
            context.subscriptions.push(vscode.languages.registerCodeActionsProvider(mode, actionProvider));
        }
    });
}
exports.activate = activate;
function pslLintConfigurationWatchers(context) {
    return __awaiter(this, void 0, void 0, function* () {
        return Promise.all(vscode.workspace.workspaceFolders
            .map(workspace => new vscode.RelativePattern(workspace, 'psl-lint.json'))
            .map((pattern) => __awaiter(this, void 0, void 0, function* () {
            const watcher = vscode.workspace.createFileSystemWatcher(pattern);
            context.subscriptions.push(watcher.onDidChange(uri => {
                (0, config_1.setConfig)(uri.fsPath);
            }), watcher.onDidCreate(uri => {
                (0, config_1.setConfig)(uri.fsPath);
            }));
            watcher.onDidDelete(uri => {
                (0, config_1.removeConfig)(uri.fsPath);
            });
            const uris = yield vscode.workspace.findFiles(pattern);
            if (!uris.length)
                return;
            yield (0, config_1.setConfig)(uris[0].fsPath);
        })));
    });
}
class MemberDiagnostic extends vscode.Diagnostic {
}
exports.MemberDiagnostic = MemberDiagnostic;
function prepareRules(textDocument, lintDiagnostics, cancellationToken) {
    if (!api.ProfileComponent.isProfileComponent(textDocument.fileName))
        return;
    const lintConfigValue = vscode.workspace.getConfiguration('psl', textDocument.uri).get('lint');
    let useConfig = false;
    if (lintConfigValue === 'config') {
        // check if config exist first
        const config = (0, config_1.getConfig)(textDocument.uri.fsPath);
        if (!config)
            return;
        useConfig = true;
    }
    else if (lintConfigValue !== 'all' && lintConfigValue !== true) {
        lintDiagnostics.clear();
        return;
    }
    process.nextTick(() => {
        if (!cancellationToken.isCancellationRequested) {
            lint(textDocument, useConfig, cancellationToken, lintDiagnostics);
        }
    });
}
function lint(textDocument, useConfig, cancellationToken, lintDiagnostics) {
    const profileComponent = prepareDocument(textDocument);
    const parsedDocument = api.ProfileComponent.isPsl(profileComponent.fsPath) ?
        parser.parseText(textDocument.getText()) : undefined;
    const diagnostics = (0, activate_1.getDiagnostics)(profileComponent, parsedDocument, useConfig);
    const memberDiagnostics = transform(diagnostics, textDocument.uri);
    process.nextTick(() => {
        if (!cancellationToken.isCancellationRequested) {
            lintDiagnostics.set(textDocument.uri, memberDiagnostics);
        }
    });
}
function prepareDocument(textDocument) {
    const getTextAtLine = (n) => textDocument.lineAt(n).text;
    const profileComponent = new api.ProfileComponent(textDocument.uri.fsPath, textDocument.getText(), getTextAtLine);
    return profileComponent;
}
function transform(diagnostics, uri) {
    return diagnostics.map(pslLintDiagnostic => {
        const r = pslLintDiagnostic.range;
        const vscodeRange = new vscode.Range(r.start.line, r.start.character, r.end.line, r.end.character);
        const memberDiagnostic = new MemberDiagnostic(vscodeRange, pslLintDiagnostic.message, pslLintDiagnostic.severity);
        memberDiagnostic.source = pslLintDiagnostic.source;
        memberDiagnostic.code = pslLintDiagnostic.code;
        memberDiagnostic.ruleName = pslLintDiagnostic.ruleName;
        if (pslLintDiagnostic.member)
            memberDiagnostic.member = pslLintDiagnostic.member;
        if (pslLintDiagnostic.relatedInformation) {
            memberDiagnostic.relatedInformation = pslLintDiagnostic.relatedInformation.map(x => {
                return new vscode.DiagnosticRelatedInformation(new vscode.Location(uri, new vscode.Range(x.range.start.line, x.range.start.character, x.range.end.line, x.range.end.character)), x.message);
            });
        }
        return memberDiagnostic;
    });
}
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoiY29kZVF1YWxpdHkuanMiLCJzb3VyY2VSb290IjoiIiwic291cmNlcyI6WyIuLi8uLi8uLi9zcmMvbGFuZ3VhZ2UvY29kZVF1YWxpdHkudHMiXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6Ijs7Ozs7Ozs7Ozs7O0FBQUEsaUNBQWlDO0FBQ2pDLDRDQUErRDtBQUMvRCwyQ0FBMkM7QUFDM0Msa0RBQXFEO0FBQ3JELHNDQUFzQztBQUN0Qyw4Q0FBdUU7QUFDdkUsNkNBQWlEO0FBSWpELFNBQXNCLFFBQVEsQ0FBQyxPQUFnQzs7UUFFOUQsTUFBTSw0QkFBNEIsQ0FBQyxPQUFPLENBQUMsQ0FBQztRQUU1QyxNQUFNLGVBQWUsR0FBRyxNQUFNLENBQUMsU0FBUyxDQUFDLDBCQUEwQixDQUFDLFVBQVUsQ0FBQyxDQUFDO1FBQ2hGLE9BQU8sQ0FBQyxhQUFhLENBQUMsSUFBSSxDQUFDLGVBQWUsQ0FBQyxDQUFDO1FBRTVDLGdCQUFnQjtRQUNoQixJQUFJLFdBQVcsR0FBRyxJQUFJLE1BQU0sQ0FBQyx1QkFBdUIsRUFBRSxDQUFDO1FBRXZELElBQUksTUFBTSxDQUFDLE1BQU0sQ0FBQyxnQkFBZ0IsRUFBRTtZQUNuQyxZQUFZLENBQUMsTUFBTSxDQUFDLE1BQU0sQ0FBQyxnQkFBZ0IsQ0FBQyxRQUFRLEVBQUUsZUFBZSxFQUFFLFdBQVcsQ0FBQyxLQUFLLENBQUMsQ0FBQztTQUMxRjtRQUVELE1BQU0sQ0FBQyxNQUFNLENBQUMsMkJBQTJCLENBQUMsQ0FBQyxDQUFDLEVBQUU7WUFDN0MsSUFBSSxDQUFDLENBQUM7Z0JBQUUsT0FBTztZQUNmLFlBQVksQ0FBQyxDQUFDLENBQUMsUUFBUSxFQUFFLGVBQWUsRUFBRSxXQUFXLENBQUMsS0FBSyxDQUFDLENBQUM7UUFDOUQsQ0FBQyxDQUFDLENBQUM7UUFFSCxNQUFNLENBQUMsU0FBUyxDQUFDLHVCQUF1QixDQUFDLENBQUMsQ0FBQyxFQUFFO1lBQzVDLElBQUksQ0FBQyxDQUFDO2dCQUFFLE9BQU87WUFDZixXQUFXLENBQUMsTUFBTSxFQUFFLENBQUM7WUFDckIsV0FBVyxHQUFHLElBQUksTUFBTSxDQUFDLHVCQUF1QixFQUFFLENBQUM7WUFDbkQsWUFBWSxDQUFDLENBQUMsQ0FBQyxRQUFRLEVBQUUsZUFBZSxFQUFFLFdBQVcsQ0FBQyxLQUFLLENBQUMsQ0FBQztRQUM5RCxDQUFDLENBQUMsQ0FBQztRQUVILE1BQU0sQ0FBQyxTQUFTLENBQUMsc0JBQXNCLENBQUMsY0FBYyxDQUFDLEVBQUU7WUFDeEQsZUFBZSxDQUFDLE1BQU0sQ0FBQyxjQUFjLENBQUMsR0FBRyxDQUFDLENBQUM7UUFDNUMsQ0FBQyxDQUFDLENBQUM7UUFFSCxNQUFNLGNBQWMsR0FBRyxJQUFJLDhCQUFpQixFQUFFLENBQUM7UUFFL0MsS0FBSyxNQUFNLElBQUksSUFBSSxDQUFDLG9CQUFRLEVBQUUsc0JBQVUsRUFBRSxxQkFBUyxDQUFDLEVBQUU7WUFDckQsT0FBTyxDQUFDLGFBQWEsQ0FBQyxJQUFJLENBQ3pCLE1BQU0sQ0FBQyxTQUFTLENBQUMsMkJBQTJCLENBQzNDLElBQUksRUFBRSxjQUFjLENBQ3BCLENBQ0QsQ0FBQztTQUNGO0lBQ0YsQ0FBQztDQUFBO0FBdkNELDRCQXVDQztBQUVELFNBQWUsNEJBQTRCLENBQUMsT0FBZ0M7O1FBQzNFLE9BQU8sT0FBTyxDQUFDLEdBQUcsQ0FDakIsTUFBTSxDQUFDLFNBQVMsQ0FBQyxnQkFBZ0I7YUFDL0IsR0FBRyxDQUFDLFNBQVMsQ0FBQyxFQUFFLENBQUMsSUFBSSxNQUFNLENBQUMsZUFBZSxDQUFDLFNBQVMsRUFBRSxlQUFlLENBQUMsQ0FBQzthQUN4RSxHQUFHLENBQUMsQ0FBTSxPQUFPLEVBQUMsRUFBRTtZQUNwQixNQUFNLE9BQU8sR0FBRyxNQUFNLENBQUMsU0FBUyxDQUFDLHVCQUF1QixDQUFDLE9BQU8sQ0FBQyxDQUFDO1lBQ2xFLE9BQU8sQ0FBQyxhQUFhLENBQUMsSUFBSSxDQUFDLE9BQU8sQ0FBQyxXQUFXLENBQUMsR0FBRyxDQUFDLEVBQUU7Z0JBQ3BELElBQUEsa0JBQVMsRUFBQyxHQUFHLENBQUMsTUFBTSxDQUFDLENBQUM7WUFDdkIsQ0FBQyxDQUFDLEVBQUUsT0FBTyxDQUFDLFdBQVcsQ0FBQyxHQUFHLENBQUMsRUFBRTtnQkFDN0IsSUFBQSxrQkFBUyxFQUFDLEdBQUcsQ0FBQyxNQUFNLENBQUMsQ0FBQztZQUN2QixDQUFDLENBQUMsQ0FBQyxDQUFDO1lBQ0osT0FBTyxDQUFDLFdBQVcsQ0FBQyxHQUFHLENBQUMsRUFBRTtnQkFDekIsSUFBQSxxQkFBWSxFQUFDLEdBQUcsQ0FBQyxNQUFNLENBQUMsQ0FBQztZQUMxQixDQUFDLENBQUMsQ0FBQztZQUNILE1BQU0sSUFBSSxHQUFHLE1BQU0sTUFBTSxDQUFDLFNBQVMsQ0FBQyxTQUFTLENBQUMsT0FBTyxDQUFDLENBQUM7WUFDdkQsSUFBSSxDQUFDLElBQUksQ0FBQyxNQUFNO2dCQUFFLE9BQU87WUFDekIsTUFBTSxJQUFBLGtCQUFTLEVBQUMsSUFBSSxDQUFDLENBQUMsQ0FBQyxDQUFDLE1BQU0sQ0FBQyxDQUFDO1FBQ2pDLENBQUMsQ0FBQSxDQUFDLENBQ0gsQ0FBQztJQUNILENBQUM7Q0FBQTtBQUVELE1BQWEsZ0JBQWlCLFNBQVEsTUFBTSxDQUFDLFVBQVU7Q0FHdEQ7QUFIRCw0Q0FHQztBQUVELFNBQVMsWUFBWSxDQUNwQixZQUFpQyxFQUNqQyxlQUE0QyxFQUM1QyxpQkFBMkM7SUFFM0MsSUFBSSxDQUFDLEdBQUcsQ0FBQyxnQkFBZ0IsQ0FBQyxrQkFBa0IsQ0FBQyxZQUFZLENBQUMsUUFBUSxDQUFDO1FBQUUsT0FBTztJQUU1RSxNQUFNLGVBQWUsR0FBZSxNQUFNLENBQUMsU0FBUyxDQUFDLGdCQUFnQixDQUFDLEtBQUssRUFBRSxZQUFZLENBQUMsR0FBRyxDQUFDLENBQUMsR0FBRyxDQUFDLE1BQU0sQ0FBQyxDQUFDO0lBRTNHLElBQUksU0FBUyxHQUFHLEtBQUssQ0FBQztJQUN0QixJQUFJLGVBQWUsS0FBSyxRQUFRLEVBQUU7UUFDakMsOEJBQThCO1FBQzlCLE1BQU0sTUFBTSxHQUFHLElBQUEsa0JBQVMsRUFBQyxZQUFZLENBQUMsR0FBRyxDQUFDLE1BQU0sQ0FBQyxDQUFDO1FBQ2xELElBQUksQ0FBQyxNQUFNO1lBQUUsT0FBTztRQUNwQixTQUFTLEdBQUcsSUFBSSxDQUFDO0tBQ2pCO1NBQ0ksSUFBSSxlQUFlLEtBQUssS0FBSyxJQUFJLGVBQWUsS0FBSyxJQUFJLEVBQUU7UUFDL0QsZUFBZSxDQUFDLEtBQUssRUFBRSxDQUFDO1FBQ3hCLE9BQU87S0FDUDtJQUVELE9BQU8sQ0FBQyxRQUFRLENBQUMsR0FBRyxFQUFFO1FBQ3JCLElBQUksQ0FBQyxpQkFBaUIsQ0FBQyx1QkFBdUIsRUFBRTtZQUMvQyxJQUFJLENBQUMsWUFBWSxFQUFFLFNBQVMsRUFBRSxpQkFBaUIsRUFBRSxlQUFlLENBQUMsQ0FBQztTQUNsRTtJQUNGLENBQUMsQ0FBQyxDQUFDO0FBQ0osQ0FBQztBQUVELFNBQVMsSUFBSSxDQUNaLFlBQWlDLEVBQ2pDLFNBQWtCLEVBQUUsaUJBQTJDLEVBQy9ELGVBQTRDO0lBRTVDLE1BQU0sZ0JBQWdCLEdBQXlCLGVBQWUsQ0FBQyxZQUFZLENBQUMsQ0FBQztJQUM3RSxNQUFNLGNBQWMsR0FBRyxHQUFHLENBQUMsZ0JBQWdCLENBQUMsS0FBSyxDQUFDLGdCQUFnQixDQUFDLE1BQU0sQ0FBQyxDQUFDLENBQUM7UUFDM0UsTUFBTSxDQUFDLFNBQVMsQ0FBQyxZQUFZLENBQUMsT0FBTyxFQUFFLENBQUMsQ0FBQyxDQUFDLENBQUMsU0FBUyxDQUFDO0lBQ3RELE1BQU0sV0FBVyxHQUFHLElBQUEseUJBQWMsRUFBQyxnQkFBZ0IsRUFBRSxjQUFjLEVBQUUsU0FBUyxDQUFDLENBQUM7SUFDaEYsTUFBTSxpQkFBaUIsR0FBRyxTQUFTLENBQUMsV0FBVyxFQUFFLFlBQVksQ0FBQyxHQUFHLENBQUMsQ0FBQztJQUNuRSxPQUFPLENBQUMsUUFBUSxDQUFDLEdBQUcsRUFBRTtRQUNyQixJQUFJLENBQUMsaUJBQWlCLENBQUMsdUJBQXVCLEVBQUU7WUFDL0MsZUFBZSxDQUFDLEdBQUcsQ0FBQyxZQUFZLENBQUMsR0FBRyxFQUFFLGlCQUFpQixDQUFDLENBQUM7U0FDekQ7SUFDRixDQUFDLENBQUMsQ0FBQztBQUNKLENBQUM7QUFFRCxTQUFTLGVBQWUsQ0FBQyxZQUFpQztJQUN6RCxNQUFNLGFBQWEsR0FBRyxDQUFDLENBQVMsRUFBRSxFQUFFLENBQUMsWUFBWSxDQUFDLE1BQU0sQ0FBQyxDQUFDLENBQUMsQ0FBQyxJQUFJLENBQUM7SUFDakUsTUFBTSxnQkFBZ0IsR0FBRyxJQUFJLEdBQUcsQ0FBQyxnQkFBZ0IsQ0FBQyxZQUFZLENBQUMsR0FBRyxDQUFDLE1BQU0sRUFBRSxZQUFZLENBQUMsT0FBTyxFQUFFLEVBQUUsYUFBYSxDQUFDLENBQUM7SUFDbEgsT0FBTyxnQkFBZ0IsQ0FBQztBQUN6QixDQUFDO0FBRUQsU0FBUyxTQUFTLENBQUMsV0FBNkIsRUFBRSxHQUFlO0lBQ2hFLE9BQU8sV0FBVyxDQUFDLEdBQUcsQ0FBQyxpQkFBaUIsQ0FBQyxFQUFFO1FBQzFDLE1BQU0sQ0FBQyxHQUFHLGlCQUFpQixDQUFDLEtBQUssQ0FBQztRQUNsQyxNQUFNLFdBQVcsR0FBRyxJQUFJLE1BQU0sQ0FBQyxLQUFLLENBQUMsQ0FBQyxDQUFDLEtBQUssQ0FBQyxJQUFJLEVBQUUsQ0FBQyxDQUFDLEtBQUssQ0FBQyxTQUFTLEVBQUUsQ0FBQyxDQUFDLEdBQUcsQ0FBQyxJQUFJLEVBQUUsQ0FBQyxDQUFDLEdBQUcsQ0FBQyxTQUFTLENBQUMsQ0FBQztRQUNuRyxNQUFNLGdCQUFnQixHQUFHLElBQUksZ0JBQWdCLENBQUMsV0FBVyxFQUFFLGlCQUFpQixDQUFDLE9BQU8sRUFBRSxpQkFBaUIsQ0FBQyxRQUFRLENBQUMsQ0FBQztRQUNsSCxnQkFBZ0IsQ0FBQyxNQUFNLEdBQUcsaUJBQWlCLENBQUMsTUFBTSxDQUFDO1FBQ25ELGdCQUFnQixDQUFDLElBQUksR0FBRyxpQkFBaUIsQ0FBQyxJQUFJLENBQUM7UUFDL0MsZ0JBQWdCLENBQUMsUUFBUSxHQUFHLGlCQUFpQixDQUFDLFFBQVEsQ0FBQztRQUN2RCxJQUFJLGlCQUFpQixDQUFDLE1BQU07WUFBRSxnQkFBZ0IsQ0FBQyxNQUFNLEdBQUcsaUJBQWlCLENBQUMsTUFBTSxDQUFDO1FBQ2pGLElBQUksaUJBQWlCLENBQUMsa0JBQWtCLEVBQUU7WUFDekMsZ0JBQWdCLENBQUMsa0JBQWtCLEdBQUcsaUJBQWlCLENBQUMsa0JBQWtCLENBQUMsR0FBRyxDQUFDLENBQUMsQ0FBQyxFQUFFO2dCQUNsRixPQUFPLElBQUksTUFBTSxDQUFDLDRCQUE0QixDQUM3QyxJQUFJLE1BQU0sQ0FBQyxRQUFRLENBQUMsR0FBRyxFQUN0QixJQUFJLE1BQU0sQ0FBQyxLQUFLLENBQUMsQ0FBQyxDQUFDLEtBQUssQ0FBQyxLQUFLLENBQUMsSUFBSSxFQUNsQyxDQUFDLENBQUMsS0FBSyxDQUFDLEtBQUssQ0FBQyxTQUFTLEVBQ3ZCLENBQUMsQ0FBQyxLQUFLLENBQUMsR0FBRyxDQUFDLElBQUksRUFDaEIsQ0FBQyxDQUFDLEtBQUssQ0FBQyxHQUFHLENBQUMsU0FBUyxDQUNyQixDQUFDLEVBQ0gsQ0FBQyxDQUFDLE9BQU8sQ0FDVCxDQUFDO1lBQ0gsQ0FBQyxDQUFDLENBQUM7U0FDSDtRQUNELE9BQU8sZ0JBQWdCLENBQUM7SUFDekIsQ0FBQyxDQUFDLENBQUM7QUFDSixDQUFDIn0=