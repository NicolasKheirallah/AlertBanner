"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.PSLDiagnostic = void 0;
const vscode = require("vscode");
class PSLDiagnostic {
    constructor(message, severity, file, range) {
        this.message = message;
        this.severity = severity;
        this.file = file;
        this.range = range;
        this.diagnostic = new vscode.Diagnostic(this.range, this.message, this.severity);
    }
    static setDiagnostics(pslDiagnostics, envName, fsPath) {
        let diagnosticMap = new Map();
        pslDiagnostics.forEach(pslDiagnostic => {
            let canonicalFile = vscode.Uri.file(pslDiagnostic.file).toString();
            let diagnostics = diagnosticMap.get(canonicalFile);
            pslDiagnostic.diagnostic.source = envName;
            if (!diagnostics) {
                diagnostics = [];
            }
            diagnostics.push(pslDiagnostic.diagnostic);
            diagnosticMap.set(canonicalFile, diagnostics);
        });
        let collection = this.diagnosticCollections.find(col => col.name === envName);
        if (!collection) {
            collection = this.registerCollection(envName);
        }
        let uri = vscode.Uri.file(fsPath);
        collection.delete(uri);
        diagnosticMap.forEach((diags, file) => {
            collection.set(vscode.Uri.parse(file), diags);
        });
    }
    static registerCollection(envName) {
        let collection = vscode.languages.createDiagnosticCollection(envName);
        vscode.workspace.onDidCloseTextDocument((textDocument) => {
            let uri = textDocument.uri;
            collection.delete(uri);
        });
        this.diagnosticCollections.push(collection);
        return collection;
    }
}
exports.PSLDiagnostic = PSLDiagnostic;
PSLDiagnostic.diagnosticCollections = [];
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoiZGlhZ25vc3RpY3MuanMiLCJzb3VyY2VSb290IjoiIiwic291cmNlcyI6WyIuLi8uLi8uLi9zcmMvY29tbW9uL2RpYWdub3N0aWNzLnRzIl0sIm5hbWVzIjpbXSwibWFwcGluZ3MiOiI7OztBQUFBLGlDQUFpQztBQUVqQyxNQUFhLGFBQWE7SUE0Q3pCLFlBQVksT0FBZSxFQUFFLFFBQW1DLEVBQUUsSUFBWSxFQUFFLEtBQW1CO1FBQ2xHLElBQUksQ0FBQyxPQUFPLEdBQUcsT0FBTyxDQUFDO1FBQ3ZCLElBQUksQ0FBQyxRQUFRLEdBQUcsUUFBUSxDQUFDO1FBQ3pCLElBQUksQ0FBQyxJQUFJLEdBQUcsSUFBSSxDQUFDO1FBQ2pCLElBQUksQ0FBQyxLQUFLLEdBQUcsS0FBSyxDQUFDO1FBRW5CLElBQUksQ0FBQyxVQUFVLEdBQUcsSUFBSSxNQUFNLENBQUMsVUFBVSxDQUFDLElBQUksQ0FBQyxLQUFLLEVBQUUsSUFBSSxDQUFDLE9BQU8sRUFBRSxJQUFJLENBQUMsUUFBUSxDQUFDLENBQUM7SUFDbEYsQ0FBQztJQS9DRCxNQUFNLENBQUMsY0FBYyxDQUFDLGNBQStCLEVBQUUsT0FBZSxFQUFFLE1BQWM7UUFDckYsSUFBSSxhQUFhLEdBQXFDLElBQUksR0FBRyxFQUFFLENBQUM7UUFDaEUsY0FBYyxDQUFDLE9BQU8sQ0FBQyxhQUFhLENBQUMsRUFBRTtZQUN0QyxJQUFJLGFBQWEsR0FBRyxNQUFNLENBQUMsR0FBRyxDQUFDLElBQUksQ0FBQyxhQUFhLENBQUMsSUFBSSxDQUFDLENBQUMsUUFBUSxFQUFFLENBQUM7WUFDbkUsSUFBSSxXQUFXLEdBQUcsYUFBYSxDQUFDLEdBQUcsQ0FBQyxhQUFhLENBQUMsQ0FBQztZQUNuRCxhQUFhLENBQUMsVUFBVSxDQUFDLE1BQU0sR0FBRyxPQUFPLENBQUM7WUFDMUMsSUFBSSxDQUFDLFdBQVcsRUFBRTtnQkFBRSxXQUFXLEdBQUcsRUFBRSxDQUFDO2FBQUU7WUFFdkMsV0FBVyxDQUFDLElBQUksQ0FBQyxhQUFhLENBQUMsVUFBVSxDQUFDLENBQUM7WUFDM0MsYUFBYSxDQUFDLEdBQUcsQ0FBQyxhQUFhLEVBQUUsV0FBVyxDQUFDLENBQUM7UUFDL0MsQ0FBQyxDQUFDLENBQUM7UUFDSCxJQUFJLFVBQVUsR0FBRyxJQUFJLENBQUMscUJBQXFCLENBQUMsSUFBSSxDQUFDLEdBQUcsQ0FBQyxFQUFFLENBQUMsR0FBRyxDQUFDLElBQUksS0FBSyxPQUFPLENBQUMsQ0FBQztRQUM5RSxJQUFJLENBQUMsVUFBVSxFQUFHO1lBQ2pCLFVBQVUsR0FBRyxJQUFJLENBQUMsa0JBQWtCLENBQUMsT0FBTyxDQUFDLENBQUM7U0FDOUM7UUFDRCxJQUFJLEdBQUcsR0FBRyxNQUFNLENBQUMsR0FBRyxDQUFDLElBQUksQ0FBQyxNQUFNLENBQUMsQ0FBQztRQUNsQyxVQUFVLENBQUMsTUFBTSxDQUFDLEdBQUcsQ0FBQyxDQUFDO1FBQ3ZCLGFBQWEsQ0FBQyxPQUFPLENBQUMsQ0FBQyxLQUFLLEVBQUUsSUFBSSxFQUFFLEVBQUU7WUFDckMsVUFBVSxDQUFDLEdBQUcsQ0FBQyxNQUFNLENBQUMsR0FBRyxDQUFDLEtBQUssQ0FBQyxJQUFJLENBQUMsRUFBRSxLQUFLLENBQUMsQ0FBQztRQUUvQyxDQUFDLENBQUMsQ0FBQztJQUNKLENBQUM7SUFFRCxNQUFNLENBQUMsa0JBQWtCLENBQUMsT0FBZTtRQUN4QyxJQUFJLFVBQVUsR0FBRyxNQUFNLENBQUMsU0FBUyxDQUFDLDBCQUEwQixDQUFDLE9BQU8sQ0FBQyxDQUFDO1FBQ3RFLE1BQU0sQ0FBQyxTQUFTLENBQUMsc0JBQXNCLENBQUMsQ0FBQyxZQUFZLEVBQUUsRUFBRTtZQUN4RCxJQUFJLEdBQUcsR0FBRyxZQUFZLENBQUMsR0FBRyxDQUFDO1lBQzNCLFVBQVUsQ0FBQyxNQUFNLENBQUMsR0FBRyxDQUFDLENBQUM7UUFDeEIsQ0FBQyxDQUFDLENBQUE7UUFDRixJQUFJLENBQUMscUJBQXFCLENBQUMsSUFBSSxDQUFDLFVBQVUsQ0FBQyxDQUFDO1FBQzVDLE9BQU8sVUFBVSxDQUFDO0lBQ25CLENBQUM7O0FBbkNGLHNDQW9EQztBQWxETyxtQ0FBcUIsR0FBa0MsRUFBRSxDQUFDIn0=