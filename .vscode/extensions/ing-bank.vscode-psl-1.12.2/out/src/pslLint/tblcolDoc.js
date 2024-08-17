"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.TblColDocumentation = void 0;
const path = require("path");
const tokenizer_1 = require("../parser/tokenizer");
const api_1 = require("./api");
/**
 * Checks whether table and columns are created with documentation.
 */
class TblColDocumentation extends api_1.FileDefinitionRule {
    report() {
        const baseName = path.basename(this.profileComponent.fsPath);
        const diagnostics = [];
        const bracketMatch = this.profileComponent.textDocument.match(/^}/m);
        // Exit if no match found
        if (!bracketMatch)
            return [];
        const charcterOffset = bracketMatch.index;
        const endPos = this.profileComponent.textDocument.length;
        const tblColDoc = this.profileComponent.textDocument.substring(charcterOffset + 1, endPos).trim();
        if (!tblColDoc) {
            let message;
            if (baseName.endsWith('TBL')) {
                message = `Documentation missing for table definition "${baseName}".`;
            }
            else
                message = `Documentation missing for data item "${baseName}".`;
            const position = this.profileComponent.positionAt(charcterOffset);
            const range = new tokenizer_1.Range(position, position);
            diagnostics.push(addDiagnostic(range, message, this.ruleName));
        }
        return diagnostics;
    }
}
exports.TblColDocumentation = TblColDocumentation;
function addDiagnostic(range, message, ruleName) {
    const diagnostic = new api_1.Diagnostic(range, message, ruleName, api_1.DiagnosticSeverity.Information);
    diagnostic.source = 'lint';
    return diagnostic;
}
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoidGJsY29sRG9jLmpzIiwic291cmNlUm9vdCI6IiIsInNvdXJjZXMiOlsiLi4vLi4vLi4vc3JjL3BzbExpbnQvdGJsY29sRG9jLnRzIl0sIm5hbWVzIjpbXSwibWFwcGluZ3MiOiI7OztBQUFBLDZCQUE2QjtBQUM3QixtREFBNEM7QUFDNUMsK0JBQTJFO0FBRTNFOztHQUVHO0FBQ0gsTUFBYSxtQkFBb0IsU0FBUSx3QkFBa0I7SUFFMUQsTUFBTTtRQUNMLE1BQU0sUUFBUSxHQUFHLElBQUksQ0FBQyxRQUFRLENBQUMsSUFBSSxDQUFDLGdCQUFnQixDQUFDLE1BQU0sQ0FBQyxDQUFDO1FBRTdELE1BQU0sV0FBVyxHQUFpQixFQUFFLENBQUM7UUFDckMsTUFBTSxZQUFZLEdBQUcsSUFBSSxDQUFDLGdCQUFnQixDQUFDLFlBQVksQ0FBQyxLQUFLLENBQUMsS0FBSyxDQUFDLENBQUM7UUFDckUseUJBQXlCO1FBQ3pCLElBQUksQ0FBQyxZQUFZO1lBQUUsT0FBTyxFQUFFLENBQUM7UUFFN0IsTUFBTSxjQUFjLEdBQUcsWUFBWSxDQUFDLEtBQUssQ0FBQztRQUMxQyxNQUFNLE1BQU0sR0FBRyxJQUFJLENBQUMsZ0JBQWdCLENBQUMsWUFBWSxDQUFDLE1BQU0sQ0FBQztRQUN6RCxNQUFNLFNBQVMsR0FBRyxJQUFJLENBQUMsZ0JBQWdCLENBQUMsWUFBWSxDQUFDLFNBQVMsQ0FBQyxjQUFjLEdBQUcsQ0FBQyxFQUFFLE1BQU0sQ0FBQyxDQUFDLElBQUksRUFBRSxDQUFDO1FBRWxHLElBQUksQ0FBQyxTQUFTLEVBQUU7WUFDZixJQUFJLE9BQU8sQ0FBQztZQUVaLElBQUksUUFBUSxDQUFDLFFBQVEsQ0FBQyxLQUFLLENBQUMsRUFBRTtnQkFDN0IsT0FBTyxHQUFHLCtDQUErQyxRQUFRLElBQUksQ0FBQzthQUN0RTs7Z0JBQ0ksT0FBTyxHQUFHLHdDQUF3QyxRQUFRLElBQUksQ0FBQztZQUNwRSxNQUFNLFFBQVEsR0FBRyxJQUFJLENBQUMsZ0JBQWdCLENBQUMsVUFBVSxDQUFDLGNBQWMsQ0FBQyxDQUFDO1lBQ2xFLE1BQU0sS0FBSyxHQUFHLElBQUksaUJBQUssQ0FBQyxRQUFRLEVBQUUsUUFBUSxDQUFDLENBQUM7WUFDNUMsV0FBVyxDQUFDLElBQUksQ0FBQyxhQUFhLENBQUMsS0FBSyxFQUFFLE9BQU8sRUFBRSxJQUFJLENBQUMsUUFBUSxDQUFDLENBQUMsQ0FBQztTQUMvRDtRQUVELE9BQU8sV0FBVyxDQUFDO0lBQ3BCLENBQUM7Q0FFRDtBQTdCRCxrREE2QkM7QUFFRCxTQUFTLGFBQWEsQ0FBQyxLQUFZLEVBQUUsT0FBZSxFQUFFLFFBQWdCO0lBQ3JFLE1BQU0sVUFBVSxHQUFHLElBQUksZ0JBQVUsQ0FBQyxLQUFLLEVBQUUsT0FBTyxFQUFFLFFBQVEsRUFBRSx3QkFBa0IsQ0FBQyxXQUFXLENBQUMsQ0FBQztJQUM1RixVQUFVLENBQUMsTUFBTSxHQUFHLE1BQU0sQ0FBQztJQUMzQixPQUFPLFVBQVUsQ0FBQztBQUNuQixDQUFDIn0=