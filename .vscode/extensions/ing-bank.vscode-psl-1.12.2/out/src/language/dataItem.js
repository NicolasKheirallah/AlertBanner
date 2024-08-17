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
exports.DataDocumentHighlightProvider = exports.DataHoverProvider = void 0;
const path = require("path");
const vscode = require("vscode");
const jsonc = require("jsonc-parser");
const fs = require("fs-extra");
function getEnvBase(fileName) {
    return vscode.workspace.getWorkspaceFolder(vscode.Uri.file(fileName)).uri.fsPath;
}
class DataHoverProvider {
    provideHover(document, position) {
        return __awaiter(this, void 0, void 0, function* () {
            // array of column names
            let columnNames = document.lineAt(0).text.split('\t');
            // the text up to the cursor
            let textToPosition = document.getText(new vscode.Range(position.line, 0, position.line, position.character));
            // position of current data item
            let currentDataItemPosition = textToPosition.split('\t').length - 1;
            // full text of data item
            let dataItemText = document.lineAt(position.line).text.split('\t')[currentDataItemPosition];
            let prevTabPos = textToPosition.lastIndexOf('\t') + 1;
            let nextTabPos = prevTabPos + dataItemText.length;
            if (currentDataItemPosition <= columnNames.length) {
                let columnName = columnNames[currentDataItemPosition];
                let tableName = path.basename(document.fileName).replace('.DAT', '');
                let fileName = `${tableName.toUpperCase()}-${columnName.toUpperCase()}.COL`;
                let link = path.join(getEnvBase(document.fileName), 'dataqwik', 'table', `${tableName.toLowerCase()}`, `${fileName}`);
                let content;
                if (!fs.existsSync(link)) {
                    content = new vscode.MarkdownString(`COLUMN: **${columnName}**`);
                }
                else {
                    let uri = vscode.Uri.file(link);
                    let tbl = yield vscode.workspace.openTextDocument(uri);
                    let tblJSON = jsonc.parse(tbl.getText());
                    content = new vscode.MarkdownString(`COLUMN: **[${columnName}](command:vscode.open?${encodeURIComponent(JSON.stringify(uri))})** (*${tblJSON['DES']}*)`);
                }
                content.isTrusted = true;
                return new vscode.Hover(content, new vscode.Range(position.line, prevTabPos, position.line, nextTabPos));
            }
            return undefined;
        });
    }
}
exports.DataHoverProvider = DataHoverProvider;
class DataDocumentHighlightProvider {
    provideDocumentHighlights(document, position) {
        return __awaiter(this, void 0, void 0, function* () {
            // the text up to the cursor
            let textToPosition = document.getText(new vscode.Range(position.line, 0, position.line, position.character));
            // position of current data item
            let currentDataItemPosition = textToPosition.split('\t').length - 1;
            let highlights = [];
            for (let lineNumber = 0; lineNumber < document.lineCount; lineNumber++) {
                let text = document.lineAt(lineNumber).text;
                if (!text)
                    continue;
                let row = document.lineAt(lineNumber).text.split('\t');
                let dataItemText = row[currentDataItemPosition];
                let textToPosition = row.slice(0, currentDataItemPosition + 1).join('\t');
                let prevTabCol = textToPosition.lastIndexOf('\t') + 1;
                let nextTabCol = prevTabCol + dataItemText.length;
                let range = new vscode.Range(lineNumber, prevTabCol, lineNumber, nextTabCol);
                document.validateRange(range);
                let highlight = new vscode.DocumentHighlight(range, vscode.DocumentHighlightKind.Write);
                highlights.push(highlight);
            }
            return highlights;
        });
    }
}
exports.DataDocumentHighlightProvider = DataDocumentHighlightProvider;
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoiZGF0YUl0ZW0uanMiLCJzb3VyY2VSb290IjoiIiwic291cmNlcyI6WyIuLi8uLi8uLi9zcmMvbGFuZ3VhZ2UvZGF0YUl0ZW0udHMiXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6Ijs7Ozs7Ozs7Ozs7O0FBQUEsNkJBQTZCO0FBQzdCLGlDQUFpQztBQUNqQyxzQ0FBc0M7QUFDdEMsK0JBQThCO0FBRTlCLFNBQVMsVUFBVSxDQUFDLFFBQWdCO0lBQ25DLE9BQU8sTUFBTSxDQUFDLFNBQVMsQ0FBQyxrQkFBa0IsQ0FBQyxNQUFNLENBQUMsR0FBRyxDQUFDLElBQUksQ0FBQyxRQUFRLENBQUMsQ0FBQyxDQUFDLEdBQUcsQ0FBQyxNQUFNLENBQUE7QUFDakYsQ0FBQztBQUVELE1BQWEsaUJBQWlCO0lBQ2hCLFlBQVksQ0FBQyxRQUE2QixFQUFFLFFBQXlCOztZQUVqRix3QkFBd0I7WUFDeEIsSUFBSSxXQUFXLEdBQWtCLFFBQVEsQ0FBQyxNQUFNLENBQUMsQ0FBQyxDQUFDLENBQUMsSUFBSSxDQUFDLEtBQUssQ0FBQyxJQUFJLENBQUMsQ0FBQztZQUVyRSw0QkFBNEI7WUFDNUIsSUFBSSxjQUFjLEdBQVcsUUFBUSxDQUFDLE9BQU8sQ0FBQyxJQUFJLE1BQU0sQ0FBQyxLQUFLLENBQUMsUUFBUSxDQUFDLElBQUksRUFBRSxDQUFDLEVBQUUsUUFBUSxDQUFDLElBQUksRUFBRSxRQUFRLENBQUMsU0FBUyxDQUFDLENBQUMsQ0FBQztZQUVySCxnQ0FBZ0M7WUFDaEMsSUFBSSx1QkFBdUIsR0FBVyxjQUFjLENBQUMsS0FBSyxDQUFDLElBQUksQ0FBQyxDQUFDLE1BQU0sR0FBRyxDQUFDLENBQUM7WUFFNUUseUJBQXlCO1lBQ3pCLElBQUksWUFBWSxHQUFHLFFBQVEsQ0FBQyxNQUFNLENBQUMsUUFBUSxDQUFDLElBQUksQ0FBQyxDQUFDLElBQUksQ0FBQyxLQUFLLENBQUMsSUFBSSxDQUFDLENBQUMsdUJBQXVCLENBQUMsQ0FBQztZQUU1RixJQUFJLFVBQVUsR0FBVyxjQUFjLENB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