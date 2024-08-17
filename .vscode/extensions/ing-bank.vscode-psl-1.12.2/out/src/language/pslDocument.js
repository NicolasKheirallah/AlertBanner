"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.MumpsDocumentSymbolProvider = exports.PSLDocumentSymbolProvider = void 0;
const vscode = require("vscode");
const parser = require("../parser/parser");
const mumps_1 = require("./mumps");
class PSLDocumentSymbolProvider {
    provideDocumentSymbols(document) {
        return new Promise(resolve => {
            const parsedDoc = parser.parseText(document.getText());
            const symbols = [];
            parsedDoc.methods.forEach(method => {
                symbols.push(createMethodSymbol(method, document));
            });
            parsedDoc.properties.forEach(property => {
                const propertyNameToken = property.id;
                const name = propertyNameToken.value;
                const containerName = '';
                const position = propertyNameToken.position;
                const location = new vscode.Location(document.uri, new vscode.Position(position.line, position.character));
                symbols.push(new vscode.SymbolInformation(name, vscode.SymbolKind.Property, containerName, location));
            });
            resolve(symbols);
        });
    }
}
exports.PSLDocumentSymbolProvider = PSLDocumentSymbolProvider;
/**
 * Outline provider for MUMPS
 */
class MumpsDocumentSymbolProvider {
    provideDocumentSymbols(document) {
        const symbols = [];
        const parsedDoc = this.getParsedDoc(document);
        parsedDoc.methods.forEach(method => {
            symbols.push(createMethodSymbol(method, document));
        });
        return symbols;
    }
    getParsedDoc(document) {
        const cachedMumps = (0, mumps_1.getVirtualDocument)(document.uri);
        if (cachedMumps)
            return cachedMumps.parsedDocument;
        else
            return parser.parseText(document.getText());
    }
}
exports.MumpsDocumentSymbolProvider = MumpsDocumentSymbolProvider;
function createMethodSymbol(method, document) {
    const methodToken = method.id;
    const name = methodToken.value;
    const containerName = '';
    const startPosition = new vscode.Position(methodToken.position.line, 0);
    let endPositionNumber = method.endLine;
    if (endPositionNumber === -1)
        endPositionNumber = document.lineCount - 1; // last line
    const endPosition = new vscode.Position(endPositionNumber, 0);
    const methodRange = new vscode.Location(document.uri, new vscode.Range(startPosition, endPosition));
    const kind = method.batch ? vscode.SymbolKind.Module : vscode.SymbolKind.Function;
    return new vscode.SymbolInformation(name, kind, containerName, methodRange);
}
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoicHNsRG9jdW1lbnQuanMiLCJzb3VyY2VSb290IjoiIiwic291cmNlcyI6WyIuLi8uLi8uLi9zcmMvbGFuZ3VhZ2UvcHNsRG9jdW1lbnQudHMiXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6Ijs7O0FBQUEsaUNBQWlDO0FBQ2pDLDJDQUEyQztBQUMzQyxtQ0FBNkM7QUFFN0MsTUFBYSx5QkFBeUI7SUFFOUIsc0JBQXNCLENBQUMsUUFBNkI7UUFDMUQsT0FBTyxJQUFJLE9BQU8sQ0FBQyxPQUFPLENBQUMsRUFBRTtZQUM1QixNQUFNLFNBQVMsR0FBRyxNQUFNLENBQUMsU0FBUyxDQUFDLFFBQVEsQ0FBQyxPQUFPLEVBQUUsQ0FBQyxDQUFDO1lBQ3ZELE1BQU0sT0FBTyxHQUErQixFQUFFLENBQUM7WUFDL0MsU0FBUyxDQUFDLE9BQU8sQ0FBQyxPQUFPLENBQUMsTUFBTSxDQUFDLEVBQUU7Z0JBQ2xDLE9BQU8sQ0FBQyxJQUFJLENBQUMsa0JBQWtCLENBQUMsTUFBTSxFQUFFLFFBQVEsQ0FBQyxDQUFDLENBQUM7WUFDcEQsQ0FBQyxDQUFDLENBQUM7WUFDSCxTQUFTLENBQUMsVUFBVSxDQUFDLE9BQU8sQ0FBQyxRQUFRLENBQUMsRUFBRTtnQkFDdkMsTUFBTSxpQkFBaUIsR0FBRyxRQUFRLENBQUMsRUFBRSxDQUFDO2dCQUN0QyxNQUFNLElBQUksR0FBRyxpQkFBaUIsQ0FBQyxLQUFLLENBQUM7Z0JBQ3JDLE1BQU0sYUFBYSxHQUFHLEVBQUUsQ0FBQztnQkFDekIsTUFBTSxRQUFRLEdBQUcsaUJBQWlCLENBQUMsUUFBUSxDQUFDO2dCQUM1QyxNQUFNLFFBQVEsR0FBRyxJQUFJLE1BQU0sQ0FBQyxRQUFRLENBQUMsUUFBUSxDQUFDLEdBQUcsRUFBRSxJQUFJLE1BQU0sQ0FBQyxRQUFRLENBQUMsUUFBUSxDQUFDLElBQUksRUFBRSxRQUFRLENBQUMsU0FBUyxDQUFDLENBQUMsQ0FBQztnQkFDM0csT0FBTyxDQUFDLElBQUksQ0FBQyxJQUFJLE1BQU0sQ0FBQyxpQkFBaUIsQ0FBQyxJQUFJLEVBQUUsTUFBTSxDQUFDLFVBQVUsQ0FBQyxRQUFRLEVBQUUsYUFBYSxFQUFFLFFBQVEsQ0FBQyxDQUFDLENBQUM7WUFDdkcsQ0FBQyxDQUFDLENBQUM7WUFDSCxPQUFPLENBQUMsT0FBTyxDQUFDLENBQUM7UUFDbEIsQ0FBQyxDQUFDLENBQUM7SUFDSixDQUFDO0NBQ0Q7QUFwQkQsOERBb0JDO0FBRUQ7O0dBRUc7QUFDSCxNQUFhLDJCQUEyQjtJQUVoQyxzQkFBc0IsQ0FBQyxRQUE2QjtRQUMxRCxNQUFNLE9BQU8sR0FBK0IsRUFBRSxDQUFDO1FBQy9DLE1BQU0sU0FBUyxHQUFHLElBQUksQ0FBQyxZQUFZLENBQUMsUUFBUSxDQUFDLENBQUM7UUFDOUMsU0FBUyxDQUFDLE9BQU8sQ0FBQyxPQUFPLENBQUMsTUFBTSxDQUFDLEVBQUU7WUFDbEMsT0FBTyxDQUFDLElBQUksQ0FBQyxrQkFBa0IsQ0FBQyxNQUFNLEVBQUUsUUFBUSxDQUFDLENBQUMsQ0FBQztRQUNwRCxDQUFDLENBQUMsQ0FBQztRQUNILE9BQU8sT0FBTyxDQUFDO0lBQ2hCLENBQUM7SUFFRCxZQUFZLENBQUMsUUFBNkI7UUFDekMsTUFBTSxXQUFXLEdBQUcsSUFBQSwwQkFBa0IsRUFBQyxRQUFRLENBQUMsR0FBRyxDQUFDLENBQUM7UUFDckQsSUFBSSxXQUFXO1lBQUUsT0FBTyxXQUFXLENBQUMsY0FBYyxDQUFDOztZQUM5QyxPQUFPLE1BQU0sQ0FBQyxTQUFTLENBQUMsUUFBUSxDQUFDLE9BQU8sRUFBRSxDQUFDLENBQUM7SUFDbEQsQ0FBQztDQUNEO0FBaEJELGtFQWdCQztBQUVELFNBQVMsa0JBQWtCLENBQUMsTUFBcUIsRUFBRSxRQUE2QjtJQUMvRSxNQUFNLFdBQVcsR0FBRyxNQUFNLENBQUMsRUFBRSxDQUFDO0lBQzlCLE1BQU0sSUFBSSxHQUFHLFdBQVcsQ0FBQyxLQUFLLENBQUM7SUFDL0IsTUFBTSxhQUFhLEdBQUcsRUFBRSxDQUFDO0lBRXpCLE1BQU0sYUFBYSxHQUFHLElBQUksTUFBTSxDQUFDLFFBQVEsQ0FBQyxXQUFXLENBQUMsUUFBUSxDQUFDLElBQUksRUFBRSxDQUFDLENBQUMsQ0FBQztJQUV4RSxJQUFJLGlCQUFpQixHQUFHLE1BQU0sQ0FBQyxPQUFPLENBQUM7SUFDdkMsSUFBSSxpQkFBaUIsS0FBSyxDQUFDLENBQUM7UUFBRSxpQkFBaUIsR0FBRyxRQUFRLENBQUMsU0FBUyxHQUFHLENBQUMsQ0FBQyxDQUFDLFlBQVk7SUFDdEYsTUFBTSxXQUFXLEdBQUcsSUFBSSxNQUFNLENBQUMsUUFBUSxDQUFDLGlCQUFpQixFQUFFLENBQUMsQ0FBQyxDQUFDO0lBQzlELE1BQU0sV0FBVyxHQUFHLElBQUksTUFBTSxDQUFDLFFBQVEsQ0FBQyxRQUFRLENBQUMsR0FBRyxFQUFFLElBQUksTUFBTSxDQUFDLEtBQUssQ0FBQyxhQUFhLEVBQUUsV0FBVyxDQUFDLENBQUMsQ0FBQztJQUNwRyxNQUFNLElBQUksR0FBRyxNQUFNLENBQUMsS0FBSyxDQUFDLENBQUMsQ0FBQyxNQUFNLENBQUMsVUFBVSxDQUFDLE1BQU0sQ0FBQyxDQUFDLENBQUMsTUFBTSxDQUFDLFVBQVUsQ0FBQyxRQUFRLENBQUM7SUFDbEYsT0FBTyxJQUFJLE1BQU0sQ0FBQyxpQkFBaUIsQ0FBQyxJQUFJLEVBQUUsSUFBSSxFQUFFLGFBQWEsRUFBRSxXQUFXLENBQUMsQ0FBQztBQUM3RSxDQUFDIn0=