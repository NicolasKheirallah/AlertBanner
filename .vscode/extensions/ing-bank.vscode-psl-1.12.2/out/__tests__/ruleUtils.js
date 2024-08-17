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
exports.getDiagnostics = exports.diagnosticsOnLine = void 0;
const fs = require("fs-extra");
const path = require("path");
const parser_1 = require("../src/parser/parser");
const activate = require("../src/pslLint/activate");
const api = require("../src/pslLint/api");
/**
 * Returns the specific diagnostics on a given line
 *
 * @param lineNumber Zero-based line number to check, i.e. line 1 of the document is lineNumber 0.
 * @param diagnostics The reports to filter
 */
function diagnosticsOnLine(lineNumber, diagnostics) {
    const lineDiagnostics = diagnostics.filter(r => r.range.start.line === lineNumber);
    return lineDiagnostics;
}
exports.diagnosticsOnLine = diagnosticsOnLine;
/**
 * Gets the diagnostics for the given file.
 *
 * @param testFileName The name of the file located in `${PROJECT_ROOT}/__tests__/files/`
 * @param ruleName Optional parameter to return only diagnostics corresponding to the ruleName
 */
function getDiagnostics(testFileName, ruleName) {
    return __awaiter(this, void 0, void 0, function* () {
        const testFilePath = path.resolve('__tests__', 'files', testFileName);
        const text = yield fs.readFile(testFilePath).then(b => b.toString());
        const profileComponent = new api.ProfileComponent(testFilePath, text);
        const diagnostics = activate.getDiagnostics(profileComponent, (0, parser_1.parseText)(text), false);
        if (ruleName)
            return diagnostics.filter(d => d.ruleName === ruleName);
        return diagnostics;
    });
}
exports.getDiagnostics = getDiagnostics;
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoicnVsZVV0aWxzLmpzIiwic291cmNlUm9vdCI6IiIsInNvdXJjZXMiOlsiLi4vLi4vX190ZXN0c19fL3J1bGVVdGlscy50cyJdLCJuYW1lcyI6W10sIm1hcHBpbmdzIjoiOzs7Ozs7Ozs7Ozs7QUFBQSwrQkFBK0I7QUFDL0IsNkJBQTZCO0FBQzdCLGlEQUFpRDtBQUNqRCxvREFBb0Q7QUFDcEQsMENBQTBDO0FBRTFDOzs7OztHQUtHO0FBQ0gsU0FBZ0IsaUJBQWlCLENBQUMsVUFBa0IsRUFBRSxXQUE2QjtJQUNsRixNQUFNLGVBQWUsR0FBRyxXQUFXLENBQUMsTUFBTSxDQUFDLENBQUMsQ0FBQyxFQUFFLENBQUMsQ0FBQyxDQUFDLEtBQUssQ0FBQyxLQUFLLENBQUMsSUFBSSxLQUFLLFVBQVUsQ0FBQyxDQUFDO0lBQ25GLE9BQU8sZUFBZSxDQUFDO0FBQ3hCLENBQUM7QUFIRCw4Q0FHQztBQUVEOzs7OztHQUtHO0FBQ0gsU0FBc0IsY0FBYyxDQUFDLFlBQW9CLEVBQUUsUUFBaUI7O1FBQzNFLE1BQU0sWUFBWSxHQUFHLElBQUksQ0FBQyxPQUFPLENBQUMsV0FBVyxFQUFFLE9BQU8sRUFBRSxZQUFZLENBQUMsQ0FBQztRQUN0RSxNQUFNLElBQUksR0FBRyxNQUFNLEVBQUUsQ0FBQyxRQUFRLENBQUMsWUFBWSxDQUFDLENBQUMsSUFBSSxDQUFDLENBQUMsQ0FBQyxFQUFFLENBQUMsQ0FBQyxDQUFDLFFBQVEsRUFBRSxDQUFDLENBQUM7UUFFckUsTUFBTSxnQkFBZ0IsR0FBRyxJQUFJLEdBQUcsQ0FBQyxnQkFBZ0IsQ0FBQyxZQUFZLEVBQUUsSUFBSSxDQUFDLENBQUM7UUFDdEUsTUFBTSxXQUFXLEdBQUcsUUFBUSxDQUFDLGNBQWMsQ0FBQyxnQkFBZ0IsRUFBRSxJQUFBLGtCQUFTLEVBQUMsSUFBSSxDQUFDLEVBQUUsS0FBSyxDQUFDLENBQUM7UUFDdEYsSUFBSSxRQUFRO1lBQUUsT0FBTyxXQUFXLENBQUMsTUFBTSxDQUFDLENBQUMsQ0FBQyxFQUFFLENBQUMsQ0FBQyxDQUFDLFFBQVEsS0FBSyxRQUFRLENBQUMsQ0FBQztRQUN0RSxPQUFPLFdBQVcsQ0FBQztJQUNwQixDQUFDO0NBQUE7QUFSRCx3Q0FRQyJ9