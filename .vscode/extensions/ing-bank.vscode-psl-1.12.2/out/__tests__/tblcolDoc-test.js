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
const tblcolDoc_1 = require("../src/pslLint/tblcolDoc");
const utils = require("./ruleUtils");
function messageOnLine(lineNumber, allDiagnostics) {
    const diagnosticsOnLine = utils.diagnosticsOnLine(lineNumber, allDiagnostics);
    if (!diagnosticsOnLine.length)
        return '';
    return diagnosticsOnLine[0].message;
}
describe('Table and Column Documentation tests', () => {
    let bracesInColDocDiagnostics = [];
    let withColDocDiagnostics = [];
    let withoutColDocDiagnostics = [];
    let withSpaceColDocDiagnostics = [];
    let bracesInsideColDefDiagnostics = [];
    let withoutClosedBracesColDiagnostics = [];
    let bracesInTblDocDiagnostics = [];
    let withTblDocDiagnostics = [];
    let withoutTblDocDiagnostics = [];
    let withSpaceTblDocDiagnostics = [];
    let bracesInsideTblDefDiagnostics = [];
    let withoutClosedBracesTblDiagnostics = [];
    beforeAll(() => __awaiter(void 0, void 0, void 0, function* () {
        bracesInColDocDiagnostics = yield utils.getDiagnostics('ZTblColDocTst-Col1.COL', tblcolDoc_1.TblColDocumentation.name);
        withColDocDiagnostics = yield utils.getDiagnostics('ZTblColDocTst-Col2.COL', tblcolDoc_1.TblColDocumentation.name);
        withoutColDocDiagnostics = yield utils.getDiagnostics('ZTblColDocTst-Col3.COL', tblcolDoc_1.TblColDocumentation.name);
        withSpaceColDocDiagnostics = yield utils.getDiagnostics('ZTblColDocTst-Col4.COL', tblcolDoc_1.TblColDocumentation.name);
        bracesInsideColDefDiagnostics = yield utils.getDiagnostics('ZTblColDocTst-Col5.COL', tblcolDoc_1.TblColDocumentation.name);
        withoutClosedBracesColDiagnostics = yield utils.getDiagnostics('ZTblColDocTst-Col6.COL', tblcolDoc_1.TblColDocumentation.name);
        bracesInTblDocDiagnostics = yield utils.getDiagnostics('ZTblColDocTst1.TBL', tblcolDoc_1.TblColDocumentation.name);
        withTblDocDiagnostics = yield utils.getDiagnostics('ZTblColDocTst2.TBL', tblcolDoc_1.TblColDocumentation.name);
        withoutTblDocDiagnostics = yield utils.getDiagnostics('ZTblColDocTst3.TBL', tblcolDoc_1.TblColDocumentation.name);
        withSpaceTblDocDiagnostics = yield utils.getDiagnostics('ZTblColDocTst4.TBL', tblcolDoc_1.TblColDocumentation.name);
        bracesInsideTblDefDiagnostics = yield utils.getDiagnostics('ZTblColDocTst5.TBL', tblcolDoc_1.TblColDocumentation.name);
        withoutClosedBracesTblDiagnostics = yield utils.getDiagnostics('ZTblColDocTst6.TBL', tblcolDoc_1.TblColDocumentation.name);
    }));
    test('Column documentation', () => {
        // Column documentation exists with '{' '}' braces
        expect(bracesInColDocDiagnostics.length).toBe(0);
        // Column documentation exists
        expect(withColDocDiagnostics.length).toBe(0);
        // Without } in the column definition.This should be ignored as the compiler should handle it
        expect(withoutClosedBracesColDiagnostics.length).toBe(0);
        // Without Column documentation and '{' '}' inside the definition
        expect(bracesInsideColDefDiagnostics.length).toBe(0);
        // Without Column documentation
        expect(messageOnLine(36, withoutColDocDiagnostics))
            .toBe(`Documentation missing for data item "ZTblColDocTst-Col3.COL".`);
        // Without Column documentation but only space exists after '}' braces
        expect(messageOnLine(36, withSpaceColDocDiagnostics))
            .toBe(`Documentation missing for data item "ZTblColDocTst-Col4.COL".`);
    });
    test('Table documentation', () => {
        // Table documentation exists with '{' '}' braces
        expect(bracesInTblDocDiagnostics.length).toBe(0);
        // Table documentation exists
        expect(withTblDocDiagnostics.length).toBe(0);
        // Without } in the Table definition.This should be ignored as the compiler should handle it
        expect(withoutClosedBracesTblDiagnostics.length).toBe(0);
        // Without Table documentation and '{' '}' inside the definition
        expect(bracesInsideTblDefDiagnostics.length).toBe(0);
        // Without Table documentation
        expect(messageOnLine(41, withoutTblDocDiagnostics))
            .toBe(`Documentation missing for table definition "ZTblColDocTst3.TBL".`);
        // Without Table documentation but only space exists after '}' braces
        expect(messageOnLine(41, withSpaceTblDocDiagnostics))
            .toBe(`Documentation missing for table definition "ZTblColDocTst4.TBL".`);
    });
});
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoidGJsY29sRG9jLXRlc3QuanMiLCJzb3VyY2VSb290IjoiIiwic291cmNlcyI6WyIuLi8uLi9fX3Rlc3RzX18vdGJsY29sRG9jLXRlc3QudHMiXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6Ijs7Ozs7Ozs7Ozs7QUFDQSx3REFBK0Q7QUFDL0QscUNBQXFDO0FBRXJDLFNBQVMsYUFBYSxDQUFDLFVBQWtCLEVBQUUsY0FBZ0M7SUFDMUUsTUFBTSxpQkFBaUIsR0FBRyxLQUFLLENBQUMsaUJBQWlCLENBQUMsVUFBVSxFQUFFLGNBQWMsQ0FBQyxDQUFDO0lBQzlFLElBQUksQ0FBQyxpQkFBaUIsQ0FBQyxNQUFNO1FBQUUsT0FBTyxFQUFFLENBQUM7SUFDekMsT0FBTyxpQkFBaUIsQ0FBQyxDQUFDLENBQUMsQ0FBQyxPQUFPLENBQUM7QUFDckMsQ0FBQztBQUVELFFBQVEsQ0FBQyxzQ0FBc0MsRUFBRSxHQUFHLEVBQUU7SUFFckQsSUFBSSx5QkFBeUIsR0FBcUIsRUFBRSxDQUFDO0lBQ3JELElBQUkscUJBQXFCLEdBQXFCLEVBQUUsQ0FBQztJQUNqRCxJQUFJLHdCQUF3QixHQUFxQixFQUFFLENBQUM7SUFDcEQsSUFBSSwwQkFBMEIsR0FBcUIsRUFBRSxDQUFDO0lBQ3RELElBQUksNkJBQTZCLEdBQXFCLEVBQUUsQ0FBQztJQUN6RCxJQUFJLGlDQUFpQyxHQUFxQixFQUFFLENBQUM7SUFFN0QsSUFBSSx5QkFBeUIsR0FBcUIsRUFBRSxDQUFDO0lBQ3JELElBQUkscUJBQXFCLEdBQXFCLEVBQUUsQ0FBQztJQUNqRCxJQUFJLHdCQUF3QixHQUFxQixFQUFFLENBQUM7SUFDcEQsSUFBSSwwQkFBMEIsR0FBcUIsRUFBRSxDQUFDO0lBQ3RELElBQUksNkJBQTZCLEdBQXFCLEVBQUUsQ0FBQztJQUN6RCxJQUFJLGlDQUFpQyxHQUFxQixFQUFFLENBQUM7SUFFN0QsU0FBUyxDQUFDLEdBQVMsRUFBRTtRQUNwQix5QkFBeUIsR0FBRyxNQUFNLEtBQUssQ0FBQyxjQUFjLENBQUMsd0JB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