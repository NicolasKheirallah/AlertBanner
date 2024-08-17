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
const methodDoc_1 = require("../src/pslLint/methodDoc");
const utils = require("./ruleUtils");
function messageOnLine(lineNumber, allDiagnostics) {
    const diagnosticsOnLine = utils.diagnosticsOnLine(lineNumber, allDiagnostics);
    if (!diagnosticsOnLine.length)
        return '';
    return diagnosticsOnLine[0].message;
}
describe('Parameter tests', () => {
    const procName = 'ZMethodDoc.PROC';
    let docDiagnostics = [];
    let emptyLineDiagnostics = [];
    let separatorDiagnostics = [];
    beforeAll(() => __awaiter(void 0, void 0, void 0, function* () {
        docDiagnostics = yield utils.getDiagnostics(procName, methodDoc_1.MethodDocumentation.name);
        emptyLineDiagnostics = yield utils.getDiagnostics(procName, methodDoc_1.TwoEmptyLines.name);
        separatorDiagnostics = yield utils.getDiagnostics(procName, methodDoc_1.MethodSeparator.name);
    }));
    test('allProblemsNoLineAbove', () => {
        expect(messageOnLine(0, docDiagnostics)).toBe(`Documentation missing for label "allProblemsNoLineAbove".`);
        expect(messageOnLine(0, separatorDiagnostics)).toBe(`Separator missing for label "allProblemsNoLineAbove".`);
        expect(messageOnLine(0, emptyLineDiagnostics)).toBe(`There should be two empty lines above label "allProblemsNoLineAbove".`);
    });
    test('allProblems', () => {
        expect(messageOnLine(2, docDiagnostics)).toBe(`Documentation missing for label "allProblems".`);
        expect(messageOnLine(2, separatorDiagnostics)).toBe(`Separator missing for label "allProblems".`);
        expect(messageOnLine(2, emptyLineDiagnostics)).toBe(`There should be two empty lines above label "allProblems".`);
    });
    test('onlySeparator', () => {
        expect(messageOnLine(5, docDiagnostics)).toBe(`Documentation missing for label "onlySeparator".`);
        expect(messageOnLine(5, separatorDiagnostics)).toBe('');
        expect(messageOnLine(5, emptyLineDiagnostics)).toBe(`There should be two empty lines above label "onlySeparator".`);
    });
    test('twoLineSeparator', () => {
        expect(messageOnLine(14, docDiagnostics)).toBe(`Documentation missing for label "twoLineSeparator".`);
        expect(messageOnLine(14, separatorDiagnostics)).toBe('');
        expect(messageOnLine(14, emptyLineDiagnostics)).toBe('');
    });
    test('onlyDoc', () => {
        expect(messageOnLine(16, docDiagnostics)).toBe('');
        expect(messageOnLine(16, separatorDiagnostics)).toBe(`Separator missing for label "onlyDoc".`);
        expect(messageOnLine(16, emptyLineDiagnostics)).toBe(`There should be two empty lines above label "onlyDoc".`);
    });
    test('oneLineDoc', () => {
        expect(messageOnLine(20, docDiagnostics)).toBe('');
        expect(messageOnLine(20, separatorDiagnostics)).toBe(`Separator missing for label "oneLineDoc".`);
        expect(messageOnLine(20, emptyLineDiagnostics)).toBe(`There should be two empty lines above label "oneLineDoc".`);
    });
    test('twoLineDoc', () => {
        expect(messageOnLine(25, docDiagnostics)).toBe('');
        expect(messageOnLine(25, separatorDiagnostics)).toBe(`Separator missing for label "twoLineDoc".`);
        expect(messageOnLine(25, emptyLineDiagnostics)).toBe('');
    });
    test('withEverything', () => {
        expect(messageOnLine(31, docDiagnostics)).toBe('');
        expect(messageOnLine(31, separatorDiagnostics)).toBe('');
        expect(messageOnLine(31, emptyLineDiagnostics)).toBe('');
    });
});
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoibWV0aG9kRG9jLXRlc3QuanMiLCJzb3VyY2VSb290IjoiIiwic291cmNlcyI6WyIuLi8uLi9fX3Rlc3RzX18vbWV0aG9kRG9jLXRlc3QudHMiXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6Ijs7Ozs7Ozs7Ozs7QUFDQSx3REFBK0Y7QUFDL0YscUNBQXFDO0FBRXJDLFNBQVMsYUFBYSxDQUFDLFVBQWtCLEVBQUUsY0FBZ0M7SUFDMUUsTUFBTSxpQkFBaUIsR0FBRyxLQUFLLENBQUMsaUJBQWlCLENBQUMsVUFBVSxFQUFFLGNBQWMsQ0FBQyxDQUFDO0lBQzlFLElBQUksQ0FBQyxpQkFBaUIsQ0FBQyxNQUFNO1FBQUUsT0FBTyxFQUFFLENBQUM7SUFDekMsT0FBTyxpQkFBaUIsQ0FBQyxDQUFDLENBQUMsQ0FBQyxPQUFPLENBQUM7QUFDckMsQ0FBQztBQUVELFFBQVEsQ0FBQyxpQkFBaUIsRUFBRSxHQUFHLEVBQUU7SUFDaEMsTUFBTSxRQUFRLEdBQUcsaUJBQWlCLENBQUM7SUFDbkMsSUFBSSxjQUFjLEdBQXFCLEVBQUUsQ0FBQztJQUMxQyxJQUFJLG9CQUFvQixHQUFxQixFQUFFLENBQUM7SUFDaEQsSUFBSSxvQkFBb0IsR0FBcUIsRUFBRSxDQUFDO0lBRWhELFNBQVMsQ0FBQyxHQUFTLEVBQUU7UUFDcEIsY0FBYyxHQUFHLE1BQU0sS0FBSyxDQUFDLGNBQWMsQ0FBQyxRQUFRLEVBQUUsK0JBQW1CLENBQUMsSUFBSSxDQUFDLENBQUM7UUFDaEYsb0JBQW9CLEdBQUcsTUFBTSxLQUFLLENBQUMsY0FBYyxDQUFDLFFBQVEsRUFBRSx5QkFBYSxDQUFDLElBQUksQ0FBQyxDQUFDO1FBQ2hGLG9CQUFvQixHQUFHLE1BQU0sS0FBSyxDQUFDLGNBQWMsQ0FBQyxRQUFRLEVBQUUsMkJBQWUsQ0FBQyxJQUFJLENBQUMsQ0FBQztJQUNuRixDQUFDLENBQUEsQ0FBQyxDQUFDO0lBRUgsSUFBSSxDQUFDLHdCQUF3QixFQUFFLEdBQUcsRUFBRTtRQUNuQyxNQUFNLENBQUMsYUFBYSxD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