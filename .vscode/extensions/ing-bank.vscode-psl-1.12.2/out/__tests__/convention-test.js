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
const api = require("../src/pslLint/api");
const elementsConventionChecker_1 = require("../src/pslLint/elementsConventionChecker");
const utils = require("./ruleUtils");
describe('Members tests', () => {
    let literalDiagnostics = [];
    let camelCaseDiagnostics = [];
    let lengthDiagnostics = [];
    let vDiagnostics = [];
    let withoutDummyDiagnostics = [];
    let withDummyDiagnostics = [];
    beforeAll(() => __awaiter(void 0, void 0, void 0, function* () {
        literalDiagnostics = yield utils.getDiagnostics('ZTestConvention.PROC', elementsConventionChecker_1.MemberLiteralCase.name);
        camelCaseDiagnostics = yield utils.getDiagnostics('ZTestConvention.PROC', elementsConventionChecker_1.MemberCamelCase.name);
        lengthDiagnostics = yield utils.getDiagnostics('ZTestConvention.PROC', elementsConventionChecker_1.MemberLength.name);
        vDiagnostics = yield utils.getDiagnostics('ZTestConvention.PROC', elementsConventionChecker_1.MemberStartsWithV.name);
        withoutDummyDiagnostics = yield utils.getDiagnostics('ZTestConvention.PROC', elementsConventionChecker_1.PropertyIsDummy.name);
        withDummyDiagnostics = yield utils.getDiagnostics('ZParent.PROC', elementsConventionChecker_1.PropertyIsDummy.name);
    }));
    test('Upper case literal report', () => {
        expect(utils.diagnosticsOnLine(5, literalDiagnostics).length).toBe(1);
    });
    test('Camel case literal report', () => {
        expect(utils.diagnosticsOnLine(4, camelCaseDiagnostics).length).toBe(1);
    });
    test('More than 25 characters', () => {
        expect(utils.diagnosticsOnLine(14, lengthDiagnostics).length).toBe(1);
    });
    test('Starts with v', () => {
        const diagnosticsOnLine = utils.diagnosticsOnLine(23, vDiagnostics);
        expect(diagnosticsOnLine.length).toBe(1);
        expect(diagnosticsOnLine[0].message).toBe(`Declaration "vString" starts with 'v'.`);
        expect(diagnosticsOnLine[0].severity).toBe(api.DiagnosticSeverity.Warning);
    });
    test('Public starts with v', () => {
        const diagnosticsOnLine = utils.diagnosticsOnLine(24, vDiagnostics);
        expect(diagnosticsOnLine.length).toBe(1);
        expect(diagnosticsOnLine[0].message).toBe(`Declaration "vNumber" is public and starts with 'v'.`);
        expect(diagnosticsOnLine[0].severity).toBe(api.DiagnosticSeverity.Information);
    });
    test('Property was not called \'dummy\'', () => {
        expect(withoutDummyDiagnostics.length).toBe(0);
    });
    test('Property was called \'dummy\'', () => {
        expect(utils.diagnosticsOnLine(2, withDummyDiagnostics).length).toBe(1);
    });
});
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoiY29udmVudGlvbi10ZXN0LmpzIiwic291cmNlUm9vdCI6IiIsInNvdXJjZXMiOlsiLi4vLi4vX190ZXN0c19fL2NvbnZlbnRpb24tdGVzdC50cyJdLCJuYW1lcyI6W10sIm1hcHBpbmdzIjoiOzs7Ozs7Ozs7OztBQUFBLDBDQUEwQztBQUMxQyx3RkFFeUc7QUFDekcscUNBQXFDO0FBRXJDLFFBQVEsQ0FBQyxlQUFlLEVBQUUsR0FBRyxFQUFFO0lBQzlCLElBQUksa0JBQWtCLEdBQXFCLEVBQUUsQ0FBQztJQUM5QyxJQUFJLG9CQUFvQixHQUFxQixFQUFFLENBQUM7SUFDaEQsSUFBSSxpQkFBaUIsR0FBcUIsRUFBRSxDQUFDO0lBQzdDLElBQUksWUFBWSxHQUFxQixFQUFFLENBQUM7SUFDeEMsSUFBSSx1QkFBdUIsR0FBcUIsRUFBRSxDQUFDO0lBQ25ELElBQUksb0JBQW9CLEdBQXFCLEVBQUUsQ0FBQztJQUVoRCxTQUFTLENBQUMsR0FBUyxFQUFFO1FBQ3BCLGtCQUFrQixHQUFHLE1BQU0sS0FBSyxDQUFDLGNBQWMsQ0FBQyxzQkFBc0IsRUFBRSw2Q0FBaUIsQ0FBQyxJQUFJLENBQUMsQ0FBQztRQUNoRyxvQkFBb0IsR0FBRyxNQUFNLEtBQUssQ0FBQyxjQUFjLENBQUMsc0JBQXNCLEVBQUUsMkNBQWUsQ0FBQyxJQUFJLENBQUMsQ0FBQztRQUNoRyxpQkFBaUIsR0FBRyxNQUFNLEtBQUssQ0FBQyxjQUFjLENBQUMsc0JBQXNCLEVBQUUsd0NBQVksQ0FBQyxJQUFJLENBQUMsQ0FBQztRQUMxRixZQUFZLEdBQUcsTUFBTSxLQUFLLENBQUMsY0FBYyxDQUFDLHNCQUFzQixFQUFFLDZDQUFpQixDQUFDLElBQUksQ0FBQyxDQUFDO1FBQzFGLHVCQUF1QixHQUFHLE1BQU0sS0FBSyxDQUFDLGNBQWMsQ0FBQyxzQkFBc0IsRUFBRSwyQ0FBZSxDQUFDLElBQUksQ0FBQyxDQUFDO1FBQ25HLG9CQUFvQixHQUFHLE1BQU0sS0FBSyxDQUFDLGNBQWMsQ0FBQyxjQUFjLEVBQUUs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