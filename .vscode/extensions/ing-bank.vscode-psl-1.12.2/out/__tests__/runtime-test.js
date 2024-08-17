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
const runtime_1 = require("../src/pslLint/runtime");
const utils = require("./ruleUtils");
describe('Parameter tests', () => {
    let runtimeDiagnostics = [];
    beforeAll(() => __awaiter(void 0, void 0, void 0, function* () {
        runtimeDiagnostics = yield utils.getDiagnostics('ZRuntime.PROC', runtime_1.RuntimeStart.name);
    }));
    test('Diagnostic count', () => {
        expect(runtimeDiagnostics.length).toBe(5);
    });
    test('No diagnostic first start', () => {
        expect(utils.diagnosticsOnLine(9, runtimeDiagnostics)).toMatchObject([]);
    });
    test('One diagnostic second start', () => {
        const reports = utils.diagnosticsOnLine(13, runtimeDiagnostics);
        expect(reports.length).toBe(1);
        expect(reports[0].message).toBe(`Declaration "flagged" referenced inside Runtime.start but not in variable list.`);
    });
    test('One diagnostic third start', () => {
        const reports = utils.diagnosticsOnLine(18, runtimeDiagnostics);
        expect(reports.length).toBe(1);
        expect(reports[0].message).toBe(`Declaration "flagged" referenced inside Runtime.start but not in variable list.`);
    });
    test('No diagnostic with comment', () => {
        expect(utils.diagnosticsOnLine(24, runtimeDiagnostics)).toMatchObject([]);
    });
    test('No diagnostic fifth start', () => {
        expect(utils.diagnosticsOnLine(29, runtimeDiagnostics)).toMatchObject([]);
    });
    test('Method in middle with new variable', () => {
        const reports = utils.diagnosticsOnLine(35, runtimeDiagnostics);
        expect(reports.length).toBe(1);
        expect(reports[0].message).toBe(`Parameter "flaggedParam" referenced inside Runtime.start but not in variable list.`);
    });
    test('Method in middle with new variable', () => {
        const reports = utils.diagnosticsOnLine(42, runtimeDiagnostics);
        const diagnostic = reports[0];
        const relatedArray = diagnostic.relatedInformation;
        const source = relatedArray[0];
        const reference = relatedArray[1];
        expect(reports.length).toBe(1);
        expect(diagnostic.message).toBe(`Declaration "notFlaggedTwice" referenced inside Runtime.start but not in variable list.`);
        expect(relatedArray.length).toBe(2);
        expect(source.message).toBe(`Source of "notFlaggedTwice"`);
        expect(reference.message).toBe(`Reference to "notFlaggedTwice"`);
    });
    test('No literal', () => {
        const reports = utils.diagnosticsOnLine(49, runtimeDiagnostics);
        expect(reports.length).toBe(1);
        expect(reports[0].message).toBe(`Declaration "flagged" referenced inside Runtime.start but not in variable list.`);
    });
});
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoicnVudGltZS10ZXN0LmpzIiwic291cmNlUm9vdCI6IiIsInNvdXJjZXMiOlsiLi4vLi4vX190ZXN0c19fL3J1bnRpbWUtdGVzdC50cyJdLCJuYW1lcyI6W10sIm1hcHBpbmdzIjoiOzs7Ozs7Ozs7OztBQUNBLG9EQUFzRDtBQUN0RCxxQ0FBcUM7QUFFckMsUUFBUSxDQUFDLGlCQUFpQixFQUFFLEdBQUcsRUFBRTtJQUVoQyxJQUFJLGtCQUFrQixHQUFxQixFQUFFLENBQUM7SUFFOUMsU0FBUyxDQUFDLEdBQVMsRUFBRTtRQUNwQixrQkFBa0IsR0FBRyxNQUFNLEtBQUssQ0FBQyxjQUFjLENBQUMsZUFBZSxFQUFFLHNCQUFZLENBQUMsSUFBSSxDQUFDLENBQUM7SUFDckYsQ0FBQyxDQUFBLENBQUMsQ0FBQztJQUVILElBQUksQ0FBQyxrQkFBa0IsRUFBRSxHQUFHLEVBQUU7UUFDN0IsTUFBTSxDQUFDLGtCQUFrQixDQUFDLE1BQU0sQ0FBQyxDQUFDLElBQUksQ0FBQyxDQUFDLENBQUMsQ0FBQztJQUMzQyxDQUFDLENBQUMsQ0FBQztJQUVILElBQUksQ0FBQywyQkFBMkIsRUFBRSxHQUFHLEVBQUU7UUFDdEMsTUFBTSxDQUFDLEtBQUssQ0FBQyxpQkFBaUIsQ0FBQyxDQUFDLEVBQUUsa0JBQWtCLENBQUMsQ0FBQyxDQUFDLGFBQWEsQ0FBQyxFQUFFLENBQUMsQ0FBQztJQUMxRSxDQUFDLENBQUMsQ0FBQztJQUVILElBQUksQ0FBQyw2QkFBNkIsRUFBRSxHQUFHLEVBQUU7UUFDeEMsTUFBTSxPQUFPLEdBQUcsS0FBSyxDQUFDLGlCQUFpQixDQUFDLEVBQUUsRUFBRSxrQkFBa0IsQ0FBQyxDQUFDO1FBQ2hFLE1BQU0sQ0FBQyxPQUFPLENBQUMsTUFBTSxDQUFDLENBQUMsSUFBSSxDQUFDLENBQUMsQ0FBQyxDQUFDO1FBQy9CLE1BQU0sQ0FBQyxPQUFPLENBQUMsQ0FBQyxDQUFDLENBQUMsT0FBTyxDQUFDLENBQUMsSUFBSSxDQUFDLGlGQUFpRixDQUFDLENBQUM7SUFD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