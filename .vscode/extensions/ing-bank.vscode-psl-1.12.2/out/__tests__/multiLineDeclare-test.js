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
const multiLineDeclare_1 = require("../src/pslLint/multiLineDeclare");
const utils = require("./ruleUtils");
describe('Parameter tests', () => {
    let multiLineDiagnostics = [];
    beforeAll(() => __awaiter(void 0, void 0, void 0, function* () {
        multiLineDiagnostics = yield utils.getDiagnostics('ZMultiLineDeclare.PROC', multiLineDeclare_1.MultiLineDeclare.name);
    }));
    test('line 2', () => {
        const test1Message = 'Declaration test1 should be initialized on a new line.';
        const test2Message = 'Declaration test2 should be initialized on a new line.';
        const diagnostics = utils.diagnosticsOnLine(1, multiLineDiagnostics);
        expect(diagnostics.length).toBe(2);
        expect(diagnostics[0].message).toBe(test1Message);
        expect(diagnostics[1].message).toBe(test2Message);
    });
    test('line 3', () => {
        const number1Message = 'Declaration number1 should be initialized on a new line.';
        const number2Message = 'Declaration number2 should be initialized on a new line.';
        const diagnostics = utils.diagnosticsOnLine(2, multiLineDiagnostics);
        expect(diagnostics.length).toBe(2);
        expect(diagnostics[0].message).toBe(number1Message);
        expect(diagnostics[1].message).toBe(number2Message);
    });
    test('line 4', () => {
        const rs1Message = 'Declaration rs1 should be initialized on a new line.';
        const rs2Message = 'Declaration rs2 should be initialized on a new line.';
        const rs3Message = 'Declaration rs3 should be initialized on a new line.';
        const diagnostics = utils.diagnosticsOnLine(3, multiLineDiagnostics);
        expect(diagnostics.length).toBe(3);
        expect(diagnostics[0].message).toBe(rs1Message);
        expect(diagnostics[1].message).toBe(rs2Message);
        expect(diagnostics[2].message).toBe(rs3Message);
    });
    test('line 5', () => {
        const diagnostics = utils.diagnosticsOnLine(4, multiLineDiagnostics);
        expect(diagnostics.length).toBe(0);
    });
    test('line 6', () => {
        const number11Message = 'Declaration number11 should be initialized on a new line.';
        const number13Message = 'Declaration number13 should be initialized on a new line.';
        const number15Message = 'Declaration number15 should be initialized on a new line.';
        const diagnostics = utils.diagnosticsOnLine(5, multiLineDiagnostics);
        expect(diagnostics.length).toBe(3);
        expect(diagnostics[0].message).toBe(number11Message);
        expect(diagnostics[1].message).toBe(number13Message);
        expect(diagnostics[2].message).toBe(number15Message);
    });
    test('line 7', () => {
        const number19Message = 'Declaration number19 should be initialized on a new line.';
        const number20Message = 'Declaration number20 should be initialized on a new line.';
        const number21Message = 'Declaration number21 should be initialized on a new line.';
        const number22Message = 'Declaration number22 should be initialized on a new line.';
        const number23Message = 'Declaration number23 should be initialized on a new line.';
        const diagnostics = utils.diagnosticsOnLine(6, multiLineDiagnostics);
        expect(diagnostics.length).toBe(5);
        expect(diagnostics[0].message).toBe(number19Message);
        expect(diagnostics[1].message).toBe(number20Message);
        expect(diagnostics[2].message).toBe(number21Message);
        expect(diagnostics[3].message).toBe(number22Message);
        expect(diagnostics[4].message).toBe(number23Message);
    });
    test('line 8', () => {
        const diagnostics = utils.diagnosticsOnLine(7, multiLineDiagnostics);
        expect(diagnostics.length).toBe(0);
    });
    test('line 9', () => {
        const number32Message = 'Declaration number32 should be initialized on a new line.';
        const number34Message = 'Declaration number34 should be initialized on a new line.';
        const number36Message = 'Declaration number36 should be initialized on a new line.';
        const diagnostics = utils.diagnosticsOnLine(8, multiLineDiagnostics);
        expect(diagnostics.length).toBe(3);
        expect(diagnostics[0].message).toBe(number32Message);
        expect(diagnostics[1].message).toBe(number34Message);
        expect(diagnostics[2].message).toBe(number36Message);
    });
    test('line 10', () => {
        const number37Message = 'Declaration number37 should be initialized on a new line.';
        const number38Message = 'Declaration number38 should be initialized on a new line.';
        const diagnostics = utils.diagnosticsOnLine(9, multiLineDiagnostics);
        expect(diagnostics.length).toBe(2);
        expect(diagnostics[0].message).toBe(number37Message);
        expect(diagnostics[1].message).toBe(number38Message);
    });
    test('line 11', () => {
        const rs4Message = 'Declaration rs4 should be initialized on a new line.';
        const rs5Message = 'Declaration rs5 should be initialized on a new line.';
        const diagnostics = utils.diagnosticsOnLine(10, multiLineDiagnostics);
        expect(diagnostics.length).toBe(2);
        expect(diagnostics[0].message).toBe(rs4Message);
        expect(diagnostics[1].message).toBe(rs5Message);
    });
    test('line 12', () => {
        const number38Message = 'Declaration number38 should be initialized on a new line.';
        const number39Message = 'Declaration number39 should be initialized on a new line.';
        const diagnostics = utils.diagnosticsOnLine(11, multiLineDiagnostics);
        expect(diagnostics.length).toBe(2);
        expect(diagnostics[0].message).toBe(number38Message);
        expect(diagnostics[1].message).toBe(number39Message);
    });
    test('line 13', () => {
        const number40Message = 'Declaration number40 should be initialized on a new line.';
        const number41Message = 'Declaration number41 should be initialized on a new line.';
        const diagnostics = utils.diagnosticsOnLine(12, multiLineDiagnostics);
        expect(diagnostics.length).toBe(2);
        expect(diagnostics[0].message).toBe(number40Message);
        expect(diagnostics[1].message).toBe(number41Message);
    });
    test('line 14', () => {
        const number42Message = 'Declaration number42 should be initialized on a new line.';
        const number41Message = 'Declaration number41 should be initialized on a new line.';
        const number43Message = 'Declaration number43 should be initialized on a new line.';
        const diagnostics = utils.diagnosticsOnLine(13, multiLineDiagnostics);
        expect(diagnostics.length).toBe(3);
        expect(diagnostics[0].message).toBe(number42Message);
        expect(diagnostics[1].message).toBe(number41Message);
        expect(diagnostics[2].message).toBe(number43Message);
    });
});
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoibXVsdGlMaW5lRGVjbGFyZS10ZXN0LmpzIiwic291cmNlUm9vdCI6IiIsInNvdXJjZXMiOlsiLi4vLi4vX190ZXN0c19fL211bHRpTGluZURlY2xhcmUtdGVzdC50cyJdLCJuYW1lcyI6W10sIm1hcHBpbmdzIjoiOzs7Ozs7Ozs7OztBQUNBLHNFQUFtRTtBQUNuRSxxQ0FBcUM7QUFFckMsUUFBUSxDQUFDLGlCQUFpQixFQUFFLEdBQUcsRUFBRTtJQUVoQyxJQUFJLG9CQUFvQixHQUFxQixFQUFFLENBQUM7SUFFaEQsU0FBUyxDQUFDLEdBQVMsRUFBRTtRQUNwQixvQkFBb0IsR0FBRyxNQUFNLEtBQUssQ0FBQyxjQUFjLENBQUMsd0JBQXdCLEVBQUUsbUNBQWdCLENBQUMsSUFBSSxDQUFDLENBQUM7SUFDcEcsQ0FBQyxDQUFBLENBQUMsQ0FBQztJQUVILElBQUksQ0FBQyxRQUFRLEVBQUUsR0FBRyxFQUFFO1FBQ25CLE1BQU0sWUFBWSxHQUFHLHdEQUF3RCxDQUFDO1FBQzlFLE1BQU0sWUFBWSxHQUFHLHdEQUF3RCxDQUFDO1FBQzlFLE1BQU0sV0FBVyxHQUFHLEtBQUssQ0FBQyxpQkFBaUIsQ0FBQyxDQUFDLEVBQUUsb0JBQW9CLENBQUMsQ0FBQztRQUNyRSxNQUFNLENBQUMsV0FBVyxDQUFDLE1BQU0sQ0FBQyxDQUFDLElBQUksQ0FBQyxDQUFDLENBQUMsQ0FBQztRQUNuQyxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxZQUFZLENBQUMsQ0FBQztRQUNsRCxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxZQUFZLENBQUMsQ0FBQztJQUNuRCxDQUFDLENBQUMsQ0FBQztJQUVILElBQUksQ0FBQyxRQUFRLEVBQUUsR0FBRyxFQUFFO1FBQ25CLE1BQU0sY0FBYyxHQUFHLDBEQUEwRCxDQUFDO1FBQ2xGLE1BQU0sY0FBYyxHQUFHLDBEQUEwRCxDQUFDO1FBQ2xGLE1BQU0sV0FBVyxHQUFHLEtBQUssQ0FBQyxpQkFBaUIsQ0FBQyxDQUFDLEVBQUUsb0JBQW9CLENBQUMsQ0FBQztRQUNyRSxNQUFNLENBQUMsV0FBVyxDQUFDLE1BQU0sQ0FBQyxDQUFDLElBQUksQ0FBQyxDQUFDLENBQUMsQ0FBQztRQUNuQyxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxjQUFjLENBQUMsQ0FBQztRQUNwRCxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxjQUFjLENBQUMsQ0FBQztJQUNyRCxDQUFDLENBQUMsQ0FBQztJQUVILElBQUksQ0FBQyxRQUFRLEVBQUUsR0FBRyxFQUFFO1FBQ25CLE1BQU0sVUFBVSxHQUFHLHNEQUFzRCxDQUFDO1FBQzFFLE1BQU0sVUFBVSxHQUFHLHNEQUFzRCxDQUFDO1FBQzFFLE1BQU0sVUFBVSxHQUFHLHNEQUFzRCxDQUFDO1FBQzFFLE1BQU0sV0FBVyxHQUFHLEtBQUssQ0FBQyxpQkFBaUIsQ0FBQyxDQUFDLEVBQUUsb0JBQW9CLENBQUMsQ0FBQztRQUNyRSxNQUFNLENBQUMsV0FBVyxDQUFDLE1BQU0sQ0FBQyxDQUFDLElBQUksQ0FBQyxDQUFDLENBQUMsQ0FBQztRQUNuQyxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxVQUFVLENBQUMsQ0FBQztRQUNoRCxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxVQUFVLENBQUMsQ0FBQztRQUNoRCxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxVQUFVLENBQUMsQ0FBQztJQUNqRCxDQUFDLENBQUMsQ0FBQztJQUVILElBQUksQ0FBQyxRQUFRLEVBQUUsR0FBRyxFQUFFO1FBQ25CLE1BQU0sV0FBVyxHQUFHLEtBQUssQ0FBQyxpQkFBaUIsQ0FBQyxDQUFDLEVBQUUsb0JBQW9CLENBQUMsQ0FBQztRQUNyRSxNQUFNLENBQUMsV0FBVyxDQUFDLE1BQU0sQ0FBQyxDQUFDLElBQUksQ0FBQyxDQUFDLENBQUMsQ0FBQztJQUNwQyxDQUFDLENBQUMsQ0FBQztJQUVILElBQUksQ0FBQyxRQUFRLEVBQUUsR0FBRyxFQUFFO1FBQ25CLE1BQU0sZUFBZSxHQUFHLDJEQUEyRCxDQUFDO1FBQ3BGLE1BQU0sZUFBZSxHQUFHLDJEQUEyRCxDQUFDO1FBQ3BGLE1BQU0sZUFBZSxHQUFHLDJEQUEyRCxDQUFDO1FBQ3BGLE1BQU0sV0FBVyxHQUFHLEtBQUssQ0FBQyxpQkFBaUIsQ0FBQyxDQUFDLEVBQUUsb0JBQW9CLENBQUMsQ0FBQztRQUNyRSxNQUFNLENBQUMsV0FBVyxDQUFDLE1BQU0sQ0FBQyxDQUFDLElBQUksQ0FBQyxDQUFDLENBQUMsQ0FBQztRQUNuQyxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxlQUFlLENBQUMsQ0FBQztRQUNyRCxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxlQUFlLENBQUMsQ0FBQztRQUNyRCxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxlQUFlLENBQUMsQ0FBQztJQUN0RCxDQUFDLENBQUMsQ0FBQztJQUVILElBQUksQ0FBQyxRQUFRLEVBQUUsR0FBRyxFQUFFO1FBQ25CLE1BQU0sZUFBZSxHQUFHLDJEQUEyRCxDQUFDO1FBQ3BGLE1BQU0sZUFBZSxHQUFHLDJEQUEyRCxDQUFDO1FBQ3BGLE1BQU0sZUFBZSxHQUFHLDJEQUEyRCxDQUFDO1FBQ3BGLE1BQU0sZUFBZSxHQUFHLDJEQUEyRCxDQUFDO1FBQ3BGLE1BQU0sZUFBZSxHQUFHLDJEQUEyRCxDQUFDO1FBQ3BGLE1BQU0sV0FBVyxHQUFHLEtBQUssQ0FBQyxpQkFBaUIsQ0FBQyxDQUFDLEVBQUUsb0JBQW9CLENBQUMsQ0FBQztRQUNyRSxNQUFNLENBQUMsV0FBVyxDQUFDLE1BQU0sQ0FBQyxDQUFDLElBQUksQ0FBQyxDQUFDLENBQUMsQ0FBQztRQUNuQyxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxlQUFlLENBQUMsQ0FBQztRQUNyRCxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxlQUFlLENBQUMsQ0FBQztRQUNyRCxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxlQUFlLENBQUMsQ0FBQztRQUNyRCxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxlQUFlLENBQUMsQ0FBQztRQUNyRCxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxlQUFlLENBQUMsQ0FBQztJQUN0RCxDQUFDLENBQUMsQ0FBQztJQUVILElBQUksQ0FBQyxRQUFRLEVBQUUsR0FBRyxFQUFFO1FBQ25CLE1BQU0sV0FBVyxHQUFHLEtBQUssQ0FBQyxpQkFBaUIsQ0FBQyxDQUFDLEVBQUUsb0JBQW9CLENBQUMsQ0FBQztRQUNyRSxNQUFNLENBQUMsV0FBVyxDQUFDLE1BQU0sQ0FBQyxDQUFDLElBQUksQ0FBQyxDQUFDLENBQUMsQ0FBQztJQUNwQyxDQUFDLENBQUMsQ0FBQztJQUVILElBQUksQ0FBQyxRQUFRLEVBQUUsR0FBRyxFQUFFO1FBQ25CLE1BQU0sZUFBZSxHQUFHLDJEQUEyRCxDQUFDO1FBQ3BGLE1BQU0sZUFBZSxHQUFHLDJEQUEyRCxDQUFDO1FBQ3BGLE1BQU0sZUFBZSxHQUFHLDJEQUEyRCxDQUFDO1FBQ3BGLE1BQU0sV0FBVyxHQUFHLEtBQUssQ0FBQyxpQkFBaUIsQ0FBQyxDQUFDLEVBQUUsb0JBQW9CLENBQUMsQ0FBQztRQUNyRSxNQUFNLENBQUMsV0FBVyxDQUFDLE1BQU0sQ0FBQyxDQUFDLElBQUksQ0FBQyxDQUFDLENBQUMsQ0FBQztRQUNuQyxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxlQUFlLENBQUMsQ0FBQztRQUNyRCxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxlQUFlLENBQUMsQ0FBQztRQUNyRCxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxlQUFlLENBQUMsQ0FBQztJQUN0RCxDQUFDLENBQUMsQ0FBQztJQUVILElBQUksQ0FBQyxTQUFTLEVBQUUsR0FBRyxFQUFFO1FBQ3BCLE1BQU0sZUFBZSxHQUFHLDJEQUEyRCxDQUFDO1FBQ3BGLE1BQU0sZUFBZSxHQUFHLDJEQUEyRCxDQUFDO1FBQ3BGLE1BQU0sV0FBVyxHQUFHLEtBQUssQ0FBQyxpQkFBaUIsQ0FBQyxDQUFDLEVBQUUsb0JBQW9CLENBQUMsQ0FBQztRQUNyRSxNQUFNLENBQUMsV0FBVyxDQUFDLE1BQU0sQ0FBQyxDQUFDLElBQUksQ0FBQyxDQUFDLENBQUMsQ0FBQztRQUNuQyxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxlQUFlLENBQUMsQ0FBQztRQUNyRCxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxlQUFlLENBQUMsQ0FBQztJQUN0RCxDQUFDLENBQUMsQ0FBQztJQUVILElBQUksQ0FBQyxTQUFTLEVBQUUsR0FBRyxFQUFFO1FBQ3BCLE1BQU0sVUFBVSxHQUFHLHNEQUFzRCxDQUFDO1FBQzFFLE1BQU0sVUFBVSxHQUFHLHNEQUFzRCxDQUFDO1FBQzFFLE1BQU0sV0FBVyxHQUFHLEtBQUssQ0FBQyxpQkFBaUIsQ0FBQyxFQUFFLEVBQUUsb0JBQW9CLENBQUMsQ0FBQztRQUN0RSxNQUFNLENBQUMsV0FBVyxDQUFDLE1BQU0sQ0FBQyxDQUFDLElBQUksQ0FBQyxDQUFDLENBQUMsQ0FBQztRQUNuQyxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxVQUFVLENBQUMsQ0FBQztRQUNoRCxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxVQUFVLENBQUMsQ0FBQztJQUNqRCxDQUFDLENBQUMsQ0FBQztJQUVILElBQUksQ0FBQyxTQUFTLEVBQUUsR0FBRyxFQUFFO1FBQ3BCLE1BQU0sZUFBZSxHQUFHLDJEQUEyRCxDQUFDO1FBQ3BGLE1BQU0sZUFBZSxHQUFHLDJEQUEyRCxDQUFDO1FBQ3BGLE1BQU0sV0FBVyxHQUFHLEtBQUssQ0FBQyxpQkFBaUIsQ0FBQyxFQUFFLEVBQUUsb0JBQW9CLENBQUMsQ0FBQztRQUN0RSxNQUFNLENBQUMsV0FBVyxDQUFDLE1BQU0sQ0FBQyxDQUFDLElBQUksQ0FBQyxDQUFDLENBQUMsQ0FBQztRQUNuQyxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxlQUFlLENBQUMsQ0FBQztRQUNyRCxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxlQUFlLENBQUMsQ0FBQztJQUN0RCxDQUFDLENBQUMsQ0FBQztJQUVILElBQUksQ0FBQyxTQUFTLEVBQUUsR0FBRyxFQUFFO1FBQ3BCLE1BQU0sZUFBZSxHQUFHLDJEQUEyRCxDQUFDO1FBQ3BGLE1BQU0sZUFBZSxHQUFHLDJEQUEyRCxDQUFDO1FBQ3BGLE1BQU0sV0FBVyxHQUFHLEtBQUssQ0FBQyxpQkFBaUIsQ0FBQyxFQUFFLEVBQUUsb0JBQW9CLENBQUMsQ0FBQztRQUN0RSxNQUFNLENBQUMsV0FBVyxDQUFDLE1BQU0sQ0FBQyxDQUFDLElBQUksQ0FBQyxDQUFDLENBQUMsQ0FBQztRQUNuQyxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxlQUFlLENBQUMsQ0FBQztRQUNyRCxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxlQUFlLENBQUMsQ0FBQztJQUN0RCxDQUFDLENBQUMsQ0FBQztJQUVILElBQUksQ0FBQyxTQUFTLEVBQUUsR0FBRyxFQUFFO1FBQ3BCLE1BQU0sZUFBZSxHQUFHLDJEQUEyRCxDQUFDO1FBQ3BGLE1BQU0sZUFBZSxHQUFHLDJEQUEyRCxDQUFDO1FBQ3BGLE1BQU0sZUFBZSxHQUFHLDJEQUEyRCxDQUFDO1FBQ3BGLE1BQU0sV0FBVyxHQUFHLEtBQUssQ0FBQyxpQkFBaUIsQ0FBQyxFQUFFLEVBQUUsb0JBQW9CLENBQUMsQ0FBQztRQUN0RSxNQUFNLENBQUMsV0FBVyxDQUFDLE1BQU0sQ0FBQyxDQUFDLElBQUksQ0FBQyxDQUFDLENBQUMsQ0FBQztRQUNuQyxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxlQUFlLENBQUMsQ0FBQztRQUNyRCxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxlQUFlLENBQUMsQ0FBQztRQUNyRCxNQUFNLENBQUMsV0FBVyxDQUFDLENBQUMsQ0FBQyxDQUFDLE9BQU8sQ0FBQyxDQUFDLElBQUksQ0FBQyxlQUFlLENBQUMsQ0FBQztJQUN0RCxDQUFDLENBQUMsQ0FBQztBQUVKLENBQUMsQ0FBQyxDQUFDIn0=