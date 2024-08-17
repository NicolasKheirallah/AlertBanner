"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.updateStatus = void 0;
const vscode = require("vscode");
function updateStatus(status, langs) {
    if (langs.length === 0) {
        status.show();
    }
    else if (vscode.window.activeTextEditor && langs.indexOf(vscode.window.activeTextEditor.document.languageId) >= 0) {
        status.show();
    }
    else {
        status.hide();
    }
}
exports.updateStatus = updateStatus;
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoic3RhdHVzVXRpbHMuanMiLCJzb3VyY2VSb290IjoiIiwic291cmNlcyI6WyIuLi8uLi8uLi9zcmMvY29tbW9uL3N0YXR1c1V0aWxzLnRzIl0sIm5hbWVzIjpbXSwibWFwcGluZ3MiOiI7OztBQUFBLGlDQUFpQztBQUVqQyxTQUFnQixZQUFZLENBQUMsTUFBNEIsRUFBRSxLQUFvQjtJQUM5RSxJQUFJLEtBQUssQ0FBQyxNQUFNLEtBQUssQ0FBQyxFQUFFO1FBQ3ZCLE1BQU0sQ0FBQyxJQUFJLEVBQUUsQ0FBQztLQUNkO1NBQ0ksSUFBSSxNQUFNLENBQUMsTUFBTSxDQUFDLGdCQUFnQixJQUFJLEtBQUssQ0FBQyxPQUFPLENBQUMsTUFBTSxDQUFDLE1BQU0sQ0FBQyxnQkFBZ0IsQ0FBQyxRQUFRLENBQUMsVUFBVSxDQUFDLElBQUksQ0FBQyxFQUFFO1FBQ2xILE1BQU0sQ0FBQyxJQUFJLEVBQUUsQ0FBQztLQUNkO1NBQ0k7UUFDSixNQUFNLENBQUMsSUFBSSxFQUFFLENBQUM7S0FDZDtBQUNGLENBQUM7QUFWRCxvQ0FVQyJ9