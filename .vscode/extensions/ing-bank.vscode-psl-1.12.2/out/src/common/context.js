"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.getFullContext = void 0;
const vscode = require("vscode");
const fsExtra = require("fs-extra");
function getFullContext(context, useActiveTextEditor) {
    let fsPath = '';
    let mode;
    let activeTextEditor = vscode.window.activeTextEditor;
    if (context) {
        fsPath = context.fsPath;
        mode = fsExtra.lstatSync(fsPath).isFile() ? 1 /* ContextMode.FILE */ : 2 /* ContextMode.DIRECTORY */;
        return { fsPath, mode };
    }
    else if (useActiveTextEditor && activeTextEditor) {
        fsPath = activeTextEditor.document.fileName;
        mode = 1 /* ContextMode.FILE */;
        return { fsPath, mode };
    }
    else {
        mode = 3 /* ContextMode.EMPTY */;
        return { fsPath, mode };
    }
}
exports.getFullContext = getFullContext;
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoiY29udGV4dC5qcyIsInNvdXJjZVJvb3QiOiIiLCJzb3VyY2VzIjpbIi4uLy4uLy4uL3NyYy9jb21tb24vY29udGV4dC50cyJdLCJuYW1lcyI6W10sIm1hcHBpbmdzIjoiOzs7QUFBQSxpQ0FBaUM7QUFDakMsb0NBQW9DO0FBaUJwQyxTQUFnQixjQUFjLENBQUMsT0FBNEMsRUFBRSxtQkFBNkI7SUFDekcsSUFBSSxNQUFNLEdBQVcsRUFBRSxDQUFDO0lBQ3hCLElBQUksSUFBaUIsQ0FBQztJQUN0QixJQUFJLGdCQUFnQixHQUFHLE1BQU0sQ0FBQyxNQUFNLENBQUMsZ0JBQWdCLENBQUM7SUFFdEQsSUFBSSxPQUFPLEVBQUU7UUFDWixNQUFNLEdBQUcsT0FBTyxDQUFDLE1BQU0sQ0FBQztRQUN4QixJQUFJLEdBQUcsT0FBTyxDQUFDLFNBQVMsQ0FBQyxNQUFNLENBQUMsQ0FBQyxNQUFNLEVBQUUsQ0FBQyxDQUFDLDBCQUFrQixDQUFDLDhCQUFzQixDQUFDO1FBQ3JGLE9BQU8sRUFBQyxNQUFNLEVBQUUsSUFBSSxFQUFDLENBQUM7S0FDdEI7U0FDSSxJQUFJLG1CQUFtQixJQUFJLGdCQUFnQixFQUFFO1FBQ2pELE1BQU0sR0FBRyxnQkFBZ0IsQ0FBQyxRQUFRLENBQUMsUUFBUSxDQUFDO1FBQzVDLElBQUksMkJBQW1CLENBQUM7UUFDeEIsT0FBTyxFQUFDLE1BQU0sRUFBRSxJQUFJLEVBQUMsQ0FBQztLQUN0QjtTQUNJO1FBQ0osSUFBSSw0QkFBb0IsQ0FBQztRQUN6QixPQUFPLEVBQUMsTUFBTSxFQUFFLElBQUksRUFBQyxDQUFDO0tBQ3RCO0FBQ0YsQ0FBQztBQW5CRCx3Q0FtQkMifQ==