"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const vscode = require("vscode");
class Output extends vscode.Disposable {
    constructor() {
        super(() => {
            this.channel.dispose();
        });
        this.channel = vscode.window.createOutputChannel('Rencore Deploy SPFx Package');
    }
    show(preserveFocus = true) {
        this.channel.show(preserveFocus);
    }
    write(s) {
        this.channel.appendLine(s);
    }
}
exports.Output = Output;
//# sourceMappingURL=Output.js.map