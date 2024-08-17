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
exports.activate = void 0;
const vscode = require("vscode");
function activate(context) {
    return __awaiter(this, void 0, void 0, function* () {
        context.subscriptions.push(vscode.commands.registerCommand('psl.stepIn', stepIn));
        context.subscriptions.push(vscode.commands.registerCommand('psl.stepOut', stepOut));
        context.subscriptions.push(vscode.commands.registerCommand('psl.stepOver', stepOver));
        context.subscriptions.push(vscode.commands.registerCommand('psl.sendToHostTerminal', sendToHostTerminal));
        terminalSendSettings();
        configureGtmDebug(context);
    });
}
exports.activate = activate;
function terminalSendSettings() {
    return __awaiter(this, void 0, void 0, function* () {
        const pslTerminalCommands = ['psl.stepIn', 'psl.stepOut', 'psl.stepOver', 'psl.sendToHostTerminal'];
        const terminalSettings = vscode.workspace.getConfiguration('terminal');
        const commandsToSkip = terminalSettings.get('integrated.commandsToSkipShell');
        if (commandsToSkip) {
            const merged = commandsToSkip.concat(pslTerminalCommands);
            const filteredMerge = merged.filter((item, pos) => merged.indexOf(item) === pos);
            terminalSettings.update('integrated.commandsToSkipShell', filteredMerge, true);
        }
    });
}
function stepIn() {
    terminalSend('ZSTEP INTO:"W $ZPOS,! ZP @$ZPOS B"');
}
function stepOut() {
    terminalSend('ZSTEP OUTOF:"W $ZPOS,! ZP @$ZPOS B"');
}
function stepOver() {
    terminalSend('ZSTEP OVER:"W $ZPOS,! ZP @$ZPOS B"');
}
function sendToHostTerminal(text) {
    terminalSend(text);
}
function terminalSend(text) {
    const activeTerminal = vscode.window.activeTerminal;
    if (activeTerminal) {
        activeTerminal.show();
        activeTerminal.sendText(text, true);
    }
}
function configureGtmDebug(context) {
    const statusBar = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Right, 901);
    const commandName = 'psl.setGtmDebug';
    let gtmDebug = vscode.workspace.getConfiguration('psl').get('gtmDebugEnabled');
    const set = () => {
        if (gtmDebug)
            showInformation(context);
        statusBar.text = `GT.M Debug ${gtmDebug ? '$(check)' : '$(circle-slash)'}`;
        vscode.commands.executeCommand('setContext', 'psl.gtmDebug', gtmDebug);
    };
    set();
    context.subscriptions.push(vscode.commands.registerCommand(commandName, () => {
        gtmDebug = !gtmDebug;
        set();
    }));
    vscode.workspace.onDidChangeConfiguration(event => {
        if (event.affectsConfiguration('psl.gtmDebugEnabled')) {
            gtmDebug = true;
            set();
        }
    });
    statusBar.command = commandName;
    statusBar.tooltip = 'GT.M Debug hotkeys';
    statusBar.show();
}
function showInformation(context) {
    return __awaiter(this, void 0, void 0, function* () {
        const doNotShow = context.globalState.get('gtmDebugShow');
        if (doNotShow)
            return;
        const response = yield vscode.window.showInformationMessage('INTO Ctrl+Q | OVER Ctrl+W | OUTOF Ctrl+E', 'Do not show again');
        if (response) {
            context.globalState.update('gtmDebugShow', true);
        }
    });
}
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoidGVybWluYWwuanMiLCJzb3VyY2VSb290IjoiIiwic291cmNlcyI6WyIuLi8uLi8uLi9zcmMvY29tbW9uL3Rlcm1pbmFsLnRzIl0sIm5hbWVzIjpbXSwibWFwcGluZ3MiOiI7Ozs7Ozs7Ozs7OztBQUFBLGlDQUFpQztBQUVqQyxTQUFzQixRQUFRLENBQUMsT0FBZ0M7O1FBRTlELE9BQU8sQ0FBQyxhQUFhLENBQUMsSUFBSSxDQUN6QixNQUFNLENBQUMsUUFBUSxDQUFDLGVBQWUsQ0FDOUIsWUFBWSxFQUFFLE1BQU0sQ0FDcEIsQ0FDRCxDQUFDO1FBRUYsT0FBTyxDQUFDLGFBQWEsQ0FBQyxJQUFJLENBQ3pCLE1BQU0sQ0FBQyxRQUFRLENBQUMsZUFBZSxDQUM5QixhQUFhLEVBQUUsT0FBTyxDQUN0QixDQUNELENBQUM7UUFFRixPQUFPLENBQUMsYUFBYSxDQUFDLElBQUksQ0FDekIsTUFBTSxDQUFDLFFBQVEsQ0FBQyxlQUFlLENBQzlCLGNBQWMsRUFBRSxRQUFRLENBQ3hCLENBQ0QsQ0FBQztRQUVGLE9BQU8sQ0FBQyxhQUFhLENBQUMsSUFBSSxDQUN6QixNQUFNLENBQUMsUUFBUSxDQUFDLGVBQWUsQ0FDOUIsd0JBQXdCLEVBQUUsa0JBQWtCLENBQzVDLENBQ0QsQ0FBQztRQUNGLG9CQUFvQixFQUFFLENBQUM7UUFFdkIsaUJBQWlCLENBQUMsT0FBTyxDQUFDLENBQUM7SUFDNUIsQ0FBQztDQUFBO0FBNUJELDRCQTRCQztBQUVELFNBQWUsb0JBQW9COztRQUNsQyxNQUFNLG1CQUFtQixHQUFHLENBQUMsWUFBWSxFQUFFLGFBQWEsRUFBRSxjQUFjLEVBQUUsd0JBQXdCLENBQUMsQ0FBQztRQUNwRyxNQUFNLGdCQUFnQixHQUFHLE1BQU0sQ0FBQyxTQUFTLENBQUMsZ0JBQWdCLENBQUMsVUFBVSxDQUFDLENBQUM7UUFDdkUsTUFBTSxjQUFjLEdBQXlCLGdCQUFnQixDQUFDLEdBQUcsQ0FBQyxnQ0FBZ0MsQ0FBQyxDQUFDO1FBQ3BHLElBQUksY0FBYyxFQUFFO1lB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