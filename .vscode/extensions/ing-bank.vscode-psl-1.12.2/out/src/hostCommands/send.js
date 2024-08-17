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
exports.sendTableHandler = exports.sendElementHandler = void 0;
const vscode = require("vscode");
const utils = require("./hostCommandUtils");
const path = require("path");
const fs = require("fs-extra");
const environment = require("../common/environment");
const icon = "\u21E7" /* utils.icons.SEND */;
function sendElementHandler(context) {
    return __awaiter(this, void 0, void 0, function* () {
        let c = utils.getFullContext(context);
        if (c.mode === 3 /* utils.ContextMode.EMPTY */) {
            let quickPick = yield environment.workspaceQuickPick();
            if (!quickPick)
                return;
            let chosenEnv = quickPick;
            let files = yield vscode.window.showOpenDialog({ defaultUri: vscode.Uri.file(chosenEnv.fsPath), canSelectMany: true, openLabel: 'Send' });
            if (!files)
                return;
            for (let fsPath of files.map(file => file.fsPath).sort(tableFirst)) {
                yield sendElement(fsPath).catch(() => { });
            }
        }
        else if (c.mode === 2 /* utils.ContextMode.DIRECTORY */) {
            let files = yield vscode.window.showOpenDialog({ defaultUri: vscode.Uri.file(c.fsPath), canSelectMany: true, openLabel: 'Send' });
            if (!files)
                return;
            let sortedFiles = files.map(uri => uri.fsPath).sort(tableFirst);
            for (let fsPath of sortedFiles) {
                yield sendElement(fsPath).catch(() => { });
            }
        }
        if (c.mode === 1 /* utils.ContextMode.FILE */) {
            return sendElement(c.fsPath).catch(() => { });
        }
        return;
    });
}
exports.sendElementHandler = sendElementHandler;
function sendTableHandler(context) {
    return __awaiter(this, void 0, void 0, function* () {
        let c = utils.getFullContext(context);
        if (c.mode === 3 /* utils.ContextMode.EMPTY */) {
            return;
        }
        if (c.mode === 1 /* utils.ContextMode.FILE */) {
            let tableName;
            if (path.extname(c.fsPath) === '.TBL') {
                tableName = path.basename(c.fsPath).split('.TBL')[0];
            }
            else if (path.extname(c.fsPath) === '.COL') {
                tableName = path.basename(c.fsPath).split('.COL')[0].split('-')[0];
            }
            else {
                return;
            }
            let files = yield fs.readdir(path.dirname(c.fsPath));
            let sortedFiles = files.filter(f => f.startsWith(tableName)).sort(tableFirst);
            if (sortedFiles.length > 99) {
                let resp = yield vscode.window.showInformationMessage(`Send ${sortedFiles.length} elements of ${tableName}?`, { modal: true }, 'Yes');
                if (resp !== 'Yes')
                    return;
            }
            for (let file of sortedFiles) {
                yield sendElement(path.join(path.dirname(c.fsPath), file)).catch(() => { });
            }
        }
        return;
    });
}
exports.sendTableHandler = sendTableHandler;
// async function sendDirectory(targetDir: string) {
// 	let fileNames = await fs.readdir(targetDir);
// 	let word = fileNames.length === 1 ? 'file' : 'files';
// 	let resp = await vscode.window.showInformationMessage(`Send contents of ${targetDir} (${fileNames.length} ${word})?`, { modal: true }, 'Yes');
// 	if (resp !== 'Yes') return;
// 	fileNames.sort(tableFirst);
// 	for (let index = 0; index < fileNames.length; index++) {
// 		let fileName = fileNames[index];
// 		// TODO what if element is a directory?
// 		await sendElement(path.join(targetDir, fileName));
// 	}
// }
function sendElement(fsPath) {
    return __awaiter(this, void 0, void 0, function* () {
        if (!fs.statSync(fsPath).isFile())
            return;
        let envs;
        try {
            envs = yield utils.getEnvironment(fsPath);
        }
        catch (e) {
            utils.logger.error(`${"\u274C" /* utils.icons.ERROR */} ${icon} Invalid environment configuration.`);
            return;
        }
        if (envs.length === 0) {
            utils.logger.error(`${"\u274C" /* utils.icons.ERROR */} ${icon} No environments selected.`);
            return;
        }
        let promises = [];
        for (let env of envs) {
            promises.push(utils.executeWithProgress(`${icon} ${path.basename(fsPath)} SEND`, () => __awaiter(this, void 0, void 0, function* () {
                yield vscode.workspace.openTextDocument(fsPath).then(doc => doc.save());
                utils.logger.info(`${"\u2026" /* utils.icons.WAIT */} ${icon} ${path.basename(fsPath)} SEND to ${env.name}`);
                let connection = yield utils.getConnection(env);
                yield connection.send(fsPath);
                connection.close();
                utils.logger.info(`${"\u2714" /* utils.icons.SUCCESS */} ${icon} ${path.basename(fsPath)} SEND to ${env.name} successful`);
            })).catch((e) => {
                utils.logger.error(`${"\u274C" /* utils.icons.ERROR */} ${icon} error in ${env.name} ${e.message}`);
            }));
        }
        ;
        yield Promise.all(promises);
    });
}
function tableFirst(a, b) {
    let aIsTable = a.endsWith('.TBL');
    let bIsTable = b.endsWith('.TBL');
    if (aIsTable && !bIsTable) {
        return -1;
    }
    else if (bIsTable && !aIsTable) {
        return 1;
    }
    return a.localeCompare(b);
}
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoic2VuZC5qcyIsInNvdXJjZVJvb3QiOiIiLCJzb3VyY2VzIjpbIi4uLy4uLy4uL3NyYy9ob3N0Q29tbWFuZHMvc2VuZC50cyJdLCJuYW1lcyI6W10sIm1hcHBpbmdzIjoiOzs7Ozs7Ozs7Ozs7QUFBQSxpQ0FBaUM7QUFDakMsNENBQTRDO0FBQzVDLDZCQUE2QjtBQUM3QiwrQkFBK0I7QUFDL0IscURBQXFEO0FBRXJELE1BQU0sSUFBSSxrQ0FBbUIsQ0FBQztBQUU5QixTQUFzQixrQkFBa0IsQ0FBQyxPQUFzQzs7UUFDOUUsSUFBSSxDQUFDLEdBQUcsS0FBSyxDQUFDLGNBQWMsQ0FBQyxPQUFPLENBQUMsQ0FBQztRQUN0QyxJQUFJLENBQUMsQ0FBQyxJQUFJLG9DQUE0QixFQUFFO1lBQ3ZDLElBQUksU0FBUyxHQUFHLE1BQU0sV0FBVyxDQUFDLGtCQUFrQixFQUFFLENBQUM7WUFDdkQsSUFBSSxDQUFDLFNBQVM7Z0JBQUUsT0FBTztZQUN2QixJQUFJLFNBQVMsR0FBRyxTQUFTLENBQUM7WUFDMUIsSUFBSSxLQUFLLEdBQUcsTUFBTSxNQUFNLENBQUMsTUFBTSxDQUFDLGNBQWMsQ0FBQyxFQUFFLFVBQVUsRUFBRSxNQUFNLENBQUMsR0FBRyxDQUFDLElBQUksQ0FBQyxTQUFTLENBQUMsTUFBTSxDQUFDLEVBQUUsYUFBYSxFQUFFLElBQUksRUFBRSxTQUFTLEVBQUUsTUFBTSxFQUFFLENBQUMsQ0FBQTtZQUN6SSxJQUFJLENBQUMsS0FBSztnQkFBRSxPQUFPO1lBQ25CLEtBQUssSUFBSSxNQUFNLElBQUksS0FBSyxDQUFDLEdBQUcsQ0FBQyxJQUFJLENBQUMsRUFBRSxDQUFDLElBQUksQ0FBQyxNQUFNLENBQUMsQ0FBQyxJQUFJLENBQUMsVUFBVSxDQUFDLEVBQUU7Z0JBQ25FLE1BQU0sV0FBVyxDQUFDLE1BQU0sQ0FBQyxDQUFDLEtBQUssQ0FBQyxHQUFHLEVBQUUsR0FBRyxDQUFDLENBQUMsQ0FBQzthQUMzQztTQUNEO2FBQ0ksSUFBSSxDQUFDLENBQUMsSUFBSSx3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