"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.PSLFormatProvider = exports.activate = void 0;
const vscode = require("vscode");
const parser = require("../parser/parser");
const extension_1 = require("../extension");
function activate(context) {
    context.subscriptions.push(vscode.languages.registerDocumentFormattingEditProvider(extension_1.PSL_MODE, new PSLFormatProvider()));
}
exports.activate = activate;
class PSLFormatProvider {
    provideDocumentFormattingEdits(document) {
        let textEdits = [];
        return new Promise(resolve => {
            let p = parser.parseText(document.getText());
            p.methods.forEach(method => {
                if (!method.closeParen)
                    return;
                method.memberClass;
                let methodLine = method.id.position.line;
                let closePosition = method.closeParen.position;
                let methodRange = new vscode.Range(methodLine, 0, closePosition.line, closePosition.character + 1);
                textEdits.push(new vscode.TextEdit(methodRange, buildText(method)));
            });
            resolve(textEdits);
        });
    }
}
exports.PSLFormatProvider = PSLFormatProvider;
function buildText(method) {
    let methodString = '';
    if (method.modifiers.length > 0) {
        methodString += method.modifiers.map(m => m.value).join(' ') + ' ';
    }
    methodString += `${method.id.value}(`;
    let parameterStrings = method.parameters.map(p => {
        let param = { parameter: '', comment: '' };
        let parameterString = '';
        if (p.req) {
            parameterString += 'req ';
        }
        if (p.ret) {
            parameterString += 'ret ';
        }
        if (p.literal) {
            parameterString += 'literal ';
        }
        parameterString += p.types[0].value + ' ' + p.id.value;
        if (p.types.length > 1) {
            parameterString += `( ${p.types.map(t => t.value).slice(1).join(', ')})`;
        }
        if (p.comment) {
            param.comment = `\t// ${p.comment.value.trim()}`;
        }
        param.parameter = parameterString;
        return param;
    });
    if (parameterStrings.length === 0) {
        methodString += ')';
    }
    else if (parameterStrings.length === 1) {
        methodString += parameterStrings[0].parameter + ')' + parameterStrings[0].comment;
    }
    else {
        methodString += '\n\t\t  ' + parameterStrings.map(p => p.parameter + p.comment).join('\n\t\t, ');
        methodString += '\n\t\t)';
    }
    return methodString;
}
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoicHNsRm9ybWF0LmpzIiwic291cmNlUm9vdCI6IiIsInNvdXJjZXMiOlsiLi4vLi4vLi4vc3JjL2xhbmd1YWdlL3BzbEZvcm1hdC50cyJdLCJuYW1lcyI6W10sIm1hcHBpbmdzIjoiOzs7QUFBQSxpQ0FBaUM7QUFDakMsMkNBQTJDO0FBQzNDLDRDQUF3QztBQUV4QyxTQUFnQixRQUFRLENBQUMsT0FBZ0M7SUFFeEQsT0FBTyxDQUFDLGFBQWEsQ0FBQyxJQUFJLENBQ3pCLE1BQU0sQ0FBQyxTQUFTLENBQUMsc0NBQXNDLENBQ3RELG9CQUFRLEVBQUUsSUFBSSxpQkFBaUIsRUFBRSxDQUNqQyxDQUNELENBQUM7QUFFSCxDQUFDO0FBUkQsNEJBUUM7QUFFRCxNQUFhLGlCQUFpQjtJQUM3Qiw4QkFBOEIsQ0FBQyxRQUE2QjtRQUMzRCxJQUFJLFNBQVMsR0FBc0IsRUFBRSxDQUFDO1FBQ3RDLE9BQU8sSUFBSSxPQUFPLENBQUMsT0FBTyxDQUFDLEVBQUU7WUFDNUIsSUFBSSxDQUFDLEdBQUcsTUFBTSxDQUFDLFNBQVMsQ0FBQyxRQUFRLENBQUMsT0FBTyxFQUFFLENBQUMsQ0FBQztZQUM3QyxDQUFDLENBQUMsT0FBTyxDQUFDLE9BQU8sQ0FBQyxNQUFNLENBQUMsRUFBRTtnQkFDMUIsSUFBSSxDQUFDLE1BQU0sQ0FBQyxVQUFVO29CQUFFLE9BQU87Z0JBQy9CLE1BQU0sQ0FBQyxXQUFXLENBQUE7Z0JBQ2xCLElBQUksVUFBVSxHQUFHLE1BQU0sQ0FBQyxFQUFFLENBQUMsUUFBUSxDQUFDLElBQUksQ0FBQztnQkFDekMsSUFBSSxhQUFhLEdBQUcsTUFBTSxDQUFDLFVBQVUsQ0FBQyxRQUFRLENBQUM7Z0JBQy9DLElBQUksV0FBVyxHQUFHLElBQUksTUFBTSxDQUFDLEtBQUssQ0FBQyxVQUFVLEVBQUUsQ0FBQyxFQUFFLGFBQWEsQ0FBQyxJQUFJLEVBQUUsYUFBYSxDQUFDLFNBQVMsR0FBRyxDQUFDLENBQUMsQ0FBQTtnQkFDbEcsU0FBUyxDQUFDLElBQUksQ0FBQyxJQUFJ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