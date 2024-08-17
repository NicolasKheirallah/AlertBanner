"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.TodoInfo = void 0;
const tokenizer_1 = require("../parser/tokenizer");
const api_1 = require("./api");
class TodoInfo extends api_1.PslRule {
    report() {
        let todos = [];
        for (const token of this.parsedDocument.comments) {
            if (token.value.includes('TODO')) {
                const startLine = token.position.line;
                const startChar = token.position.character;
                todos = todos.concat(getTodosFromComment(token.value, startLine, startChar));
            }
        }
        return todos.map(todo => {
            const diagnostic = new api_1.Diagnostic(todo.range, todo.message, this.ruleName, api_1.DiagnosticSeverity.Information);
            diagnostic.source = 'TODO';
            return diagnostic;
        });
    }
}
exports.TodoInfo = TodoInfo;
function getTodosFromComment(commentText, startLine, startChar) {
    let todos = [];
    let todo;
    let currentLine;
    let currentChar;
    const finalize = () => {
        if (!todo)
            return;
        const start = todo.range.start;
        const end = new tokenizer_1.Position(currentLine, todo.range.end.character + todo.message.trimRight().length);
        todo.range = new tokenizer_1.Range(start, end);
        todo.message = todo.message.trim().replace(/^:/gm, '').trim();
        if (!todo.message)
            todo.message = `TODO on line ${todo.range.start.line + 1}.`;
        todos.push(todo);
        todo = undefined;
    };
    const tokens = (0, tokenizer_1.getTokens)(commentText);
    for (const token of tokens) {
        currentLine = startLine + token.position.line;
        currentChar = startLine === currentLine ? token.position.character + startChar : token.position.character;
        if (token.isBlockCommentInit() || token.isLineCommentInit())
            continue;
        else if (token.isBlockComment() || token.isLineComment()) {
            todos = todos.concat(getTodosFromComment(token.value, currentLine, currentChar));
        }
        else if (token.value === 'TODO' && !todo) {
            const range = new tokenizer_1.Range(currentLine, currentChar, currentLine, currentChar + 4);
            const message = '';
            todo = { range, message };
        }
        else if (todo) {
            if (token.isNewLine())
                finalize();
            else
                todo.message += token.value;
        }
    }
    if (todo)
        finalize();
    return todos;
}
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoidG9kb3MuanMiLCJzb3VyY2VSb290IjoiIiwic291cmNlcyI6WyIuLi8uLi8uLi9zcmMvcHNsTGludC90b2Rvcy50cyJdLCJuYW1lcyI6W10sIm1hcHBpbmdzIjoiOzs7QUFBQSxtREFBaUU7QUFDakUsK0JBQWdFO0FBRWhFLE1BQWEsUUFBUyxTQUFRLGFBQU87SUFFcEMsTUFBTTtRQUNMLElBQUksS0FBSyxHQUFXLEVBQUUsQ0FBQztRQUN2QixLQUFLLE1BQU0sS0FBSyxJQUFJLElBQUksQ0FBQyxjQUFjLENBQUMsUUFBUSxFQUFFO1lBQ2pELElBQUksS0FBSyxDQUFDLEtBQUssQ0FBQyxRQUFRLENBQUMsTUFBTSxDQUFDLEVBQUU7Z0JBQ2pDLE1BQU0sU0FBUyxHQUFHLEtBQUssQ0FBQyxRQUFRLENBQUMsSUFBSSxDQUFDO2dCQUN0QyxNQUFNLFNBQVMsR0FBRyxLQUFLLENBQUMsUUFBUSxDQUFDLFNBQVMsQ0FBQztnQkFDM0MsS0FBSyxHQUFHLEtBQUssQ0FBQyxNQUFNLENBQUMsbUJBQW1CLENBQUMsS0FBSyxDQUFDLEtBQUssRUFBRSxTQUFTLEVBQUUsU0FBUyxDQUFDLENBQUMsQ0FBQzthQUM3RTtTQUNEO1FBQ0QsT0FBTyxLQUFLLENBQUMsR0FBRyxDQUFDLElBQUksQ0FBQyxFQUFFO1lBQ3ZCLE1BQU0sVUFBVSxHQUFHLElBQUksZ0JBQVUsQ0FBQyxJQUFJLENBQUMsS0FBSyxFQUFFLElBQUksQ0FBQyxPQUFPLEVBQUUsSUFBSSxDQUFDLFFBQVEsRUFBRSx3QkFBa0IsQ0FBQyxXQUFXLENBQUMsQ0FBQztZQUMzRyxVQUFVLENBQUMsTUFBTSxHQUFHLE1BQU0sQ0FBQztZQUMzQixPQUFPLFVBQVUsQ0FBQztRQUNuQixDQUFDLENBQUMsQ0FBQztJQUNKLENBQUM7Q0FDRDtBQWpCRCw0QkFpQkM7QUFPRCxTQUFTLG1CQUFtQixDQUFDLFdBQW1CLEVBQUUsU0FBaUIsRUFBRSxTQUFpQjtJQUNyRixJQUFJLEtBQUssR0FBVyxFQUFFLENBQUM7SUFDdkIsSUFB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