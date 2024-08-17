"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
;
class TextHelper {
    /**
     * Strip quotes at the beginning and end of the string
     *
     * @param value
     */
    static stripQuotes(value) {
        // Strip the comma
        if (value.endsWith(",")) {
            value = value.substring(0, value.length - 1);
        }
        if ((value.startsWith(`'`) && value.endsWith(`'`)) ||
            (value.startsWith(`"`) && value.endsWith(`"`)) ||
            (value.startsWith("`") && value.endsWith("`"))) {
            return value.substring(1, value.length - 1);
        }
        return value;
    }
    // find proper position for inserting text in typescript file
    static findInsertPosition(fileLines, localeKey, regexSet) {
        let result = -1;
        let inScope = false;
        for (let row = 0; row < fileLines.length; ++row) {
            const line = fileLines[row];
            if (inScope) {
                const idMatches = line.match(regexSet.line);
                if (idMatches !== null && idMatches.length > 1) {
                    const rowKey = idMatches[1];
                    if (rowKey.toLowerCase() < localeKey.toLowerCase()) {
                        result = row;
                    }
                }
                const endMatches = line.match(regexSet.end);
                if (endMatches !== null && endMatches.length > 1) {
                    inScope = false;
                }
            }
            const startMatches = line.match(regexSet.start);
            if (startMatches !== null && startMatches.length > 1) {
                result = row;
                inScope = true;
            }
        }
        return result;
    }
}
exports.default = TextHelper;
//start: "declare interface {", line: "some_id: string;", end: "}" 
TextHelper.findPositionTs = {
    start: /^\s*declare\s*interface\s*(\w+)\s*\{\s*$/,
    line: /^\s*(\w+)\s*\:\s*string\s*;\s*$/,
    end: /^\s*(\})\s*$/,
};
//start: "return {", line: "some_id: string,", end: "}" 
TextHelper.findPositionJs = {
    start: /^\s*(return\s*\{)\s*$/,
    line: /^\s*(\w+)\s*\:.*,\s*$/,
    end: /^\s*(\})\s*;\s*$/,
};
//# sourceMappingURL=TextHelper.js.map