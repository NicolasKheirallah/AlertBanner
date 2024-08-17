"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const request = require("request-promise");
var ScriptType;
(function (ScriptType) {
    ScriptType[ScriptType["module"] = 1] = "module";
    ScriptType[ScriptType["nonModule"] = 2] = "nonModule";
})(ScriptType = exports.ScriptType || (exports.ScriptType = {}));
class ScriptChecker {
    /**
     * Check the script type
     * @param scriptData
     */
    static check(scriptData) {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                // Check the script type
                const scriptTypeData = yield request(this._scriptCheckAPI, {
                    method: "POST",
                    headers: {
                        "content-type": "application/json",
                        "accept": "application/json"
                    },
                    body: JSON.stringify(scriptData)
                });
                if (!scriptTypeData) {
                    return null;
                }
                const script = JSON.parse(scriptTypeData);
                if (!script.scriptType) {
                    return null;
                }
                switch (script.scriptType) {
                    case "module":
                        return ScriptType.module;
                    case "non-module":
                        return ScriptType.nonModule;
                    default:
                        return ScriptType.module;
                }
                return null;
            }
            catch (err) {
                return 'Sorry, something went wrong';
            }
        });
    }
}
ScriptChecker._scriptCheckAPI = "https://scriptcheck-weu-fn.azurewebsites.net/api/script-check";
exports.default = ScriptChecker;
//# sourceMappingURL=scriptChecker.js.map