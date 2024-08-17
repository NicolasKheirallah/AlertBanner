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
class CDNQuality {
    /**
     * Test the CDN quality
     * @param scriptData
     */
    static test(scriptData, isSharePointUrl) {
        return __awaiter(this, void 0, void 0, function* () {
            let scriptCDN = undefined;
            if (isSharePointUrl) {
                scriptCDN = {
                    max: 6,
                    score: 4
                };
            }
            else {
                scriptCDN = yield request(this._scriptInfoAPI, {
                    method: "POST",
                    headers: {
                        "content-type": "application/json",
                        "accept": "application/json"
                    },
                    body: JSON.stringify(scriptData)
                });
            }
            if (!scriptCDN) {
                return null;
            }
            let quality = null;
            // Show notification about the script CDN
            if (scriptCDN.score === scriptCDN.max) {
                quality = "good";
            }
            else if (scriptCDN.score >= 3 && scriptCDN.score < scriptCDN.max) {
                quality = "average";
            }
            else {
                quality = "poor";
            }
            if (quality) {
                return `The quality of the CDN that you are using is: ${quality}`;
            }
            return null;
        });
    }
}
CDNQuality._scriptInfoAPI = "https://scriptcheck-weu-fn.azurewebsites.net/api/script-info";
exports.default = CDNQuality;
//# sourceMappingURL=cdnQuality.js.map