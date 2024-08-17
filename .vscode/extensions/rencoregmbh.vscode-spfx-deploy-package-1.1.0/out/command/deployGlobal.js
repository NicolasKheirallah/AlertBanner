"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const __1 = require("../");
function deployGlobal(fileUri, auth, output) {
    if (!__1.Utils.warnIfNotSppkg(fileUri)) {
        return;
    }
    __1.Utils.deploySolution(fileUri, true, auth, output);
}
exports.deployGlobal = deployGlobal;
//# sourceMappingURL=deployGlobal.js.map