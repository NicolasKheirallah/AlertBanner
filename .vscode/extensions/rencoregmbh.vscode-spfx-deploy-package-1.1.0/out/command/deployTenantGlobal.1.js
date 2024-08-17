"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const __1 = require("..");
function deployTenantGlobal(fileUri, auth, output) {
    if (!__1.Utils.warnIfNotSppkg(fileUri)) {
        return;
    }
    __1.Utils.deploySolution(fileUri, true, true, auth, output);
}
exports.deployTenantGlobal = deployTenantGlobal;
//# sourceMappingURL=deployTenantGlobal.1.js.map