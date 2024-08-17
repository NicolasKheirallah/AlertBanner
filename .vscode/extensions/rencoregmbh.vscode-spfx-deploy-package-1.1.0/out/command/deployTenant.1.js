"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const __1 = require("..");
function deployTenant(fileUri, auth, output) {
    if (!__1.Utils.warnIfNotSppkg(fileUri)) {
        return;
    }
    __1.Utils.deploySolution(fileUri, true, false, auth, output);
}
exports.deployTenant = deployTenant;
//# sourceMappingURL=deployTenant.1.js.map