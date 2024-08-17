"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const __1 = require("../");
function deploy(fileUri, auth, output) {
    if (!__1.Utils.warnIfNotSppkg(fileUri)) {
        return;
    }
    __1.Utils.deploySolution(fileUri, false, auth, output);
}
exports.deploy = deploy;
//# sourceMappingURL=deploy.js.map