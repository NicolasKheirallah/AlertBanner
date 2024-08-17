"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const __1 = require("..");
function deploySiteCollectionGlobal(fileUri, auth, output) {
    if (!__1.Utils.warnIfNotSppkg(fileUri)) {
        return;
    }
    __1.Utils.deploySolution(fileUri, false, true, auth, output);
}
exports.deploySiteCollectionGlobal = deploySiteCollectionGlobal;
//# sourceMappingURL=deploySiteCollectionGlobal.js.map