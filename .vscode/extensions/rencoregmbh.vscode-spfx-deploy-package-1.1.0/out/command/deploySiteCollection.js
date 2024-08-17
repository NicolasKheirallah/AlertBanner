"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const __1 = require("..");
function deploySiteCollection(fileUri, auth, output) {
    if (!__1.Utils.warnIfNotSppkg(fileUri)) {
        return;
    }
    __1.Utils.deploySolution(fileUri, false, false, auth, output);
}
exports.deploySiteCollection = deploySiteCollection;
//# sourceMappingURL=deploySiteCollection.js.map