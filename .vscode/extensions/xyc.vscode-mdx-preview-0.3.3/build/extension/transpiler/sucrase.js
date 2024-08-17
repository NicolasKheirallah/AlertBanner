"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const sucrase_1 = require("sucrase");
exports.transform = (code) => {
    return sucrase_1.transform(code, { transforms: ["jsx", "typescript", "imports"] });
};
//# sourceMappingURL=sucrase.js.map