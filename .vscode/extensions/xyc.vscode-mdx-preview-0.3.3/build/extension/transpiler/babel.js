"use strict";
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (Object.hasOwnProperty.call(mod, k)) result[k] = mod[k];
    result["default"] = mod;
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
const babel = __importStar(require("@babel/core"));
// can probably load directory .babelrc config from here?
const babelOptions = {
    "presets": [
        babel.createConfigItem([require("@babel/preset-env"), { exclude: ["transform-regenerator"] }]),
        babel.createConfigItem(require("@babel/preset-react")),
    ],
    "plugins": [
        babel.createConfigItem(require("@babel/plugin-proposal-export-default-from")),
        // selected plugins from https://github.com/babel/babel/blob/master/packages/babel-preset-stage-2/README.md
        babel.createConfigItem(require("@babel/plugin-proposal-export-namespace-from")),
        babel.createConfigItem(require("@babel/plugin-proposal-class-properties")),
        babel.createConfigItem(require("@babel/plugin-proposal-optional-chaining")),
        babel.createConfigItem(require("@babel/plugin-proposal-nullish-coalescing-operator")),
        babel.createConfigItem(require("@babel/plugin-syntax-dynamic-import")),
        babel.createConfigItem(require("babel-plugin-transform-dynamic-import")),
    ]
};
exports.transformAsync = code => {
    return babel.transformAsync(code, babelOptions);
};
//# sourceMappingURL=babel.js.map