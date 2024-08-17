"use strict";
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (Object.hasOwnProperty.call(mod, k)) result[k] = mod[k];
    result["default"] = mod;
    return result;
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const path = __importStar(require("path"));
const is_module_1 = __importDefault(require("is-module"));
const mdx_1 = require("../transpiler/mdx/mdx");
const typescript_1 = require("typescript");
const babel_1 = require("../transpiler/babel");
const sucrase_1 = require("../transpiler/sucrase");
async function transformEntry(code, fsPath, preview) {
    const { languageId, uri } = preview.doc;
    if (languageId === 'markdown' || languageId === 'mdx' || uri.scheme === 'untitled') {
        code = await mdx_1.mdxTranspileAsync(code, true, preview);
    }
    let useSucrase = preview.configuration.useSucraseTranspiler;
    if ((languageId === 'typescript' || languageId === 'typescriptreact') && !useSucrase) {
        // in case user hasn't provided a tsconfig.json, generate a default one
        if (!preview.typescriptConfiguration) {
            preview.generateTypescriptConfiguration(null);
        }
        const { tsCompilerOptions } = preview.typescriptConfiguration;
        code = typescript_1.transpileModule(code, {
            compilerOptions: tsCompilerOptions,
            fileName: fsPath,
        }).outputText;
    }
    if (useSucrase) {
        try {
            code = (sucrase_1.transform(code)).code;
        }
        catch (error) {
            code = (await babel_1.transformAsync(code)).code;
        }
    }
    else {
        code = (await babel_1.transformAsync(code)).code;
    }
    return code;
}
exports.transformEntry = transformEntry;
async function transform(code, fsPath, preview) {
    const extname = path.extname(fsPath);
    if (/\.mdx?$/i.test(extname)) {
        code = await mdx_1.mdxTranspileAsync(code, false, preview);
    }
    let useSucrase = preview.configuration.useSucraseTranspiler;
    if (/\.tsx?$/i.test(extname) && !useSucrase) {
        if (!preview.typescriptConfiguration) {
            preview.generateTypescriptConfiguration(null);
        }
        const { tsCompilerOptions } = preview.typescriptConfiguration;
        code = typescript_1.transpileModule(code, {
            compilerOptions: tsCompilerOptions,
            fileName: fsPath,
        }).outputText;
    }
    // Transform:
    // - exclude node_modules
    // - include file in node_modules only if it's es module
    const isInNodeModules = fsPath.split(path.sep).includes('node_modules');
    if (!isInNodeModules ||
        is_module_1.default(code)) {
        console.log(`Transpiling: ${fsPath}`);
        // use sucrase to transpile node_module files as it's faster
        if (isInNodeModules || useSucrase) {
            try {
                code = (sucrase_1.transform(code)).code;
            }
            catch (error) {
                code = (await babel_1.transformAsync(code)).code;
            }
        }
        else {
            code = (await babel_1.transformAsync(code)).code;
        }
    }
    return code;
}
exports.transform = transform;
//# sourceMappingURL=transform.js.map