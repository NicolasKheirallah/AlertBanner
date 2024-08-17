"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (Object.hasOwnProperty.call(mod, k)) result[k] = mod[k];
    result["default"] = mod;
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
const mdx_1 = __importDefault(require("@mdx-js/mdx"));
const hasDefaultExport_1 = __importDefault(require("./hasDefaultExport"));
const path = __importStar(require("path"));
const injectMDXStyles = (mdxText, preview) => {
    const { customLayoutFilePath, useVscodeMarkdownStyles, useWhiteBackground, } = preview.configuration;
    if (customLayoutFilePath) {
        try {
            const currentPreviewDirname = path.dirname(preview.doc.uri.fsPath);
            const relativeCustomLayoutPath = path.relative(currentPreviewDirname, customLayoutFilePath);
            return `import Layout from '.${path.sep}${relativeCustomLayoutPath}';

export default Layout;

${mdxText}`;
        }
        catch (error) {
            return mdxText;
        }
    }
    else if (useVscodeMarkdownStyles) {
        let layoutOptions = useWhiteBackground ? '{ forceLightTheme: true }' : '{}';
        return `import { createLayout } from 'vscode-markdown-layout';

export default createLayout(${layoutOptions});

${mdxText}`;
    }
    else {
        return mdxText;
    }
};
const wrapCompiledMdx = (compiledMDX, isEntry) => {
    if (isEntry) {
        // entry MDX, render on webview DOM
        return `import React from 'react';
import ReactDOM from 'react-dom';
import { MDXTag } from '@mdx-js/tag';
${compiledMDX}
ReactDOM.render(<MDXContent></MDXContent>, document.getElementById('root'));`;
    }
    else {
        // transclusion
        return `import React from 'react';
import { MDXTag } from '@mdx-js/tag';
${compiledMDX}`;
    }
};
exports.mdxTranspileAsync = async (mdxText, isEntry, preview) => {
    let mdxTextToCompile;
    if (!hasDefaultExport_1.default(mdxText)) {
        // inject vscode markdown styles if we haven't found a default export
        mdxTextToCompile = injectMDXStyles(mdxText, preview);
    }
    else {
        mdxTextToCompile = mdxText;
    }
    const compiledMDX = await mdx_1.default(mdxTextToCompile);
    return wrapCompiledMdx(compiledMDX, isEntry);
};
//# sourceMappingURL=mdx.js.map