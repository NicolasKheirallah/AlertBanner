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
const vscode = __importStar(require("vscode"));
const fs = __importStar(require("fs"));
const path = __importStar(require("path"));
const typescript = __importStar(require("typescript"));
const webview_manager_1 = require("./webview-manager");
const evaluate_in_webview_1 = __importDefault(require("./evaluate-in-webview"));
const { performance, PerformanceObserver } = require('perf_hooks');
class Preview {
    constructor(doc) {
        this.setDoc(doc);
        const extensionConfig = vscode.workspace.getConfiguration('mdx-preview', doc.uri);
        this.configuration = {
            previewOnChange: extensionConfig.get('preview.previewOnChange', true),
            useSucraseTranspiler: extensionConfig.get('build.useSucraseTranspiler', false),
            useVscodeMarkdownStyles: extensionConfig.get('preview.useVscodeMarkdownStyles', true),
            useWhiteBackground: extensionConfig.get('preview.useWhiteBackground', false),
            customLayoutFilePath: extensionConfig.get('preview.mdx.customLayoutFilePath', ""),
            securityPolicy: extensionConfig.get('preview.security', "strict" /* Strict */)
        };
        if (process.env.NODE_ENV === 'development') {
            this.performanceObserver = new PerformanceObserver((list, observer) => {
                this.previewDuration = list.getEntries()[0].duration;
                vscode.window.showInformationMessage(`Previewing used: ${Number(this.previewDuration / 1000).toFixed(2)} seconds. 
            Evaluation used: ${Number(this.evaluationDuration / 1000).toFixed(2)} seconds.`);
                performance.clearMarks();
            });
            this.performanceObserver.observe({ entryTypes: ['measure'] });
        }
    }
    initWebviewHandshakePromise() {
        this.webviewHandshakePromise = new Promise(resolve => {
            this.resolveWebviewHandshakePromise = () => {
                resolve();
            };
        });
    }
    get styleConfiguration() {
        return {
            useVscodeMarkdownStyles: this.configuration.useVscodeMarkdownStyles,
            useWhiteBackground: this.configuration.useWhiteBackground
        };
    }
    get securityConfiguration() {
        return { securityPolicy: this.configuration.securityPolicy };
    }
    generateTypescriptConfiguration(configFile) {
        let tsCompilerOptions, tsCompilerHost;
        if (configFile) {
            const configContents = fs.readFileSync(configFile).toString();
            const configJson = typescript.parseConfigFileTextToJson(configFile, configContents).config.compilerOptions;
            tsCompilerOptions = typescript.convertCompilerOptionsFromJson(configJson, configFile).options;
        }
        else {
            tsCompilerOptions = typescript.getDefaultCompilerOptions();
        }
        delete tsCompilerOptions.emitDeclarationOnly;
        delete tsCompilerOptions.declaration;
        tsCompilerOptions.module = typescript.ModuleKind.ESNext;
        tsCompilerOptions.target = typescript.ScriptTarget.ESNext;
        tsCompilerOptions.noEmitHelpers = false;
        tsCompilerOptions.importHelpers = false;
        tsCompilerHost = typescript.createCompilerHost(tsCompilerOptions);
        this.typescriptConfiguration = {
            tsCompilerHost,
            tsCompilerOptions
        };
    }
    setDoc(doc) {
        this.doc = doc;
        this.dependentFsPaths = new Set([doc.uri.fsPath]);
        let configFile = typescript.findConfigFile(this.entryFsDirectory, typescript.sys.fileExists);
        if (configFile) {
            this.generateTypescriptConfiguration(configFile);
        }
        else {
            this.typescriptConfiguration = undefined;
        }
    }
    get fsPath() {
        return this.doc.uri.fsPath;
    }
    get text() {
        return this.doc.getText();
    }
    /**
     * Entry fs directory for resolving bare imports.
     * For untitled documents, we are trying to get the workspace root. If there's no workspace root,
     * we will not resolve any bare imports.
     * For file documents, it's directory that document is located.
     */
    get entryFsDirectory() {
        if (this.doc.uri.scheme === 'untitled') {
            const rootWorkspaceFolder = vscode.workspace.workspaceFolders[0];
            if (!rootWorkspaceFolder) {
                return null;
            }
            return rootWorkspaceFolder.uri.fsPath;
        }
        else if (this.doc.uri.scheme === 'file') {
            return path.dirname(this.fsPath);
        }
        return null;
    }
    updateWebview() {
        const preview = this;
        const { uri } = preview.doc;
        const { scheme, fsPath } = uri;
        switch (scheme) {
            case 'untitled': {
                evaluate_in_webview_1.default(preview, preview.text, preview.entryFsDirectory);
                return;
            }
            case 'file': {
                if (this.configuration.previewOnChange) {
                    evaluate_in_webview_1.default(preview, preview.text, fsPath);
                }
                else {
                    const text = fs.readFileSync(fsPath, { encoding: 'utf8' });
                    evaluate_in_webview_1.default(preview, text, fsPath);
                }
                break;
            }
            default:
                break;
        }
    }
    refreshWebview() {
        webview_manager_1.refreshPanel(exports.currentPreview);
        this.updateWebview();
    }
    handleDidChangeTextDocument(fsPath, doc) {
        if (this.active) {
            if (this.configuration.previewOnChange) {
                if (this.dependentFsPaths.has(fsPath)) {
                    this.editingDoc = doc;
                    if (fsPath !== this.fsPath) {
                        this.webviewHandle.invalidate(fsPath)
                            .then(() => {
                            this.updateWebview();
                        });
                    }
                    else {
                        // not necessary to invalidate entry path
                        this.updateWebview();
                    }
                }
            }
        }
    }
    handleDidSaveTextDocument(fsPath) {
        if (this.active) {
            if (this.dependentFsPaths.has(fsPath)) {
                if (fsPath !== this.fsPath) {
                    this.webviewHandle.invalidate(fsPath)
                        .then(() => {
                        this.updateWebview();
                    });
                }
                else {
                    this.updateWebview();
                }
            }
        }
    }
    updateConfiguration() {
        const extensionConfig = vscode.workspace.getConfiguration('mdx-preview', this.doc.uri);
        const previewOnChange = extensionConfig.get('preview.previewOnChange', true);
        const useSucraseTranspiler = extensionConfig.get('build.useSucraseTranspiler', false);
        const useVscodeMarkdownStyles = extensionConfig.get('preview.useVscodeMarkdownStyles', true);
        const useWhiteBackground = extensionConfig.get('preview.useWhiteBackground', false);
        const customLayoutFilePath = extensionConfig.get('preview.mdx.customLayoutFilePath', "");
        const securityPolicy = extensionConfig.get('preview.security', "strict" /* Strict */);
        const needsWebviewRefresh = useVscodeMarkdownStyles !== this.configuration.useVscodeMarkdownStyles
            || useWhiteBackground !== this.configuration.useWhiteBackground
            || customLayoutFilePath !== this.configuration.customLayoutFilePath
            || securityPolicy !== this.configuration.securityPolicy;
        Object.assign(this.configuration, {
            previewOnChange,
            useSucraseTranspiler,
            useVscodeMarkdownStyles,
            useWhiteBackground,
            customLayoutFilePath,
            securityPolicy
        });
        if (needsWebviewRefresh) {
            this.refreshWebview();
        }
    }
}
exports.Preview = Preview;
function openPreview() {
    if (!vscode.window.activeTextEditor) {
        return;
    }
    const doc = vscode.window.activeTextEditor.document;
    if (!exports.currentPreview) {
        exports.currentPreview = new Preview(doc);
    }
    else {
        exports.currentPreview.setDoc(doc);
    }
    webview_manager_1.createOrShowPanel(exports.currentPreview);
    exports.currentPreview.updateWebview();
}
exports.openPreview = openPreview;
function refreshPreview() {
    if (!exports.currentPreview) {
        return;
    }
    // don't set doc
    webview_manager_1.refreshPanel(exports.currentPreview);
    exports.currentPreview.updateWebview();
}
exports.refreshPreview = refreshPreview;
//# sourceMappingURL=preview-manager.js.map