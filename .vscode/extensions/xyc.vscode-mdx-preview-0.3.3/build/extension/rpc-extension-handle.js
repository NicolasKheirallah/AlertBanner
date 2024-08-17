"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const { performance } = require('perf_hooks');
const module_fetcher_1 = require("./module-fetcher/module-fetcher");
class ExtensionHandle {
    constructor(preview) {
        this.preview = preview;
    }
    handshake() {
        this.preview.resolveWebviewHandshakePromise();
    }
    reportPerformance(evaluationDuration) {
        this.preview.evaluationDuration = evaluationDuration;
        performance.mark('preview/end');
        performance.measure('preview duration', 'preview/start', 'preview/end');
    }
    async fetch(request, isBare, parentId) {
        return module_fetcher_1.fetchLocal(request, isBare, parentId, this.preview);
    }
}
exports.default = ExtensionHandle;
//# sourceMappingURL=rpc-extension-handle.js.map