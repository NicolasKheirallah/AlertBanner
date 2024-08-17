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
const comlink = __importStar(require("@dxflow/comlink/umd/comlink"));
const rpc_extension_handle_1 = __importDefault(require("./rpc-extension-handle"));
class ExtensionEndpoint {
    constructor(webview, disposables) {
        this.webview = webview;
        this.disposables = disposables;
    }
    postMessage(message) {
        this.webview.postMessage(message);
    }
    addEventListener(type, listener) {
        this.currentListener = listener;
        this.disposeEventListener = this.webview.onDidReceiveMessage(message => {
            const messageEvent = {
                data: message,
                lastEventId: null,
                origin: null,
                ports: null,
                source: null,
                bubbles: null,
                cancelBubble: null,
                cancelable: null,
                composed: null,
                currentTarget: null,
                defaultPrevented: null,
                eventPhase: null,
                isTrusted: null,
                returnValue: null,
                srcElement: null,
                target: null,
                timeStamp: null,
                type: null,
                composedPath: null,
                initEvent: null,
                preventDefault: null,
                stopImmediatePropagation: null,
                AT_TARGET: null,
                stopPropagation: null,
                BUBBLING_PHASE: null,
                CAPTURING_PHASE: null,
                NONE: null,
            };
            listener(messageEvent);
        }, null, this.disposables);
    }
    removeEventListener(type, listener) {
        if (this.currentListener === listener && this.disposeEventListener) {
            this.disposeEventListener.dispose();
        }
    }
}
// NOTE: we only support 1 webview at this time
function initRPCExtensionSide(preview, webview, disposables) {
    const extensionEndpoint = new ExtensionEndpoint(webview, disposables);
    // Webview to extension calls
    const handle = new rpc_extension_handle_1.default(preview);
    comlink.expose(handle, extensionEndpoint);
    // Extension to webview calls
    const WebviewHandle = comlink.proxy(extensionEndpoint);
    return WebviewHandle;
}
exports.initRPCExtensionSide = initRPCExtensionSide;
//# sourceMappingURL=rpc-extension.js.map