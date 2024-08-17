"use strict";
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (Object.hasOwnProperty.call(mod, k)) result[k] = mod[k];
    result["default"] = mod;
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
const vscode = __importStar(require("vscode"));
const selectSecurityPolicy = async () => {
    const extensionConfig = vscode.workspace.getConfiguration('mdx-preview');
    const securityPolicy = extensionConfig.get('preview.security', "strict" /* Strict */);
    const pickItems = [
        {
            type: "strict" /* Strict */,
            label: 'strict',
            description: 'Do not allow insecure content or eval',
        },
        {
            type: "disabled" /* Disabled */,
            label: 'disabled',
            description: 'Allow insecure content (not recommended)',
        },
    ];
    const currentPolicyItem = pickItems.find(pickItem => {
        return pickItem.type === securityPolicy;
    });
    if (currentPolicyItem) {
        currentPolicyItem.label = `• ${currentPolicyItem.label}`;
    }
    const selectedSecurityPolicyItem = await vscode.window.showQuickPick(pickItems);
    extensionConfig.update('preview.security', selectedSecurityPolicyItem.type);
};
exports.selectSecurityPolicy = selectSecurityPolicy;
//# sourceMappingURL=security.js.map