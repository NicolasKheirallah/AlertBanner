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
const path = __importStar(require("path"));
const path_is_inside_1 = __importDefault(require("path-is-inside"));
const rootDirectoryCache = new Map();
function getRootDirectoryPath(entryFsDirectory) {
    if (rootDirectoryCache.has(entryFsDirectory)) {
        return rootDirectoryCache.get(entryFsDirectory);
    }
    // maybe use vscode.workspace.getWorkspaceFolder(uri)?
    const rootDirectories = vscode.workspace.workspaceFolders.filter(workspaceFolder => {
        return path_is_inside_1.default(entryFsDirectory, workspaceFolder.uri.fsPath);
    });
    const rootDirectory = rootDirectories.sort((d1, d2) => {
        return d1.uri.fsPath.length - d2.uri.fsPath.length;
    })[0];
    if (rootDirectory && rootDirectory.uri.fsPath) {
        const rootDirectoryPath = rootDirectory.uri.fsPath;
        if (rootDirectoryPath) {
            rootDirectoryCache.set(entryFsDirectory, rootDirectoryPath);
            return rootDirectoryPath;
        }
    }
}
function handleDidChangeWorkspaceFolders() {
    rootDirectoryCache.clear();
}
exports.handleDidChangeWorkspaceFolders = handleDidChangeWorkspaceFolders;
function checkFsPath(entryFsDirectory, fsPath) {
    const rootDirectory = getRootDirectoryPath(entryFsDirectory);
    if (path.sep === '\\') {
        fsPath = path.normalize(fsPath);
    }
    if (!path_is_inside_1.default(fsPath, rootDirectory)) {
        return false;
    }
    return true;
}
exports.checkFsPath = checkFsPath;
class CustomError extends Error {
    constructor(...args) {
        super(...args);
        Object.setPrototypeOf(this, new.target.prototype);
        this.name = new.target.name;
        if (Error.captureStackTrace) {
            Error.captureStackTrace(this, new.target);
        }
    }
}
class PathAccessDeniedError extends CustomError {
    constructor(fsPath) {
        super(`Accessing ${fsPath} denied. This path is outside of your workspace folders. Please make sure you have all dependencies inside your workspace.`);
        this.fsPath = fsPath;
    }
}
exports.PathAccessDeniedError = PathAccessDeniedError;
//# sourceMappingURL=checkFsPath.js.map