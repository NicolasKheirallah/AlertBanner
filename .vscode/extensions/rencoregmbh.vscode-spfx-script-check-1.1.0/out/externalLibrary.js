"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const scriptChecker_1 = require("./scriptChecker");
class ExternalLibrary {
    /**
     * Update the external configuration
     * @param configJson
     * @param type
     * @param moduleName
     * @param url
     */
    static update(configJson, type, moduleName, url, globalName, scriptDependencies) {
        if (type === scriptChecker_1.ScriptType.module) {
            configJson.externals[moduleName] = url;
        }
        else {
            // Check if it is a plugin or a module
            if (globalName) {
                configJson.externals[moduleName] = {
                    path: url,
                    globalName: globalName,
                    globalDependencies: scriptDependencies.split(',')
                };
            }
            else {
                configJson.externals[moduleName] = {
                    path: url,
                    globalName: moduleName
                };
            }
        }
        return configJson;
    }
}
exports.default = ExternalLibrary;
//# sourceMappingURL=externalLibrary.js.map