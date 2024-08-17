"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const vscode = require("vscode");
const path = require("path");
class ProjectFileHelper {
    /**
     * Fetch the project config file
     */
    static async getConfig(errorLog) {
        // Start the search for the loc folder in the project
        const configFileUrls = await vscode.workspace.findFiles('**/config/config.json', "**/node_modules/**", 1);
        if (!configFileUrls || configFileUrls.length === 0) {
            if (errorLog) {
                errorLog(`Solution config file could not be retrieved.`);
            }
            return null;
        }
        // Take the first config file
        const configFileUrl = configFileUrls[0];
        if (configFileUrl) {
            // Fetch the the config file contents
            const configFile = await vscode.workspace.openTextDocument(configFileUrl);
            if (!configFile) {
                if (errorLog) {
                    errorLog(`Could not read the config file.`);
                }
                return null;
            }
            // Get the file contents
            const contents = configFile.getText();
            if (!contents) {
                if (errorLog) {
                    errorLog(`Could not retrieve the file contents.`);
                }
                return null;
            }
            // Fetch the config information and check if localizedResources were defined
            const configInfo = JSON.parse(contents);
            return configInfo;
        }
        return null;
    }
    /**
     * Retrieve the resource path for the file
     *
     * @param resx
     */
    static getResourcePath(resx) {
        // Create the key in the localized resource file
        let resourcePath = resx.value.substring(0, resx.value.lastIndexOf('/'));
        // Check if the path starts with 'lib/', if this is the case, it needs to be changed to 'src/'
        if (resourcePath.startsWith("lib/")) {
            resourcePath = resourcePath.replace("lib/", "src/");
        }
        // vscode stopped finding localization files if they start with './'
        if (resourcePath.startsWith('./')) {
            resourcePath = resourcePath.replace("./", "");
        }
        return resourcePath;
    }
    /**
     * Get the absolute path for the file
     *
     * @param fileLocation
     */
    static getAbsPath(fileLocation) {
        // Create the file path
        const rootPath = vscode.workspace.rootPath || __dirname;
        return path.join(rootPath, fileLocation);
    }
}
exports.default = ProjectFileHelper;
//# sourceMappingURL=ProjectFileHelper.js.map