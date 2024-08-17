"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const TextHelper_1 = require("./TextHelper");
class ResourceHelper {
    /**
     * Exclude all the none related project resource paths
     *
     * @param configInfo
     */
    static excludeResourcePaths(configInfo) {
        const lrKeys = Object.keys(configInfo.localizedResources);
        const resx = [];
        for (const key of lrKeys) {
            const value = configInfo.localizedResources[key];
            if (!value.includes("node_modules")) {
                resx.push({
                    key,
                    value
                });
            }
        }
        return resx;
    }
    /**
     * Search the word/key in the resource file
     *
     * @param contents
     * @param key
     */
    static getResourceValue(contents, key) {
        // Select the return object
        if (contents.includes(key)) {
            const regEx = new RegExp(`\\b${key}\\b`);
            const keyIdx = contents.search(regEx);
            if (keyIdx !== -1) {
                const colonIdx = contents.indexOf(":", keyIdx);
                const commaIdx = contents.indexOf(",", keyIdx);
                if (colonIdx !== -1 && commaIdx !== -1) {
                    let value = contents.substring((colonIdx + 1), commaIdx);
                    value = value.trim();
                    value = TextHelper_1.default.stripQuotes(value);
                    return value;
                }
            }
        }
        return null;
    }
    /**
     * Retrieve the key value pairs from the locale file contents
     *
     * @param fileContents
     */
    static getKeyValuePairs(fileContents) {
        let localeKeyValue = [];
        // Check if file contents were passed
        if (fileContents) {
            // Find the position of the return statement
            const fileLines = fileContents.split("\n");
            const returnIdx = fileLines.findIndex(line => {
                const matches = line.trim().match(/(^return|{$)/gi);
                return matches !== null && matches.length >= 2;
            });
            // Check if the index has been found
            if (returnIdx !== -1) {
                // Loop over all the lines
                let x = 0;
                for (const line of fileLines) {
                    if (x > returnIdx) {
                        const lineVal = line.trim();
                        // Get the colon location
                        const colonIdx = lineVal.indexOf(":");
                        if (colonIdx !== -1) {
                            let keyName = lineVal.substring(0, colonIdx);
                            keyName = keyName.trim();
                            let keyValue = lineVal.substring((colonIdx + 1));
                            keyValue = keyValue.trim();
                            keyValue = TextHelper_1.default.stripQuotes(keyValue);
                            // Add the key and value to the array
                            if (keyName && keyValue) {
                                localeKeyValue.push({
                                    key: TextHelper_1.default.stripQuotes(keyName),
                                    value: keyValue
                                });
                            }
                        }
                    }
                    x++;
                }
            }
        }
        return localeKeyValue;
    }
}
exports.default = ResourceHelper;
//# sourceMappingURL=ResourceHelper.js.map