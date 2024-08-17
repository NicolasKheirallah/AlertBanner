"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const vscode = require("vscode");
const path = require("path");
const ResourceHelper_1 = require("./ResourceHelper");
const CsvHelper_1 = require("./CsvHelper");
class ExportLocaleHelper {
    /**
     * Start the localization export to the CSV file
     *
     * @param err
     * @param csvData
     * @param localeFiles
     * @param csvLocation
     * @param delimiter
     * @param resourceName
     */
    static async startExport(csvData, localeFiles, csvLocation, delimiter, resourceName, useBom, useComment, useTimestamp) {
        // Start looping over the JS Locale files
        for (const localeFile of localeFiles) {
            const localeData = await vscode.workspace.openTextDocument(localeFile);
            if (localeData) {
                const keyValuePairs = ResourceHelper_1.default.getKeyValuePairs(localeData.getText());
                // Check if key value pairs have been retrieved
                if (keyValuePairs && keyValuePairs.length > 0) {
                    const fileName = path.basename(localeData.fileName);
                    const localeName = fileName.split('.').slice(0, -1).join('.');
                    // Start adding/updating the key and values to the CSV data
                    CsvHelper_1.default.updateData(csvData, keyValuePairs, localeName, resourceName, useComment, useTimestamp);
                }
            }
        }
        // Once all data has been processed, the CSV file can be created
        await CsvHelper_1.default.writeToCsvFile(csvLocation, csvData, delimiter, useBom);
    }
}
exports.default = ExportLocaleHelper;
//# sourceMappingURL=ExportLocaleHelper.js.map