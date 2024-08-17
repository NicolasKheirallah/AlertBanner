"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const vscode = require("vscode");
const Logging_1 = require("./Logging");
const ProjectFileHelper_1 = require("../helpers/ProjectFileHelper");
const ResourceHelper_1 = require("../helpers/ResourceHelper");
const CsvHelper_1 = require("../helpers/CsvHelper");
const ExportLocaleHelper_1 = require("../helpers/ExportLocaleHelper");
const ExtensionSettings_1 = require("../helpers/ExtensionSettings");
class CsvCommands {
    /**
     * Import locale labels and keys from a CSV file
     *
     * Logic
     * 1. get the CSV file ✅
     * 2. get the headers from the CSV file (key, locale, localizedResource) ✅
     * 3. ask which localized resource to use (if multiple are configured) ✅
     * 4. ask if the localized resource files can be overwritten - atm files will be created from the CSV file
     * 5. start (creating and) writing to the files ✅
     */
    static async import() {
        const config = vscode.workspace.getConfiguration(ExtensionSettings_1.CONFIG_KEY);
        // Retrieve the delimiter
        let delimiter = config.get(ExtensionSettings_1.CONFIG_CSV_DELIMITER);
        if (!delimiter) {
            delimiter = ";";
            Logging_1.default.warning(`The delimiter setting was empty, ";" will be used instead.`);
        }
        const useBom = !!config.get(ExtensionSettings_1.CONFIG_CSV_USE_BOM);
        const filePath = this.getCsvFilePath();
        if (filePath) {
            const csvData = await CsvHelper_1.default.openFile(filePath, delimiter, useBom);
            if (csvData) {
                this.initializeImport(csvData);
            }
            else {
                Logging_1.default.error(`The CSV/XLSX file could not be retrieved. Used file location: "${filePath}".`);
                return null;
            }
        }
    }
    /**
     * Export locale labels and keys to a CSV file
     *
     * Logic
     * 1. select the localized resource to output (if multiple) ✅
     * 2. get the localized resource files ✅
     * 3. fetch the csv file or create it if it doesn't exist ✅
     * 4. get all the headers or create them if they do not exist ✅
     * 5. add the keys and values ✅
     * 6. ask to override the data in the CSV file - atm the CSV becomes the master of the data
     */
    static async export(resxToUse = null) {
        try {
            const config = vscode.workspace.getConfiguration(ExtensionSettings_1.CONFIG_KEY);
            // Use the provided resource or ask which resource file to use
            const resources = resxToUse ? [resxToUse] : await this.getResourceToUse();
            if (resources && resources.length > 0) {
                for (const resource of resources) {
                    if (resource) {
                        let fileExtension = config.get(ExtensionSettings_1.CONFIG_FILE_EXTENSION);
                        if (!fileExtension) {
                            fileExtension = "js";
                        }
                        // Get all the localized resource files
                        const resourcePath = ProjectFileHelper_1.default.getResourcePath(resource);
                        let localeFiles = await vscode.workspace.findFiles(`${resourcePath}/*.${fileExtension}`);
                        if (!localeFiles || localeFiles.length === 0) {
                            Logging_1.default.error(`No locale files were found for the selected resource: ${resource.key}.`);
                        }
                        // Exclude the mystrings file
                        if (fileExtension === "ts") {
                            localeFiles = localeFiles.filter(f => !f.path.includes("mystrings.d.ts"));
                        }
                        let delimiter = config.get(ExtensionSettings_1.CONFIG_CSV_DELIMITER);
                        if (!delimiter) {
                            delimiter = ";";
                            Logging_1.default.warning(`The delimiter setting was empty, ";" will be used instead.`);
                        }
                        // Retrieve the settings for the extension
                        const csvFileLocation = config.get(ExtensionSettings_1.CONFIG_CSV_FILELOCATION);
                        if (!csvFileLocation) {
                            Logging_1.default.error(`The "spfxLocalization.csvFileLocation" configuration setting is not provided.`);
                            throw new Error(`The "spfxLocalization.csvFileLocation" configuration setting is not provided.`);
                        }
                        const useBom = !!config.get(ExtensionSettings_1.CONFIG_CSV_USE_BOM);
                        const useComment = !!config.get(ExtensionSettings_1.CONFIG_CSV_USE_COMMENT);
                        const useTimestamp = !!config.get(ExtensionSettings_1.CONFIG_CSV_USE_TIMESTAMP);
                        // Get the CSV file or create one
                        const filePath = await this.getCsvFilePath();
                        if (filePath) {
                            // Start the export
                            try {
                                let csvData = await CsvHelper_1.default.openFile(filePath, delimiter, useBom);
                                if (!csvData) {
                                    csvData = await CsvHelper_1.default.createCsvData(localeFiles, resource, csvFileLocation, fileExtension, useComment, useTimestamp);
                                }
                                ExportLocaleHelper_1.default.startExport(csvData, localeFiles, csvFileLocation, delimiter, resource.key, useBom, useComment, useTimestamp);
                            }
                            catch (err) {
                                Logging_1.default.error(`Unable to read the file ${filePath}. ${err}`);
                            }
                        }
                    }
                }
            }
        }
        catch (e) {
            // Nothing to do here
        }
    }
    /**
     * Ask for which component you want to export the localization
     */
    static async getResourceToUse() {
        const configInfo = await ProjectFileHelper_1.default.getConfig();
        if (configInfo && configInfo.localizedResources) {
            // Retrieve all the project related localized resources
            const resx = ResourceHelper_1.default.excludeResourcePaths(configInfo);
            if (resx && resx.length > 0) {
                // Take the default one to import
                let defaultResx = resx[0].key;
                if (resx.length > 1) {
                    // Add an option to import all
                    let opts = resx.map(r => r.key);
                    opts.push(ExtensionSettings_1.OPTION_EXPORT_ALL);
                    defaultResx = await vscode.window.showQuickPick(opts, {
                        placeHolder: "Specify for which resource file you want to perform the input.",
                        canPickMany: false
                    });
                }
                // Check if an option was provided
                if (defaultResx) {
                    if (defaultResx === ExtensionSettings_1.OPTION_EXPORT_ALL) {
                        // Return all resources
                        return resx;
                    }
                    else {
                        // Return only the one you choose
                        return resx.filter(r => r.key === defaultResx);
                    }
                }
            }
        }
        return null;
    }
    /**
     * Initialize the CSV data import
     *
     * @param err Parsing error
     * @param csvData Retrieved CSV data from the file
     */
    static async initializeImport(csvData) {
        // Check if the file contained content
        if (csvData && csvData.rowCount > 0) {
            // Retrieve the config data
            const configInfo = await ProjectFileHelper_1.default.getConfig();
            if (configInfo && configInfo.localizedResources) {
                // Retrieve all the project related localized resources
                const resx = ResourceHelper_1.default.excludeResourcePaths(configInfo);
                if (resx && resx.length > 0) {
                    // Take the default one to import
                    let defaultResx = resx[0].key;
                    if (resx.length > 1) {
                        // Add an option to import all
                        let opts = resx.map(r => r.key);
                        opts.push(ExtensionSettings_1.OPTION_IMPORT_ALL);
                        defaultResx = await vscode.window.showQuickPick(opts, {
                            placeHolder: "Specify for which resource file you want to perform the input.",
                            canPickMany: false
                        });
                    }
                    // Check if an option was provided
                    if (defaultResx) {
                        // Start the CSV data
                        await CsvHelper_1.default.startCsvImporting(csvData, defaultResx, resx);
                    }
                }
            }
            else {
                Logging_1.default.error(`SPFx project config file could not be retrieved`);
            }
        }
        else {
            Logging_1.default.warning(`The CSV/XLSX file is empty.`);
        }
    }
    /**
     * Get the CSV file
     *
     * @param needsToExists Specify if the file needs to exist
     */
    static getCsvFilePath() {
        // Retrieve the CSV file config value
        const csvFileLocation = vscode.workspace.getConfiguration(ExtensionSettings_1.CONFIG_KEY).get(ExtensionSettings_1.CONFIG_CSV_FILELOCATION);
        if (!csvFileLocation) {
            Logging_1.default.error(`The "spfxLocalization.csvFileLocation" configuration setting is not provided.`);
            return null;
        }
        // Get the absolute path for the file
        return ProjectFileHelper_1.default.getAbsPath(csvFileLocation);
    }
}
exports.default = CsvCommands;
//# sourceMappingURL=CsvCommands.js.map