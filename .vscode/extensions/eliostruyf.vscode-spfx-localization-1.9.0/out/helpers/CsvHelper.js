"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const fs = require("fs");
const path = require("path");
const Logging_1 = require("../commands/Logging");
const ImportLocaleHelper_1 = require("./ImportLocaleHelper");
const ProjectFileHelper_1 = require("./ProjectFileHelper");
const ExtensionSettings_1 = require("./ExtensionSettings");
const CsvHeaders_1 = require("../constants/CsvHeaders");
const CsvDataArray_1 = require("./CsvDataArray");
const CsvDataExcel_1 = require("./CsvDataExcel");
class CsvHelper {
    static async openFile(filePath, delimiter, bom) {
        if (!fs.existsSync(filePath)) {
            return null;
        }
        const result = this.isCsv(filePath) ? new CsvDataArray_1.CsvDataArray() : new CsvDataExcel_1.CsvDataExcel();
        await result.read(filePath, { delimiter, bom });
        return result;
    }
    /**
     * Start processing the CSV data
     *
     * @param csvData
     * @param impLocale
     * @param resx
     */
    static async startCsvImporting(csvData, impLocale, resx) {
        // Get the header information
        const csvHeaders = this.getHeaders(csvData);
        if (csvHeaders && csvHeaders.keyIdx >= 0) {
            // Process the CSV data
            const localeData = this.processCsvData(csvData, csvHeaders);
            if (localeData) {
                // Check which resx file needs to be imported
                if (impLocale === ExtensionSettings_1.OPTION_IMPORT_ALL) {
                    // Full import
                    for (const localeResx of resx) {
                        await ImportLocaleHelper_1.default.createLocaleFiles(localeResx, localeData);
                    }
                }
                else {
                    // Single import
                    await ImportLocaleHelper_1.default.createLocaleFiles(resx.find(r => r.key === impLocale), localeData);
                }
            }
        }
        else {
            Logging_1.default.error(`The header information is not correctly in place.`);
        }
    }
    /**
     * Create a new CSV file
     *
     * @param jsFiles
     * @param resource
     * @param csvFileLocation
     * @param delimiter
     * @param fileExtension
     */
    static async createCsvData(localeFiles, resource, csvFileLocation, fileExtension, useComment, useTimestamp) {
        const locales = localeFiles.map(f => {
            const filePath = f.path.substring(f.path.lastIndexOf("/") + 1);
            return filePath.replace(`.${fileExtension}`, "");
        });
        // Create the headers for the CSV file
        const headers = [CsvHeaders_1.KEY_HEADER, ...locales, resource.key];
        // add comment column if feature is enabled
        if (useComment) {
            headers.push(CsvHeaders_1.COMMENT_HEADER);
        }
        // add timestamp column if feature is enabled
        if (useTimestamp) {
            headers.push(CsvHeaders_1.TIMESTAMP_HEADER);
        }
        if (this.isCsv(csvFileLocation)) {
            return new CsvDataArray_1.CsvDataArray([headers], resource.key);
        }
        else {
            return new CsvDataExcel_1.CsvDataExcel([headers], resource.key);
        }
    }
    /**
     * Update the CSV data based on the retrieved locale pairs
     * @param csvData
     * @param keyValuePairs
     * @param localeName
     * @param resourceName
     * @param useComment
     * @param useTimestamp
     */
    static updateData(csvData, keyValuePairs, localeName, resourceName, useComment, useTimestamp) {
        const csvHeaders = this.getHeaders(csvData);
        const timestamp = useTimestamp ? this.addTimestamp() : '';
        if (csvHeaders && csvHeaders.keyIdx >= 0) {
            // Start looping over the keyValuePairs
            for (const keyValue of keyValuePairs) {
                const rowIdx = this.findRowForKey(csvData, keyValue.key, csvHeaders.keyIdx);
                // Check if rowIdx has been found
                if (rowIdx) {
                    // Update the row data
                    this.updateDataRow(csvData, rowIdx, csvHeaders, keyValue, localeName, resourceName, timestamp);
                }
                else {
                    // Key wasn't found, adding a new data row
                    this.addDataRow(csvData, csvHeaders, keyValue, localeName, resourceName, timestamp);
                }
            }
        }
    }
    static isCsv(filePath) {
        return path.extname(filePath).toLocaleLowerCase() === '.csv';
    }
    /**
     * Write the CSV data to the file
     *
     * @param fileLocation
     * @param csvData
     * @param delimiter
     * @param useBom
     */
    static async writeToCsvFile(fileLocation, csvData, delimiter, bom) {
        const filePath = ProjectFileHelper_1.default.getAbsPath(fileLocation);
        if (await csvData.write(filePath, { delimiter, bom })) {
            Logging_1.default.info(`Exported the locale data to the CSV/XLSX file.`);
        }
        else {
            Logging_1.default.error(`Something went wrong while writing to the CSV/XLSX file.`);
        }
    }
    /**
     * Update the current row data
     *
     * @param csvData
     * @param rowIndex
     * @param rowDefinition
     * @param keyValue
     * @param localeName
     * @param resourceName
     */
    static updateDataRow(csvData, rowIndex, csvHeaders, keyValue, localeName, resourceName, timestamp) {
        let rowModified = false;
        for (const locale of csvHeaders.localeIdx) {
            if (locale.key === localeName) {
                const existingValue = csvData.getValue(rowIndex, locale.idx);
                if (!existingValue) {
                    csvData.setValue(rowIndex, locale.idx, keyValue.value);
                    rowModified = true;
                }
                else {
                    if (existingValue !== keyValue.value) {
                        Logging_1.default.warning(`Ignoring overwritten ${keyValue.key} in ${localeName} '${keyValue.value}'. Keeping '${existingValue}'.`);
                    }
                }
                // rowModified = true;
            }
        }
        for (const resx of csvHeaders.resxNames) {
            if (resourceName === resx.key && !csvData.getValue(rowIndex, resx.idx)) {
                csvData.setValue(rowIndex, resx.idx, "x"); // Specify that the key is used in the specified resource
                rowModified = true;
            }
        }
        if (timestamp && csvHeaders.timestampIdx !== null && rowModified) {
            csvData.setValue(rowIndex, csvHeaders.timestampIdx, timestamp);
        }
    }
    /**
     * Add a new data row to the CSV data
     *
     * @param csvData
     * @param rowDefinition
     * @param keyValue
     * @param localeName
     * @param resourceName
     */
    static addDataRow(csvData, csvHeaders, keyValue, localeName, resourceName, timestamp) {
        if (csvHeaders.keyIdx >= 0) {
            // Add the new row
            const insertRow = this.findInsertRowForKey(csvData, keyValue.key, csvHeaders.keyIdx);
            csvData.addRow(insertRow);
            csvData.setValue(insertRow, csvHeaders.keyIdx, keyValue.key);
            for (const locale of csvHeaders.localeIdx) {
                if (locale.key === localeName) {
                    csvData.setValue(insertRow, locale.idx, keyValue.value); // Add the locale key to the CSV data
                }
            }
            for (const resx of csvHeaders.resxNames) {
                if (resourceName === resx.key) {
                    csvData.setValue(insertRow, resx.idx, "x"); // Specify that the key is used in the specified resource
                }
            }
            if (timestamp && csvHeaders.timestampIdx !== null) {
                csvData.setValue(insertRow, csvHeaders.timestampIdx, timestamp);
            }
        }
    }
    /**
     * Search for the corresponding key / row
     *
     * @param csvData
     * @param localeKey
     */
    static findRowForKey(csvData, localeKey, cellIdx) {
        for (let row = 0; row < csvData.rowCount; row++) {
            if (row && csvData.getValue(row, cellIdx) === localeKey) {
                return row;
            }
        }
        return null;
    }
    /**
     * Search for proper new row insert position (compare lines by keys, stop at the first which follows the key)
     *
     * @param csvData
     * @param localeKey
     */
    static findInsertRowForKey(csvData, localeKey, cellIdx) {
        let result = 1;
        for (let row = 1; row < csvData.rowCount; row++) {
            const rowKey = row && csvData.getValue(row, cellIdx);
            if (rowKey && rowKey.toLowerCase() < localeKey.toLowerCase()) {
                result = row + 1;
            }
        }
        return result;
    }
    /**
     * Process all the locale data from the CSV file
     *
     * @param csvData
     * @param csvHeaders
     */
    static processCsvData(csvData, csvHeaders) {
        if (csvHeaders.keyIdx >= 0) {
            const localeData = {};
            // Create all the required locale data
            for (const locale of csvHeaders.localeIdx) {
                localeData[locale.key] = [];
            }
            // Start looping over all the rows (filtering out the first row)
            for (let row = 1; row < csvData.rowCount; ++row) {
                // Loop over all the locales in the csv file
                for (const locale of csvHeaders.localeIdx) {
                    // Loop over the available resources
                    for (const resx of csvHeaders.resxNames) {
                        const resxValue = csvData.getValue(row, resx.idx) || null;
                        // Check if the label is for the current resource
                        if (resxValue && resxValue.toLowerCase() === "x") {
                            localeData[locale.key].push({
                                key: csvData.getValue(row, csvHeaders.keyIdx) || null,
                                label: csvData.getValue(row, locale.idx) || null,
                                comment: csvHeaders.commentIdx !== null ? csvData.getValue(row, csvHeaders.commentIdx) : null,
                                timestamp: csvHeaders.timestampIdx !== null ? csvData.getValue(row, csvHeaders.timestampIdx) : null,
                                resx: resx.key || null
                            });
                        }
                    }
                }
            }
            return localeData;
        }
        else {
            Logging_1.default.error(`The required ${CsvHeaders_1.KEY_HEADER} header was not found in the CSV file.`);
            return null;
        }
    }
    static getLocaleName(cell) {
        const trimmed = cell.toLowerCase().trim();
        if (trimmed.startsWith(CsvHeaders_1.LOCALE_HEADER.toLowerCase())) {
            return trimmed.replace(CsvHeaders_1.LOCALE_HEADER.toLowerCase(), "").trim();
        }
        else {
            const match = cell && /^[a-z]{2}-[a-z]{2}$/.exec(trimmed);
            return match && match[0];
        }
    }
    /**
     * Get the headers of the CSV file
     *
     * @param csvData
     */
    static getHeaders(csvData) {
        if (csvData && csvData.rowCount > 0) {
            const headerInfo = {
                keyIdx: -1,
                localeIdx: [],
                commentIdx: null,
                timestampIdx: null,
                resxNames: []
            };
            for (let i = 0; i <= csvData.colCount; i++) {
                // Get the cell
                const cell = csvData.getValue(0, i);
                if (cell) {
                    // Add the key index to the object
                    if (cell.toLowerCase() === CsvHeaders_1.KEY_HEADER.toLocaleLowerCase()) {
                        headerInfo.keyIdx = i;
                    }
                    else if (cell.toLowerCase() === CsvHeaders_1.COMMENT_HEADER.toLowerCase()) {
                        headerInfo.commentIdx = i;
                    }
                    else if (cell.toLowerCase() === CsvHeaders_1.TIMESTAMP_HEADER.toLowerCase()) {
                        headerInfo.timestampIdx = i;
                    }
                    else if (CsvHelper.getLocaleName(cell)) {
                        headerInfo.localeIdx.push({
                            key: CsvHelper.getLocaleName(cell),
                            idx: i
                        });
                    }
                    else {
                        headerInfo.resxNames.push({
                            key: cell,
                            idx: i
                        });
                    }
                }
            }
            return headerInfo;
        }
        return null;
    }
}
exports.default = CsvHelper;
/**
 * Use current timestamp as default comment for the new strings
 */
CsvHelper.addTimestamp = () => new Date().toISOString().split('.')[0];
//# sourceMappingURL=CsvHelper.js.map