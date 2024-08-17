"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.CsvDataArray = void 0;
const fs = require("fs");
const csv_stringify_1 = require("csv-stringify");
const csv_parse_1 = require("csv-parse");
const ExtensionSettings_1 = require("./ExtensionSettings");
class CsvDataArray {
    constructor(data, name) {
        this.data = [];
        if (data) {
            this.data = data;
        }
    }
    getData() {
        return this.data;
    }
    getValue(r, c) {
        if (r < this.rowCount && c < this.colCount) {
            return this.data[r][c] || '';
        }
        else {
            return '';
        }
    }
    setValue(r, c, v) {
        if (r < this.rowCount && c < this.colCount) {
            this.data[r][c] = v;
        }
    }
    addRow(r) {
        this.data.splice(r, 0, Array(this.colCount).join('.').split('.'));
    }
    get rowCount() {
        return this.data && this.data.length;
    }
    get colCount() {
        return this.data && this.data.length && this.data[0] && this.data[0].length;
    }
    read(filePath, options) {
        return new Promise((resolve, reject) => {
            const input = fs.readFileSync(filePath);
            (0, csv_parse_1.parse)(input, { delimiter: options.delimiter, bom: options.bom }, (err, data) => {
                if (err) {
                    reject(err);
                }
                else {
                    this.data = data;
                    resolve(true);
                }
            });
        });
    }
    write(filePath, options) {
        return new Promise((resolve, reject) => {
            (0, csv_stringify_1.stringify)(this.data, { delimiter: options.delimiter }, (err, output) => {
                if (err) {
                    reject(err);
                }
                else {
                    if (output) {
                        const bom = options.bom ? ExtensionSettings_1.UTF8_BOM : '';
                        fs.writeFileSync(filePath, bom + output, { encoding: "utf8" });
                        resolve(true);
                    }
                    else {
                        resolve(false);
                    }
                }
            });
        });
    }
}
exports.CsvDataArray = CsvDataArray;
//# sourceMappingURL=CsvDataArray.js.map