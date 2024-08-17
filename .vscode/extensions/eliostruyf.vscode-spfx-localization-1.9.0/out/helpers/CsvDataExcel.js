"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.CsvDataExcel = void 0;
const ExcelJS = require("@nbelyh/exceljs");
class CsvDataExcel {
    constructor(data, name) {
        this.columnCount = -1;
        this.wb = new ExcelJS.Workbook();
        this.ws = this.wb.addWorksheet(name);
        if (data) {
            this.columnCount = data[0].length;
            this.ws.addRows(data);
        }
    }
    getData() {
        const result = [];
        const ws = this.ws;
        for (let r = 0; r < ws.rowCount; ++r) {
            result.push(Array(ws.columnCount));
            for (let c = 0; c < ws.columnCount; ++c) {
                result[r][c] = ws.getCell(r + 1, c + 1).value || '';
            }
        }
        return result;
    }
    getValue(r, c) {
        if (r < this.ws.rowCount && c < this.columnCount) {
            const cell = this.ws.getCell(r + 1, c + 1);
            return cell.value;
        }
        else {
            return '';
        }
    }
    setValue(r, c, v) {
        if (r < this.ws.rowCount && c < this.columnCount) {
            const cell = this.ws.getCell(r + 1, c + 1);
            cell.value = v;
        }
    }
    addRow(r) {
        this.ws.insertRow(r + 1, Array(this.colCount).join('.').split('.'));
    }
    get rowCount() {
        return this.ws.rowCount;
    }
    get colCount() {
        return this.ws.columnCount;
    }
    write(filePath, options) {
        return new Promise((resolve, reject) => {
            this.wb.xlsx.writeFile(filePath).then(() => resolve(true), err => reject(err));
        });
    }
    read(filePath, options) {
        return new Promise((resolve, reject) => {
            this.wb.xlsx.readFile(filePath).then(wb => {
                this.wb = wb;
                this.ws = wb.worksheets[0];
                this.columnCount = this.ws.columnCount;
                resolve(true);
            }, err => {
                reject(err);
            });
        });
    }
}
exports.CsvDataExcel = CsvDataExcel;
//# sourceMappingURL=CsvDataExcel.js.map