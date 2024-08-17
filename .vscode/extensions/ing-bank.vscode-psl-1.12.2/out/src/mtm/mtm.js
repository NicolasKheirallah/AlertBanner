"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.MtmConnection = void 0;
const hostSocket_1 = require("./hostSocket");
const utils = require("./utils");
const fs = require("fs");
var ServiceClass;
(function (ServiceClass) {
    ServiceClass[ServiceClass["CONNECTION"] = 0] = "CONNECTION";
    ServiceClass[ServiceClass["MRPC"] = 3] = "MRPC";
    ServiceClass[ServiceClass["SQL"] = 5] = "SQL";
})(ServiceClass || (ServiceClass = {}));
class MtmConnection {
    constructor(serverType = 'SCA$IBS', encoding = 'utf8') {
        this.serverType = serverType;
        this.encoding = encoding;
        this.socket = new hostSocket_1.default();
        this.messageByte = String.fromCharCode(28);
        this.token = '';
        this.messageId = 0;
        this.maxRow = 30;
        this.isSql = false;
        this.recordCount = 0;
    }
    open(host, port, profileUsername, profilePassword) {
        return __awaiter(this, void 0, void 0, function* () {
            yield this.socket.connect(port, host);
            let prepareString = utils.connectionObject(profileUsername, profilePassword);
            let returnArray = yield this.execute({ serviceClass: ServiceClass.CONNECTION }, prepareString);
            this.token = returnArray;
        });
    }
    send(fileName) {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                let codeToken = yield this._send(fileName);
                let returnString = yield this.saveInProfile(fileName, codeToken);
                if (returnString !== '1') {
                    throw new Error(returnString.split('\r\n')[1]);
                }
                return returnString;
            }
            catch (err) {
                this.close();
                throw new Error(err.toString());
            }
        });
    }
    testCompile(fileName) {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                let codeToken = yield this._send(fileName);
                let returnString = yield this._testCompile(fileName, codeToken);
                return returnString;
            }
            catch (err) {
                this.close();
                throw new Error(err.toString());
            }
        });
    }
    get(fileName) {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                let returnString = yield this._get(fileName);
                return returnString;
            }
            catch (err) {
                this.close();
                throw new Error(err.toString());
            }
        });
    }
    compileAndLink(fileName) {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                let returnString = yield this._compileAndLink(fileName);
                return returnString;
            }
            catch (err) {
                this.close();
                throw new Error(err.toString());
            }
        });
    }
    runPsl(fileName) {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                let codeToken = yield this._send(fileName);
                let returnString = yield this._runPsl(codeToken);
                return returnString;
            }
            catch (err) {
                this.close();
                throw new Error(err.toString());
            }
        });
    }
    runCustom(fileName, mrpcID, request) {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                const codeToken = yield this._send(fileName);
                const returnString = yield this._runCustom(codeToken, mrpcID, request);
                return returnString;
            }
            catch (err) {
                this.close();
                throw new Error(err.toString());
            }
        });
    }
    close() {
        return __awaiter(this, void 0, void 0, function* () {
            this.socket.closeConnection();
            return this.socket.socket.destroyed;
        });
    }
    batchcomp(fileName) {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                let returnString = yield this.batchCompileAndLink(fileName);
                return returnString;
            }
            catch (err) {
                this.close();
                throw new Error(err.toString());
            }
        });
    }
    getTable(fileName) {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                this.isSql = false;
                let returnString = yield this._getTable(fileName);
                return returnString;
            }
            catch (err) {
                this.close();
                throw new Error(err.toString());
            }
        });
    }
    sqlQuery(query) {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                this.isSql = true;
                let returnString = yield this._sqlQuery(query);
                return returnString;
            }
            catch (err) {
                this.close();
                throw new Error(err.toString());
            }
        });
    }
    getPSLClasses() {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                let returnString = yield this._getPslClasses();
                return returnString;
            }
            catch (err) {
                this.close();
                throw new Error(err.toString());
            }
        });
    }
    _send(filename) {
        return __awaiter(this, void 0, void 0, function* () {
            let returnString;
            let fileString = (yield readFileAsync(filename, { encoding: this.encoding })).toString(this.encoding);
            let fileContentLength = fileString.length;
            let totalLoop = Math.ceil(fileContentLength / 1024);
            let codeToken = '';
            for (let i = 0; i < totalLoop; i++) {
                let partialString = fileString.slice(i * 1024, (i * 1024) + 1024);
                let withPipe = '';
                for (const char of partialString) {
                    withPipe += char.charCodeAt(0) + '|';
                }
                let prepareString = utils.initCodeObject(withPipe, codeToken);
                returnString = yield this.execute({ mrpcID: '121', serviceClass: ServiceClass.MRPC }, prepareString);
                codeToken = returnString;
            }
            let prepareString = utils.initCodeObject('', codeToken);
            returnString = yield this.execute({ mrpcID: '121', serviceClass: ServiceClass.MRPC }, prepareString);
            return returnString;
        });
    }
    saveInProfile(fileName, codeToken) {
        return __awaiter(this, void 0, void 0, function* () {
            let returnString;
            let fileDetails = utils.getObjectType(fileName);
            let prepareString = utils.saveObject(fileDetails.fileBaseName, codeToken, utils.getUserName());
            returnString = yield this.execute({ mrpcID: '121', serviceClass: ServiceClass.MRPC }, prepareString);
            return returnString;
        });
    }
    _testCompile(fileName, codeToken) {
        return __awaiter(this, void 0, void 0, function* () {
            let returnString;
            let fileDetails = utils.getObjectType(fileName);
            let prepareString = utils.testCompileObject(fileDetails.fileBaseName, codeToken);
            returnString = yield this.execute({ mrpcID: '121', serviceClass: ServiceClass.MRPC }, prepareString);
            return returnString;
        });
    }
    _get(fileName) {
        return __awaiter(this, void 0, void 0, function* () {
            let returnString;
            let fileDetails = utils.getObjectType(fileName);
            let prepareString = utils.initObject(fileDetails.fileId, fileDetails.fileName);
            returnString = yield this.execute({ mrpcID: '121', serviceClass: ServiceClass.MRPC }, prepareString);
            let codeToken = returnString.split('\r\n')[1];
            let hasMore = '1';
            returnString = '';
            while (hasMore === '1') {
                prepareString = utils.retObject(codeToken);
                let nextReturnString = yield this.execute({ mrpcID: '121', serviceClass: ServiceClass.MRPC }, prepareString);
                hasMore = nextReturnString.substr(0, 1);
                returnString = returnString + nextReturnString.substr(1, nextReturnString.length);
            }
            return returnString;
        });
    }
    _compileAndLink(fileName) {
        return __awaiter(this, void 0, void 0, function* () {
            let returnString;
            let fileDetails = utils.getObjectType(fileName);
            let prepareString = utils.preCompileObject(fileDetails.fileBaseName);
            let codeToken = yield this.execute({ mrpcID: '121', serviceClass: ServiceClass.MRPC }, prepareString);
            prepareString = utils.compileObject(codeToken);
            returnString = yield this.execute({ mrpcID: '121', serviceClass: ServiceClass.MRPC }, prepareString);
            return returnString;
        });
    }
    _runPsl(codeToken) {
        return __awaiter(this, void 0, void 0, function* () {
            let returnString;
            let prepareString = utils.pslRunObject(codeToken);
            returnString = yield this.execute({ mrpcID: '121', serviceClass: ServiceClass.MRPC }, prepareString);
            return returnString;
        });
    }
    _runCustom(codeToken, mrpcID, request) {
        return __awaiter(this, void 0, void 0, function* () {
            let returnString;
            let prepareString = utils.customRunObject(request, codeToken);
            returnString = yield this.execute({ mrpcID, serviceClass: ServiceClass.MRPC }, prepareString);
            return returnString;
        });
    }
    // Batch complie is not working since 81 is not fully exposed from profile
    batchCompileAndLink(fileName) {
        return __awaiter(this, void 0, void 0, function* () {
            let returnString;
            let fileDetails = utils.getObjectType(fileName);
            let dbtblTableName = utils.getDbtblInfo(fileDetails.fileId);
            let prepareString = utils.batchCompileObject(dbtblTableName, fileDetails.fileName);
            returnString = yield this.execute({ mrpcID: '121', serviceClass: ServiceClass.MRPC }, prepareString);
            return returnString;
        });
    }
    _getTable(fileName) {
        return __awaiter(this, void 0, void 0, function* () {
            let returnString;
            let columnList;
            let fileDetails = utils.getObjectType(fileName);
            let tableReturnString = fileDetails.fileBaseName + String.fromCharCode(1) + (yield this._get(fileName));
            let selectStatement = `SELECT COUNT(DI) FROM DBTBL1D WHERE FID='${fileDetails.fileName}' `;
            this.recordCount = Number(yield this._sqlQuery(selectStatement));
            selectStatement = `SELECT DI FROM DBTBL1D WHERE FID='${fileDetails.fileName}'`;
            returnString = yield this._sqlQuery(selectStatement);
            columnList = returnString.split('\r\n');
            returnString = tableReturnString;
            for (let i = 0; i < columnList.length; i++) {
                fileName = fileDetails.fileName + '-' + columnList[i] + '.COL';
                returnString = returnString + String.fromCharCode(0) + fileName + String.fromCharCode(1) + (yield this._get(fileName));
            }
            return returnString;
        });
    }
    _sqlQuery(selectQuery) {
        return __awaiter(this, void 0, void 0, function* () {
            selectQuery = selectQuery.toUpperCase();
            if (!selectQuery.startsWith('SELECT')) {
                throw new Error('Not a select query');
            }
            let cursorNumber = new Date().getTime().toString();
            let returnString = yield this.openSqlCursor(cursorNumber, selectQuery);
            returnString = yield this.fetchSqlCursor(cursorNumber);
            yield this.closeSqlCursor(cursorNumber);
            return returnString;
        });
    }
    openSqlCursor(cursorNumber, selectQuery) {
        return __awaiter(this, void 0, void 0, function* () {
            let openCursor = 'OPEN CURSOR ' + cursorNumber + ' AS ';
            let rows = '';
            let prepareString = utils.sqlObject(openCursor + selectQuery, rows);
            let returnString = yield this.execute({ serviceClass: ServiceClass.SQL }, prepareString);
            return returnString;
        });
    }
    fetchSqlCursor(cursorNumber) {
        return __awaiter(this, void 0, void 0, function* () {
            let fetchCursor = 'FETCH ' + cursorNumber;
            let rows = 'ROWS=' + this.maxRow;
            let prepareString = utils.sqlObject(fetchCursor, rows);
            let returnString = yield this.execute({ serviceClass: ServiceClass.SQL }, prepareString);
            let splitReturnSring = returnString.split(String.fromCharCode(0));
            let totalCount = Number(splitReturnSring[0]);
            returnString = splitReturnSring[1];
            if (this.isSql === false) {
                while ((totalCount < this.recordCount)) {
                    splitReturnSring = [];
                    let nextReturnString = yield this.execute({ serviceClass: ServiceClass.SQL }, prepareString);
                    splitReturnSring = nextReturnString.split(String.fromCharCode(0));
                    totalCount = totalCount + Number(splitReturnSring[0]);
                    returnString = returnString + '\r\n' + splitReturnSring[1];
                }
            }
            return returnString;
        });
    }
    closeSqlCursor(cursorNumber) {
        return __awaiter(this, void 0, void 0, function* () {
            let closeCursor = 'CLOSE ' + cursorNumber;
            let prepareString = utils.sqlObject(closeCursor, '');
            let returnString = yield this.execute({ serviceClass: ServiceClass.SQL }, prepareString);
            return returnString;
        });
    }
    _getPslClasses() {
        return __awaiter(this, void 0, void 0, function* () {
            let returnString;
            let prepareString = utils.getPslCls();
            returnString = yield this.execute({ mrpcID: '121', serviceClass: ServiceClass.MRPC }, prepareString);
            return returnString;
        });
    }
    execute(detail, prepareString) {
        return __awaiter(this, void 0, void 0, function* () {
            const sendingMessage = this.prepareSendingMessage(detail, prepareString);
            yield this.socket.send(sendingMessage);
            let message = yield this.socket.onceData();
            const { totalBytes, startByte } = utils.unpack(message);
            let messageLength = message.length;
            while (messageLength < totalBytes) {
                const nextMessage = yield this.socket.onceData();
                messageLength = messageLength + nextMessage.length;
                message = Buffer.concat([message, nextMessage], messageLength);
            }
            return (utils.parseResponse(detail.serviceClass, message.slice(startByte, message.length), this.encoding));
        });
    }
    prepareSendingMessage(detail, prepareString) {
        let tokenMessage = utils.tokenMessage(detail.serviceClass, this.token, this.messageId);
        if (detail.serviceClass === ServiceClass.MRPC) {
            let version = 1;
            prepareString = utils.mrpcMessage(detail.mrpcID, version.toString(), prepareString);
        }
        let sendingMessage = utils.sendingMessage(tokenMessage, prepareString);
        sendingMessage = this.serverType + this.messageByte + sendingMessage;
        sendingMessage = utils.pack(sendingMessage.length + 2) + sendingMessage;
        this.messageId++;
        return sendingMessage;
    }
}
exports.MtmConnection = MtmConnection;
function readFileAsync(file, options) {
    return new Promise((resolve, reject) => {
        fs.readFile(file, { encoding: null, flag: options.flag }, (err, data) => {
            if (err) {
                reject(err);
            }
            else {
                resolve(data);
            }
        });
    });
}
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoibXRtLmpzIiwic291cmNlUm9vdCI6IiIsInNvdXJjZXMiOlsiLi4vLi4vLi4vc3JjL210bS9tdG0udHMiXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6Ijs7Ozs7Ozs7Ozs7O0FBQUEsNkNBQXNDO0FBQ3RDLGlDQUFpQztBQUNqQyx5QkFBeUI7QUFFekIsSUFBSyxZQUlKO0FBSkQsV0FBSyxZQUFZO0lBQ2hCLDJEQUFjLENBQUE7SUFDZCwrQ0FBUSxDQUFBO0lBQ1IsNkNBQU8sQ0FBQTtBQUNSLENBQUMsRUFKSSxZQUFZLEtBQVosWUFBWSxRQUloQjtBQU9ELE1BQWEsYUFBYTtJQVV6QixZQUFvQixhQUFxQixTQUFTLEVBQVUsV0FBMkIsTUFBTTtRQUF6RSxlQUFVLEdBQVYsVUFBVSxDQUFvQjtRQUFVLGFBQVEsR0FBUixRQUFRLENBQXlCO1FBUnJGLFdBQU0sR0FBZSxJQUFJLG9CQUFVLEVBQUUsQ0FBQTtRQUNyQyxnQkFBVyxHQUFXLE1BQU0sQ0FBQyxZQUFZLENBQUMsRUFBRSxDQUFDLENBQUM7UUFDOUMsVUFBSyxHQUFXLEVBQUUsQ0FBQztRQUNuQixjQUFTLEdBQVcsQ0FBQyxDQUFDO1FBQ3RCLFdBQU0sR0FBVyxFQUFFLENBQUM7UUFDcEIsVUFBSyxHQUFZLEtBQUssQ0FBQztRQUN2QixnQkFBVyxHQUFXLENBQUMsQ0FBQztJQUVpRSxDQUFDO0lBRTVGLElBQUksQ0FBQyxJQUFZLEVBQUUsSUFBWSxFQUFFLGVBQXVCLEVBQUUsZUFBdUI7O1lBQ3RGLE1BQU0sSUFBSSxDQUFDLE1BQU0sQ0FBQyxPQUFPLENBQUMsSUFBSSxFQUFFLElBQUksQ0FBQyxDQUFDO1lBQ3RDLElBQUksYUFBYSxHQUFHLEtBQUssQ0FBQyxnQkFBZ0IsQ0FBQyxlQUFlLEVBQUUsZUFBZSxDQUFDLENBQUM7WUFDN0UsSUFBSSxXQUFXLEdBQUcsTUFBTSxJQUFJLENBQUMsT0FBTyxDQUFDLEVBQUUsWUFBWSxFQUFFLFlBQVksQ0FBQyxVQUFVLEVBQUUsRUFBRSxhQUFhLENBQUMsQ0FBQztZQUMvRixJQUFJLENBQUMsS0FBSyxHQUFHLFdBQVcsQ0FBQztRQUMxQixDQUFDO0tBQUE7SUFFSyxJQUFJLENBQUMsUUFBZ0I7O1lBQzFCLElBQUk7Z0JBQ0gsSUFBSSxTQUFTLEdBQUcsTUFBTSxJQUFJLENBQUMsS0FBSyxDQUFDLFFBQVEsQ0FBQyxDQUFBO2dCQUMxQyxJQUFJLFlBQVksR0FBRyxNQUFNLElBQUksQ0FBQyxhQUFhLENBQUMsUUFBUSxFQUFFLFNBQVMsQ0FBQyxDQUFBO2dCQUNoRSxJQUFJLFlBQVksS0FBSyxHQUFHLEVBQUU7b0JBQ3pCLE1BQU0sSUFBSSxLQUFLLENBQUMsWUFBWSxDQUFDLEtBQUssQ0FBQyxNQUFNLENBQUMsQ0FBQyxDQUFDLENBQUMsQ0FBQyxDQUFDO2lCQUMvQztnQkFDRCxPQUFPLFlBQVksQ0FBQTthQUNuQjtZQUNELE9BQU8sR0FBRyxFQUFFO2dCQUNYLElBQUksQ0FBQyxLQUFLLEVBQUUsQ0FBQztnQkFDYixNQUFNLElBQUksS0FBSyxDQUFDLEdBQUcsQ0FBQyxRQUFRLEVBQUUsQ0FBQyxDQUFDO2FBQ2hDO1FBQ0YsQ0FBQztLQUFBO0lBRUssV0FBVyxDQUFDLFFBQWdCOztZQUNqQyxJQUFJO2dCQUNILElBQUksU0FBUyxHQUFHLE1BQU0sSUFBSSxDQUFDLEtBQUssQ0FBQyxRQUFRLENBQUMsQ0FBQTtnQkFDMUMsSUFBSSxZQUFZLEdBQUcsTUFBTSxJQUFJLENBQUMsWUFBWSxDQUFDLFFBQVEsRUFBRSxTQUFTLENBQUMsQ0FBQTtnQkFDL0QsT0FBTyxZQUFZLENBQUE7YUFDbkI7WUFDRCxPQUFPLEdBQUcsRUFBRTtnQkFDWCxJQUFJLENBQUMsS0FBSyxFQUFFLENBQUM7Z0JBQ2IsTUFBTSxJQUFJLEtBQUssQ0FBQyxHQUFHLENBQUMsUUFBUSxFQUFFLENBQUMsQ0FBQzthQUNoQztRQUNGLENBQUM7S0FBQTtJQUVLLEdBQUcsQ0FBQyxRQUFnQjs7WUFDekIsSUFBSTtnQkFDSCxJQUFJLFlBQVksR0FBRyxNQUFNLElBQUksQ0FBQyxJQUFJLENBQUMsUUFBUSxDQUFDLENBQUE7Z0JBQzVDLE9BQU8sWUFBWSxDQUFBO2FBQ25CO1lBQ0QsT0FBTyxHQUFHLEVBQUU7Z0JBQ1gsSUFBSSxDQUFDLEtBQUssRUFBRSxDQUFDO2dCQUNiLE1BQU0sSUFBSSxLQUFLLENBQUMsR0FBRyxDQUFDLFFBQVEsRUFBRSxDQUFDLENBQUM7YUFDaEM7UUFDRixDQUFDO0tBQUE7SUFFSyxjQUFjLENBQUMsUUFBZ0I7O1lBQ3BDLElBQUk7Z0JBQ0gsSUFBSSxZQUFZLEdBQUcsTUFBTSxJQUFJLENBQUMsZUFBZSxDQUFDLFFBQVEsQ0FBQyxDQUFBO2dCQUN2RCxPQUFPLFlBQVksQ0FBQTthQUNuQjtZQUNELE9BQU8sR0FBRyxFQUFFO2dCQUNYLElBQUksQ0FBQyxLQUFLLEVBQUUsQ0FBQztnQkFDYixNQUFNLElBQUksS0FBSyxDQUFDLEdBQUcsQ0FBQyxRQUFRLEVBQUUsQ0FBQyxDQUFDO2FBQ2hDO1FBQ0YsQ0FBQztLQUFBO0lBRUssTUFBTSxDQUFDLFFBQWdCOztZQUM1QixJQUFJO2dCQUNILElBQUksU0FBUyxHQUFHLE1BQU0sSUFBSSxDQUFDLEtBQUssQ0FBQyxRQUFRLENBQUMsQ0FBQTtnQkFDMUMsSUFBSSxZQUFZLEdBQUcsTUFBTSxJQUFJLENBQUMsT0FBTyxDQUFDLFNBQVMsQ0FBQyxDQUFBO2dCQUNoRCxPQUFPLFlBQVksQ0FBQTthQUNuQjtZQUNELE9BQU8sR0FBRyxFQUFFO2dCQUNYLElBQUksQ0FBQyxLQUFLLEVBQUUsQ0FBQztnQkFDYixNQUFNLElBQUksS0FBSyxDQUFDLEdBQUcsQ0FBQyxRQUFRLEVBQUUsQ0FBQyxDQUFDO2FBQ2hDO1FBQ0YsQ0FBQztLQUFBO0lBRUssU0FBUyxDQUFDLFFBQWdCLEVBQUUsTUFBYyxFQUFFLE9BQWU7O1lBQ2hFLElBQUk7Z0JBQ0gsTUFBTSxTQUFTLEdBQUcsTUFBTSxJQUFJLENBQUMsS0FBSyxDQUFDLFFBQVEsQ0FBQyxDQUFDO2dCQUM3QyxNQUFNLFlBQVksR0FBRyxNQUFNLElBQUksQ0FBQyxVQUFVLENBQUMsU0FBUyxFQUFFLE1BQU0sRUFBRSxPQUFPLENBQUMsQ0FBQztnQkFDdkUsT0FBTyxZQUFZLENBQUM7YUFDcEI7WUFDRCxPQUFPLEdBQUcsRUFBRTtnQkFDWCxJQUFJLENBQUMsS0FBSyxFQUFFLENBQUM7Z0JBQ2IsTUFBTSxJQUFJLEtBQUssQ0FBQyxHQUFHLENBQUMsUUFBUSxFQUFFLENBQUMsQ0FBQzthQUNoQztRQUNGLENBQUM7S0FBQTtJQUVLLEtBQUs7O1lBQ1YsSUFBSSxDQUFDLE1BQU0sQ0FBQyxlQUFlLEVBQUUsQ0FBQztZQUM5QixPQUFPLElBQUksQ0FBQyxNQUFNLENBQUMsTUFBTSxDQUFDLFNBQVMsQ0FBQztRQUNyQyxDQUFDO0tBQUE7SUFFSyxTQUFTLENBQUMsUUFBZ0I7O1lBQy9CLElBQUk7Z0JBQ0gsSUFBSSxZQUFZLEdBQUcsTUFBTSxJQUFJLENBQUMsbUJBQW1CLENBQUMsUUFBUSxDQUFDLENBQUE7Z0JBQzNELE9BQU8sWUFBWSxDQUFBO2FBQ25CO1lBQ0QsT0FBTyxHQUFHLEVBQUU7Z0JBQ1gsSUFBSSxDQUFDLEtBQUssRUFBRSxDQUFDO2dCQUNiLE1BQU0sSUFBSSxLQUFLLENBQUMsR0FBRyxDQUFDLFFBQVEsRUFBRSxDQUFDLENBQUM7YUFDaEM7UUFDRixDQUFDO0tBQUE7SUFFSyxRQUFRLENBQUMsUUFBZ0I7O1lBQzlCLElBQUk7Z0JBQ0gsSUFBSSxDQUFDLEtBQUssR0FBRyxLQUFLLENBQUE7Z0JBQ2xCLElBQUksWUFBWSxHQUFHLE1BQU0sSUFBSSxDQUFDLFNBQVMsQ0FBQyxRQUFRLENBQUMsQ0FBQTtnQkFDakQsT0FBTyxZQUFZLENBQUE7YUFDbkI7WUFDRCxPQUFPLEdBQUcsRUFBRTtnQkFDWCxJQUFJLENBQUMsS0FBSyxFQUFFLENBQUM7Z0JBQ2IsTUFBTSxJQUFJLEtBQUssQ0FBQyxHQUFHLENBQUMsUUFBUSxFQUFFLENBQUMsQ0FBQzthQUNoQztRQUNGLENBQUM7S0FBQTtJQUVLLFFBQVEsQ0FBQyxLQUFhOztZQUMzQixJQUFJO2dCQUNILElBQUksQ0FBQyxLQUFLLEdBQUcsSUFBSSxDQUFBO2dCQUNqQixJQUFJLFlBQVksR0FBRyxNQUFNLElBQUksQ0FBQyxTQUFTLENBQUMsS0FBSyxDQUFDLENBQUE7Z0JBQzlDLE9BQU8sWUFBWSxDQUFBO2FBQ25CO1lBQ0QsT0FBTyxHQUFHLEVBQUU7Z0JBQ1gsSUFBSSxDQUFDLEtBQUssRUFBRSxDQUFDO2dCQUNiLE1BQU0sSUFBSSxLQUFLLENBQUMsR0FBRyxDQUFDLFFBQVEsRUFBRSxDQUFDLENBQUM7YUFDaEM7UUFDRixDQUFDO0tBQUE7SUFFSyxhQUFhOztZQUNsQixJQUFJO2dCQUNILElBQUksWUFBWSxHQUFHLE1BQU0sSUFBSSxDQUFDLGNBQWMsRUFBRSxDQUFDO2dCQUMvQyxPQUFPLFlBQVksQ0FBQzthQUNwQjtZQUNELE9BQU8sR0FBRyxFQUFFO2dCQUNYLElBQUksQ0FBQyxLQUFLLEVBQUUsQ0FBQztnQkFDYixNQUFNLElBQUksS0FBSyxDQUFDLEdBQUcsQ0FBQyxRQUFRLEVBQUUsQ0FBQyxDQUFDO2FBQ2hDO1FBQ0YsQ0FBQztLQUFBO0lBRWEsS0FBSyxDQUFDLFFBQWdCOztZQUNuQyxJQUFJLFlBQW9CLENBQUM7WUFDekIsSUFBSSxVQUFVLEdBQVcsQ0FBQyxNQUFNLGFBQWEsQ0FBQyxRQUFRLEVBQUUsRUFBQyxRQUFRLEVBQUUsSUFBSSxDQUFDLFFBQVEsRUFBQyxDQUFDLENBQUMsQ0FBQyxRQUFRLENBQUMsSUFBSSxDQUFDLFFBQVEsQ0FBQyxDQUFDO1lBQzVHLElBQUksaUJBQWlCLEdBQVcsVUFBVSxDQUFDLE1BQU0sQ0FBQztZQUNsRCxJQUFJLFNBQVMsR0FBVyxJQUFJLENBQUMsSUFBSSxDQUFDLGlCQUFpQixHQUFHLElBQUksQ0FBQyxDQUFDO1lBQzVELElBQUksU0FBUyxHQUFXLEVBQUUsQ0FBQztZQUMzQixLQUFLLElBQUksQ0FBQyxHQUFHLENBQUMsRUFBRSxDQUFDLEdBQUcsU0FBUyxFQUFFLENBQUMsRUFBRSxFQUFFO2dCQUNuQyxJQUFJLGFBQWEsR0FBVyxVQUFVLENBQUMsS0FBSyxDQUFDLENBQUMsR0FBRyxJQUFJLEVBQUUsQ0FBQyxDQUFDLEdBQUcsSUFBSSxDQUFDLEdBQUcsSUFBSSxDQUFDLENBQUM7Z0JBQzFFLElBQUksUUFBUSxHQUFXLEVBQUUsQ0FBQztnQkFDMUIsS0FBSyxNQUFNLElBQUksSUFBSSxhQUFhLEVBQUU7b0JBQ2pDLFFBQVEsSUFBSSxJQUFJLENBQUMsVUFBVSxDQUFDLENBQUMsQ0FBQyxHQUFHLEdBQUcsQ0FBQztpQkFDckM7Z0JBQ0QsSUFBSSxhQUFhLEdBQVcsS0FBSyxDQUFDLGNBQWMsQ0FBQyxRQUFRLEVBQUUsU0FBUyxDQUFDLENBQUE7Z0JBQ3JFLFlBQVksR0FBRyxNQUFNLElBQUksQ0FBQyxPQUFPLENBQUMsRUFBRSxNQUFNLEVBQUUsS0FBSyxFQUFFLFlBQVksRUFBRSxZQUFZLENBQUMsSUFBSSxFQUFFLEVBQUUsYUFBYSxDQUFDLENBQUE7Z0JBQ3BHLFNBQVMsR0FBRyxZQUFZLENBQUM7YUFDekI7WUFDRCxJQUFJLGFBQWEsR0FBVyxLQUFLLENBQUMsY0FBYyxDQUFDLEVBQUUsRUFBRSxTQUFTLENBQUMsQ0FBQTtZQUMvRCxZQUFZLEdBQUcsTUFBTSxJQUFJLENBQUMsT0FBTyxDQUFDLEVBQUUsTUFBTSxFQUFFLEtBQUssRUFBRSxZQUFZLEVBQUUsWUFBWSxDQUFDLElBQUksRUFBRSxFQUFFLGFBQWEsQ0FBQyxDQUFBO1lBQ3BHLE9BQU8sWUFBWSxDQUFDO1FBQ3JCLENBQUM7S0FBQTtJQUVhLGFBQWEsQ0FBQyxRQUFnQixFQUFFLFNBQWlCOztZQUM5RCxJQUFJLFlBQW9CLENBQUM7WUFDekIsSUFBSSxXQUFXLEdBQUcsS0FBSyxDQUFDLGFBQWEsQ0FBQyxRQUFRLENBQUMsQ0FBQztZQUNoRCxJQUFJLGFBQWEsR0FBRyxLQUFLLENBQUMsVUFBVSxDQUFDLFdBQVcsQ0FBQyxZQUFZLEVBQUUsU0FBUyxFQUFFLEtBQUssQ0FBQyxXQUFXLEVBQUUsQ0FBQyxDQUFBO1lBQzlGLFlBQVksR0FBRyxNQUFNLElBQUksQ0FBQyxPQUFPLENBQUMsRUFBRSxNQUFNLEVBQUUsS0FBSyxFQUFFLFlBQVksRUFBRSxZQUFZLENBQUMsSUFBSSxFQUFFLEVBQUUsYUFBYSxDQUFDLENBQUM7WUFDckcsT0FBTyxZQUFZLENBQUE7UUFDcEIsQ0FBQztLQUFBO0lBRWEsWUFBWSxDQUFDLFFBQWdCLEVBQUUsU0FBaUI7O1lBQzdELElBQUksWUFBb0IsQ0FBQztZQUN6QixJQUFJLFdBQVcsR0FBRyxLQUFLLENBQUMsYUFBYSxDQUFDLFFBQVEsQ0FBQyxDQUFDO1lBQ2hELElBQUksYUFBYSxHQUFHLEtBQUssQ0FBQyxpQkFBaUIsQ0FBQyxXQUFXLENBQUMsWUFBWSxFQUFFLFNBQVMsQ0FBQyxDQUFBO1lBQ2hGLFlBQVksR0FBRyxNQUFNLElBQUksQ0FBQyxPQUFPLENBQUMsRUFBRSxNQUFNLEVBQUUsS0FBSyxFQUFFLFlBQVksRUFBRSxZQUFZLENBQUMsSUFBSSxFQUFFLEVBQUUsYUFBYSxDQUFDLENBQUM7WUFDckcsT0FBTyxZQUFZLENBQUE7UUFDcEIsQ0FBQztLQUFBO0lBRWEsSUFBSSxDQUFDLFFBQWdCOztZQUNsQyxJQUFJLFlBQW9CLENBQUM7WUFDekIsSUFBSSxXQUFXLEdBQUcsS0FBSyxDQUFDLGFBQWEsQ0FBQyxRQUFRLENBQUMsQ0FBQztZQUNoRCxJQUFJLGFBQWEsR0FBRyxLQUFLLENBQUMsVUFBVSxDQUFDLFdBQVcsQ0FBQyxNQUFNLEVBQUUsV0FBVyxDQUFDLFFBQVEsQ0FBQyxDQUFBO1lBQzlFLFlBQVksR0FBRyxNQUFNLElBQUksQ0FBQyxPQUFPLENBQUMsRUFBRSxNQUFNLEVBQUUsS0FBSyxFQUFFLFlBQVksRUFBRSxZQUFZLENBQUMsSUFBSSxFQUFFLEVBQUUsYUFBYSxDQUFDLENBQUM7WUFDckcsSUFBSSxTQUFTLEdBQUcsWUFBWSxDQUFDLEtBQUssQ0FBQyxNQUFNLENBQUMsQ0FBQyxDQUFDLENBQUMsQ0FBQztZQUM5QyxJQUFJLE9BQU8sR0FBRyxHQUFHLENBQUE7WUFDakIsWUFBWSxHQUFHLEVBQUUsQ0FBQTtZQUNqQixPQUFPLE9BQU8sS0FBSyxHQUFHLEVBQUU7Z0JBQ3ZCLGFBQWEsR0FBRyxLQUFLLENBQUMsU0FBUyxDQUFDLFNBQVMsQ0FBQyxDQUFBO2dCQUMxQyxJQUFJLGdCQUFnQixHQUFHLE1BQU0sSUFBSSxDQUFDLE9BQU8sQ0FBQyxFQUFFLE1BQU0sRUFBRSxLQUFLLEVBQUUsWUFBWSxFQUFFLFlBQVksQ0FBQyxJQUFJLEVBQUUsRUFBRSxhQUFhLENBQUMsQ0FBQztnQkFDN0csT0FBTyxHQUFHLGdCQUFnQixDQUFDLE1BQU0sQ0FBQyxDQUFDLEVBQUUsQ0FBQyxDQUFDLENBQUM7Z0JBQ3hDLFlBQVksR0FBRyxZQUFZLEdBQUcsZ0JBQWdCLENBQUMsTUFBTSxDQUFDLENBQUMsRUFBRSxnQkFBZ0IsQ0FBQyxNQUFNLENBQUMsQ0FBQzthQUNsRjtZQUNELE9BQU8sWUFBWSxDQUFBO1FBQ3BCLENBQUM7S0FBQTtJQUVhLGVBQWUsQ0FBQyxRQUFnQjs7WUFDN0MsSUFBSSxZQUFvQixDQUFDO1lBQ3pCLElBQUksV0FBVyxHQUFHLEtBQUssQ0FBQyxhQUFhLENBQUMsUUFBUSxDQUFDLENBQUM7WUFDaEQsSUFBSSxhQUFhLEdBQUcsS0FBSyxDQUFDLGdCQUFnQixDQUFDLFdBQVcsQ0FBQyxZQUFZLENBQUMsQ0FBQTtZQUNwRSxJQUFJLFNBQVMsR0FBRyxNQUFNLElBQUksQ0FBQyxPQUFPLENBQUMsRUFBRSxNQUFNLEVBQUUsS0FBSyxFQUFFLFlBQVksRUFBRSxZQUFZLENBQUMsSUFBSSxFQUFFLEVBQUUsYUFBYSxDQUFDLENBQUM7WUFDdEcsYUFBYSxHQUFHLEtBQUssQ0FBQyxhQUFhLENBQUMsU0FBUyxDQUFDLENBQUE7WUFDOUMsWUFBWSxHQUFHLE1BQU0sSUFBSSxDQUFDLE9BQU8sQ0FBQyxFQUFFLE1BQU0sRUFBRSxLQUFLLEVBQUUsWUFBWSxFQUFFLFlBQVksQ0FBQyxJQUFJLEVBQUUsRUFBRSxhQUFhLENBQUMsQ0FBQztZQUNyRyxPQUFPLFlBQVksQ0FBQztRQUNyQixDQUFDO0tBQUE7SUFFYSxPQUFPLENBQUMsU0FBaUI7O1lBQ3RDLElBQUksWUFBb0IsQ0FBQztZQUN6QixJQUFJLGFBQWEsR0FBRyxLQUFLLENBQUMsWUFBWSxDQUFDLFNBQVMsQ0FBQyxDQUFBO1lBQ2pELFlBQVksR0FBRyxNQUFNLElBQUksQ0FBQyxPQUFPLENBQUMsRUFBRSxNQUFNLEVBQUUsS0FBSyxFQUFFLFlBQVksRUFBRSxZQUFZLENBQUMsSUFBSSxFQUFFLEVBQUUsYUFBYSxDQUFDLENBQUM7WUFDckcsT0FBTyxZQUFZLENBQUM7UUFDckIsQ0FBQztLQUFBO0lBRWEsVUFBVSxDQUFDLFNBQWlCLEVBQUUsTUFBYyxFQUFFLE9BQWU7O1lBQzFFLElBQUksWUFBb0IsQ0FBQztZQUN6QixJQUFJLGFBQWEsR0FBRyxLQUFLLENBQUMsZUFBZSxDQUFDLE9BQU8sRUFBRSxTQUFTLENBQUMsQ0FBQztZQUM5RCxZQUFZLEdBQUcsTUFBTSxJQUFJLENBQUMsT0FBTyxDQUFDLEVBQUUsTUFBTSxFQUFFLFlBQVksRUFBRSxZQUFZLENBQUMsSUFBSSxFQUFFLEVBQUUsYUFBYSxDQUFDLENBQUM7WUFDOUYsT0FBTyxZQUFZLENBQUM7UUFDckIsQ0FBQztLQUFBO0lBRUQsMEVBQTBFO0lBQzVELG1CQUFtQixDQUFDLFFBQWdCOztZQUNqRCxJQUFJLFlBQW9CLENBQUM7WUFDekIsSUFBSSxXQUFXLEdBQUcsS0FBSyxDQUFDLGFBQWEsQ0FBQyxRQUFRLENBQUMsQ0FBQztZQUNoRCxJQUFJLGNBQWMsR0FBRyxLQUFLLENBQUMsWUFBWSxDQUFDLFdBQVcsQ0FBQyxNQUFNLENBQUMsQ0FBQztZQUM1RCxJQUFJLGFBQWEsR0FBRyxLQUFLLENBQUMsa0JBQWtCLENBQUMsY0FBYyxFQUFFLFdBQVcsQ0FBQyxRQUFRLENBQUMsQ0FBQTtZQUNsRixZQUFZLEdBQUcsTUFBTSxJQUFJLENBQUMsT0FBTyxDQUFDLEVBQUUsTUFBTSxFQUFFLEtBQUssRUFBRSxZQUFZLEVBQUUsWUFBWSxDQUFDLElBQUksRUFBRSxFQUFFLGFBQWEsQ0FBQyxDQUFDO1lBQ3JHLE9BQU8sWUFBWSxDQUFDO1FBQ3JCLENBQUM7S0FBQTtJQUVhLFNBQVMsQ0FBQyxRQUFnQjs7WUFDdkMsSUFBSSxZQUFvQixDQUFDO1lBQ3pCLElBQUksVUFBb0IsQ0FBQztZQUN6QixJQUFJLFdBQVcsR0FBRyxLQUFLLENBQUMsYUFBYSxDQUFDLFFBQVEsQ0FBQyxDQUFDO1lBQ2hELElBQUksaUJBQWlCLEdBQUcsV0FBVyxDQUFDLFlBQVksR0FBRyxNQUFNLENBQUMsWUFBWSxDQUFDLENBQUMsQ0FBQyxJQUFHLE1BQU0sSUFBSSxDQUFDLElBQUksQ0FBQyxRQUFRLENBQUMsQ0FBQSxDQUFBO1lBQ3JHLElBQUksZUFBZSxHQUFHLDRDQUE0QyxXQUFXLENBQUMsUUFBUSxJQUFJLENBQUM7WUFDM0YsSUFBSSxDQUFDLFdBQVcsR0FBRyxNQUFNLENBQUMsTUFBTSxJQUFJLENBQUMsU0FBUyxDQUFDLGVBQWUsQ0FBQyxDQUFDLENBQUE7WUFDaEUsZUFBZSxHQUFHLHFDQUFxQyxXQUFXLENBQUMsUUFBUSxHQUFHLENBQUM7WUFDL0UsWUFBWSxHQUFHLE1BQU0sSUFBSSxDQUFDLFNBQVMsQ0FBQyxlQUFlLENBQUMsQ0FBQTtZQUNwRCxVQUFVLEdBQUcsWUFBWSxDQUFDLEtBQUssQ0FBQyxNQUFNLENBQUMsQ0FBQztZQUN4QyxZQUFZLEdBQUcsaUJBQWlCLENBQUE7WUFDaEMsS0FBSyxJQUFJLENBQUMsR0FBRyxDQUFDLEVBQUUsQ0FBQyxHQUFHLFVBQVUsQ0FBQyxNQUFNLEVBQUUsQ0FBQyxFQUFFLEVBQUU7Z0JBQzNDLFFBQVEsR0FBRyxXQUFXLENBQUMsUUFBUSxHQUFHLEdBQUcsR0FBRyxVQUFVLENBQUMsQ0FBQyxDQUFDLEdBQUcsTUFBTSxDQUFBO2dCQUM5RCxZQUFZLEdBQUcsWUFBWSxHQUFHLE1BQU0sQ0FBQyxZQUFZLENBQUMsQ0FBQyxDQUFDLEdBQUcsUUFBUSxHQUFHLE1BQU0sQ0FBQyxZQUFZLENBQUMsQ0FBQyxDQUFDLElBQUcsTUFBTSxJQUFJLENBQUMsSUFBSSxDQUFDLFFBQVEsQ0FBQyxDQUFBLENBQUE7YUFDcEg7WUFDRCxPQUFPLFlBQVksQ0FBQztRQUNyQixDQUFDO0tBQUE7SUFFYSxTQUFTLENBQUMsV0FBbUI7O1lBQzFDLFdBQVcsR0FBRyxXQUFXLENBQUMsV0FBVyxFQUFFLENBQUE7WUFDdkMsSUFBSSxDQUFDLFdBQVcsQ0FBQyxVQUFVLENBQUMsUUFBUSxDQUFDLEVBQUU7Z0JBQ3RDLE1BQU0sSUFBSSxLQUFLLENBQUMsb0JBQW9CLENBQUMsQ0FBQzthQUN0QztZQUNELElBQUksWUFBWSxHQUFHLElBQUksSUFBSSxFQUFFLENBQUMsT0FBTyxFQUFFLENBQUMsUUFBUSxFQUFFLENBQUE7WUFDbEQsSUFBSSxZQUFZLEdBQUcsTUFBTSxJQUFJLENBQUMsYUFBYSxDQUFDLFlBQVksRUFBRSxXQUFXLENBQUMsQ0FBQTtZQUN0RSxZQUFZLEdBQUcsTUFBTSxJQUFJLENBQUMsY0FBYyxDQUFDLFlBQVksQ0FBQyxDQUFBO1lBQ3RELE1BQU0sSUFBSSxDQUFDLGNBQWMsQ0FBQyxZQUFZLENBQUMsQ0FBQTtZQUN2QyxPQUFPLFlBQVksQ0FBQztRQUNyQixDQUFDO0tBQUE7SUFFYSxhQUFhLENBQUMsWUFBb0IsRUFBRSxXQUFtQjs7WUFDcEUsSUFBSSxVQUFVLEdBQUcsY0FBYyxHQUFHLFlBQVksR0FBRyxNQUFNLENBQUM7WUFDeEQsSUFBSSxJQUFJLEdBQUcsRUFBRSxDQUFDO1lBQ2QsSUFBSSxhQUFhLEdBQUcsS0FBSyxDQUFDLFNBQVMsQ0FBQyxVQUFVLEdBQUcsV0FBVyxFQUFFLElBQUksQ0FBQyxDQUFBO1lBQ25FLElBQUksWUFBWSxHQUFHLE1BQU0sSUFBSSxDQUFDLE9BQU8sQ0FBQyxFQUFFLFlBQVksRUFBRSxZQUFZLENBQUMsR0FBRyxFQUFFLEVBQUUsYUFBYSxDQUFDLENBQUM7WUFDekYsT0FBTyxZQUFZLENBQUE7UUFDcEIsQ0FBQztLQUFBO0lBRWEsY0FBYyxDQUFDLFlBQW9COztZQUNoRCxJQUFJLFdBQVcsR0FBRyxRQUFRLEdBQUcsWUFBWSxDQUFDO1lBQzFDLElBQUksSUFBSSxHQUFHLE9BQU8sR0FBRyxJQUFJLENBQUMsTUFBTSxDQUFDO1lBQ2pDLElBQUksYUFBYSxHQUFHLEtBQUssQ0FBQyxTQUFTLENBQUMsV0FBVyxFQUFFLElBQUksQ0FBQyxDQUFBO1lBQ3RELElBQUksWUFBWSxHQUFHLE1BQU0sSUFBSSxDQUFDLE9BQU8sQ0FBQyxFQUFFLFlBQVksRUFBRSxZQUFZLENBQUMsR0FBRyxFQUFFLEVBQUUsYUFBYSxDQUFDLENBQUM7WUFDekYsSUFBSSxnQkFBZ0IsR0FBYSxZQUFZLENBQUMsS0FBSyxDQUFDLE1BQU0sQ0FBQyxZQUFZLENBQUMsQ0FBQyxDQUFDLENBQUMsQ0FBQTtZQUMzRSxJQUFJLFVBQVUsR0FBRyxNQUFNLENBQUMsZ0JBQWdCLENBQUMsQ0FBQyxDQUFDLENBQUMsQ0FBQztZQUM3QyxZQUFZLEdBQUcsZ0JBQWdCLENBQUMsQ0FBQyxDQUFDLENBQUM7WUFDbkMsSUFBSSxJQUFJLENBQUMsS0FBSyxLQUFLLEtBQUssRUFBRTtnQkFDekIsT0FBTyxDQUFDLFVBQVUsR0FBRyxJQUFJLENBQUMsV0FBVyxDQUFDLEVBQUU7b0JBQ3ZDLGdCQUFnQixHQUFHLEVBQUUsQ0FBQztvQkFDdEIsSUFBSSxnQkFBZ0IsR0FBRyxNQUFNLElBQUksQ0FBQyxPQUFPLENBQUMsRUFBRSxZQUFZLEVBQUUsWUFBWSxDQUFDLEdBQUcsRUFBRSxFQUFFLGFBQWEsQ0FBQyxDQUFDO29CQUM3RixnQkFBZ0IsR0FBRyxnQkFBZ0IsQ0FBQyxLQUFLLENBQUMsTUFBTSxDQUFDLFlBQVksQ0FBQyxDQUFDLENBQUMsQ0FBQyxDQUFBO29CQUNqRSxVQUFVLEdBQUcsVUFBVSxHQUFHLE1BQU0sQ0FBQyxnQkFBZ0IsQ0FBQyxDQUFDLENBQUMsQ0FBQyxDQUFDO29CQUN0RCxZQUFZLEdBQUcsWUFBWSxHQUFHLE1BQU0sR0FBRyxnQkFBZ0IsQ0FBQyxDQUFDLENBQUMsQ0FBQTtpQkFDMUQ7YUFDRDtZQUNELE9BQU8sWUFBWSxDQUFBO1FBQ3BCLENBQUM7S0FBQTtJQUVhLGNBQWMsQ0FBQyxZQUFvQjs7WUFDaEQsSUFBSSxXQUFXLEdBQUcsUUFBUSxHQUFHLFlBQVksQ0FBQztZQUMxQyxJQUFJLGFBQWEsR0FBRyxLQUFLLENBQUMsU0FBUyxDQUFDLFdBQVcsRUFBRSxFQUFFLENBQUMsQ0FBQTtZQUNwRCxJQUFJLFlBQVksR0FBRyxNQUFNLElBQUksQ0FBQyxPQUFPLENBQUMsRUFBRSxZQUFZLEVBQUUsWUFBWSxDQUFDLEdBQUcsRUFBRSxFQUFFLGFBQWEsQ0FBQyxDQUFDO1lBQ3pGLE9BQU8sWUFBWSxDQUFBO1FBQ3BCLENBQUM7S0FBQTtJQUVhLGNBQWM7O1lBQzNCLElBQUksWUFBb0IsQ0FBQztZQUN6QixJQUFJLGFBQWEsR0FBRyxLQUFLLENBQUMsU0FBUyxFQUFFLENBQUE7WUFDckMsWUFBWSxHQUFHLE1BQU0sSUFBSSxDQUFDLE9BQU8sQ0FBQyxFQUFFLE1BQU0sRUFBRSxLQUFLLEVBQUUsWUFBWSxFQUFFLFlBQVksQ0FBQyxJQUFJLEVBQUUsRUFBRSxhQUFhLENBQUMsQ0FBQztZQUNyRyxPQUFPLFlBQVksQ0FBQztRQUNyQixDQUFDO0tBQUE7SUFFYSxPQUFPLENBQUMsTUFBcUIsRUFBRSxhQUFxQjs7WUFDakUsTUFBTSxjQUFjLEdBQUcsSUFBSSxDQUFDLHFCQUFxQixDQUFDLE1BQU0sRUFBRSxhQUFhLENBQUMsQ0FBQztZQUN6RSxNQUFNLElBQUksQ0FBQyxNQUFNLENBQUMsSUFBSSxDQUFDLGNBQWMsQ0FBQyxDQUFDO1lBQ3ZDLElBQUksT0FBTyxHQUFHLE1BQU0sSUFBSSxDQUFDLE1BQU0sQ0FBQyxRQUFRLEVBQUUsQ0FBQztZQUMzQyxNQUFNLEVBQUUsVUFBVSxFQUFFLFNBQVMsRUFBRSxHQUFHLEtBQUssQ0FBQyxNQUFNLENBQUMsT0FBTyxDQUFDLENBQUM7WUFDeEQsSUFBSSxhQUFhLEdBQUcsT0FBTyxDQUFDLE1BQU0sQ0FBQztZQUVuQyxPQUFPLGFBQWEsR0FBRyxVQUFVLEVBQUU7Z0JBQ2xDLE1BQU0sV0FBVyxHQUFHLE1BQU0sSUFBSSxDQUFDLE1BQU0sQ0FBQyxRQUFRLEVBQUUsQ0FBQztnQkFDakQsYUFBYSxHQUFHLGFBQWEsR0FBRyxXQUFXLENBQUMsTUFBTSxDQUFDO2dCQUNuRCxPQUFPLEdBQUcsTUFBTSxDQUFDLE1BQU0sQ0FBQyxDQUFDLE9BQU8sRUFBRSxXQUFXLENBQUMsRUFBRSxhQUFhLENBQUMsQ0FBQzthQUMvRDtZQUNELE9BQU8sQ0FBQyxLQUFLLENBQUMsYUFBYSxDQUFDLE1BQU0sQ0FBQyxZQUFZLEVBQUUsT0FBTyxDQUFDLEtBQUssQ0FBQyxTQUFTLEVBQUUsT0FBTyxDQUFDLE1BQU0sQ0FBQyxFQUFFLElBQUksQ0FBQyxRQUFRLENBQUMsQ0FBQyxDQUFDO1FBQzVHLENBQUM7S0FBQTtJQUVPLHFCQUFxQixDQUFDLE1BQXFCLEVBQUUsYUFBcUI7UUFDekUsSUFBSSxZQUFZLEdBQUcsS0FBSyxDQUFDLFlBQVksQ0FBQyxNQUFNLENBQUMsWUFBWSxFQUFFLElBQUksQ0FBQyxLQUFLLEVBQUUsSUFBSSxDQUFDLFNBQVMsQ0FBQyxDQUFDO1FBQ3ZGLElBQUksTUFBTSxDQUFDLFlBQVksS0FBSyxZQUFZLENBQUMsSUFBSSxFQUFFO1lBQzlDLElBQUksT0FBTyxHQUFXLENBQUMsQ0FBQztZQUN4QixhQUFhLEdBQUcsS0FBSyxDQUFDLFdBQVcsQ0FBQyxNQUFNLENBQUMsTUFBTSxFQUFFLE9BQU8sQ0FBQyxRQUFRLEVBQUUsRUFBRSxhQUFhLENBQUMsQ0FBQTtTQUNuRjtRQUNELElBQUksY0FBYyxHQUFHLEtBQUssQ0FBQyxjQUFjLENBQUMsWUFBWSxFQUFFLGFBQWEsQ0FBQyxDQUFDO1FBQ3ZFLGNBQWMsR0FBRyxJQUFJLENBQUMsVUFBVSxHQUFHLElBQUksQ0FBQyxXQUFXLEdBQUcsY0FBYyxDQUFBO1FBQ3BFLGNBQWMsR0FBRyxLQUFLLENBQUMsSUFBSSxDQUFDLGNBQWMsQ0FBQyxNQUFNLEdBQUcsQ0FBQyxDQUFDLEdBQUcsY0FBYyxDQUFDO1FBQ3hFLElBQUksQ0FBQyxTQUFTLEVBQUUsQ0FBQztRQUNqQixPQUFPLGNBQWMsQ0FBQztJQUN2QixDQUFDO0NBQ0Q7QUExVUQsc0NBMFVDO0FBRUQsU0FBUyxhQUFhLENBQUMsSUFBWSxFQUFFLE9BQTRDO0lBQ2hGLE9BQU8sSUFBSSxPQUFPLENBQUMsQ0FBQyxPQUFPLEVBQUUsTUFBTSxFQUFFLEVBQUU7UUFDdEMsRUFBRSxDQUFDLFFBQVEsQ0FBQyxJQUFJLEVBQUUsRUFBQyxRQUFRLEVBQUUsSUFBSSxFQUFFLElBQUksRUFBRSxPQUFPLENBQUMsSUFBSSxFQUFDLEVBQUUsQ0FBQyxHQUFHLEVBQUUsSUFBSSxFQUFFLEVBQUU7WUFDckUsSUFBSSxHQUFHLEVBQUU7Z0JBQ1IsTUFBTSxDQUFDLEdBQUcsQ0FBQyxDQUFDO2FBQ1o7aUJBQU07Z0JBQ04sT0FBTyxDQUFDLElBQUksQ0FBQyxDQUFDO2FBQ2Q7UUFDRixDQUFDLENBQUMsQ0FBQztJQUNKLENBQUMsQ0FBQyxDQUFDO0FBQ0osQ0FBQyJ9