"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.unpack = exports.pack = exports.getDbtblInfo = exports.getObjectType = exports.getUserName = exports.getPslCls = exports.batchCompileObject = exports.sqlObject = exports.customRunObject = exports.pslRunObject = exports.compileObject = exports.preCompileObject = exports.retObject = exports.initObject = exports.testCompileObject = exports.initCodeObject = exports.saveObject = exports.checkObject = exports.connectionObject = exports.tokenMessage = exports.mrpcMessage = exports.sendingMessage = exports.parseResponse = exports.lv2vFormat = exports.lvFormat = exports.v2lvFormat = exports.extensionToDescription = void 0;
const os = require("os");
const path = require("path");
exports.extensionToDescription = {
    'BATCH': 'Batch',
    'COL': 'Column',
    'DAT': 'Data',
    'FKY': 'Foreign Key',
    // 'G': 'Global',
    'IDX': 'Index',
    'JFD': 'Journal',
    'm': 'M routine',
    'PPL': 'Pre Post Lib',
    'PROC': 'Procedure',
    'properties': 'properties file',
    'PSL': 'psl File',
    'psl': 'psl File',
    'pslx': 'pslx File',
    'pslxtra': 'pslxtra File',
    'psql': 'PSQLScript',
    'serial': 'Serialized Data',
    'QRY': 'Query',
    'RPT': 'Report',
    'SCR': 'Screen',
    // TABLE not supported
    // 'TABLE': 'Complete Table',
    'TBL': 'Table',
    'TRIG': 'Trigger',
};
function v2lvFormat(messageValue) {
    let lvMessage = '';
    if (messageValue.length !== 0) {
        messageValue.forEach(messageString => {
            lvMessage = lvMessage + lvFormat(messageString);
        });
    }
    return lvMessage;
}
exports.v2lvFormat = v2lvFormat;
function lvFormat(messagString) {
    let returnLvFormat = '';
    let lvArray = [];
    let messageLength = messagString.length;
    let splitBytes;
    if (messageLength < 255) {
        splitBytes = 1;
    }
    else if (messageLength < 65534) {
        splitBytes = 2;
    }
    else if (messageLength < 16777213) {
        splitBytes = 3;
    }
    else {
        splitBytes = 4;
    }
    messageLength = splitBytes + messageLength;
    if (messageLength > 255) {
        for (let loop = 0; loop < splitBytes; loop++) {
            lvArray.push(messageLength % 256);
            messageLength = Math.trunc(messageLength / 256);
        }
        returnLvFormat = String.fromCharCode(0) + String.fromCharCode(splitBytes);
        for (let loop = splitBytes - 1; loop >= 0; loop--) {
            returnLvFormat = returnLvFormat + String.fromCharCode(lvArray[loop]);
        }
    }
    else {
        returnLvFormat = String.fromCharCode(messageLength);
    }
    return (returnLvFormat + messagString);
}
exports.lvFormat = lvFormat;
/**
 * This method does a thing.
 *
 * @export
 * @param {Buffer} messageString The message string
 * @returns {Buffer[]} A buffer array
 */
function lv2vFormat(messageString) {
    let returnString = [];
    let messageLength = messageString.length;
    if (messageLength === 0) {
        return [];
    }
    ;
    let bytePointer = 0;
    let extractChar = 0;
    let numberOfBufferedLine = 0;
    let byteCalcNumber;
    while (bytePointer < messageLength) {
        extractChar = messageString.readUInt8(bytePointer);
        numberOfBufferedLine = 1;
        if (extractChar === 0) {
            numberOfBufferedLine = messageString.readUInt8(bytePointer + 1);
            bytePointer = bytePointer + 2;
            if (numberOfBufferedLine === 0) {
                continue;
            }
            byteCalcNumber = 1;
            for (let loopFor = numberOfBufferedLine - 1; loopFor >= 0; loopFor--) {
                extractChar = (messageString.readUInt8(bytePointer + loopFor) * byteCalcNumber) + extractChar;
                byteCalcNumber = byteCalcNumber * 256;
            }
        }
        if (bytePointer > messageLength) {
            continue;
        }
        returnString.push(messageString.slice(bytePointer + numberOfBufferedLine, bytePointer + extractChar));
        bytePointer = bytePointer + extractChar;
    }
    return returnString;
}
exports.lv2vFormat = lv2vFormat;
function parseResponse(serviceClass, outputData, encoding) {
    // unpacking multiple times to get the token, remove the endiness by extracting from position 2
    let returnString = '';
    let returnArray;
    returnArray = lv2vFormat(outputData);
    returnArray = lv2vFormat(returnArray[1]);
    returnArray = lv2vFormat(returnArray[1]);
    returnString = returnArray[0].toString(encoding);
    if (returnString === 'ER') {
        throw returnArray.map(x => x.toString(encoding)).join('');
    }
    if (serviceClass === 5) {
        returnString = returnArray[2].toString(encoding) + String.fromCharCode(0) + returnArray[3].toString(encoding);
    }
    return returnString;
}
exports.parseResponse = parseResponse;
function sendingMessage(tokenMessage, mrpcMessage) {
    return v2lvFormat([tokenMessage, mrpcMessage]);
}
exports.sendingMessage = sendingMessage;
function mrpcMessage(mrpcId, version, prepareString) {
    let exchangeMessage = [
        mrpcId,
        version,
        prepareString,
        mrpcConnMessage()
    ];
    return v2lvFormat(exchangeMessage);
}
exports.mrpcMessage = mrpcMessage;
function tokenMessage(serviceClass, token, messageId) {
    let exchangeMessage = [
        serviceClass.toString(),
        token,
        messageId.toString(),
        '0',
        ''
    ];
    return v2lvFormat(exchangeMessage);
}
exports.tokenMessage = tokenMessage;
function connectionObject(envUser, envPassword) {
    let perpareString = [
        '1',
        envUser.toString(),
        'nowhere',
        envPassword,
        '',
        '',
        netConnMessage()
    ];
    return v2lvFormat(perpareString);
}
exports.connectionObject = connectionObject;
function checkObject(localFile, token) {
    let messageArray = [
        'CHECKOBJ',
        '', '', localFile, '', '', token
    ];
    return v2lvFormat(messageArray);
}
exports.checkObject = checkObject;
function saveObject(localFile, token, username) {
    let messageArray = [
        'SAVEOBJ',
        '', '', localFile, '', '', token, username
    ];
    return v2lvFormat(messageArray);
}
exports.saveObject = saveObject;
function initCodeObject(code, compilationToken) {
    let messageArray = [
        'INITCODE',
        code, compilationToken
    ];
    return v2lvFormat(messageArray);
}
exports.initCodeObject = initCodeObject;
function testCompileObject(fileName, compilationToken) {
    let messageArray = [
        'EXECCOMP',
        '', compilationToken, fileName
    ];
    return v2lvFormat(messageArray);
}
exports.testCompileObject = testCompileObject;
function initObject(objectId, objectName) {
    let messageArray = [
        'INITOBJ',
        '', '', '', objectId, objectName
    ];
    return v2lvFormat(messageArray);
}
exports.initObject = initObject;
function retObject(token) {
    let messageArray = [
        'RETOBJ',
        '', '', '', '', '', token
    ];
    return v2lvFormat(messageArray);
}
exports.retObject = retObject;
function preCompileObject(fileName) {
    let messageArray = [
        'PRECMP',
        '', '', fileName
    ];
    return v2lvFormat(messageArray);
}
exports.preCompileObject = preCompileObject;
function compileObject(compilationToken) {
    let messageArray = [
        'CMPLINK',
        '', compilationToken
    ];
    return v2lvFormat(messageArray);
}
exports.compileObject = compileObject;
function pslRunObject(compilationToken) {
    let messageArray = [
        'PSLRUN',
        '', compilationToken
    ];
    return v2lvFormat(messageArray);
}
exports.pslRunObject = pslRunObject;
function customRunObject(request, compilationToken) {
    let messageArray = [
        request,
        '', compilationToken
    ];
    return v2lvFormat(messageArray);
}
exports.customRunObject = customRunObject;
function sqlObject(query, rows) {
    let messageArray = [
        query,
        rows,
        ''
    ];
    return v2lvFormat(messageArray);
}
exports.sqlObject = sqlObject;
function batchCompileObject(dbtblTableName, elementName) {
    let messageArray = [
        dbtblTableName,
        elementName
    ];
    return v2lvFormat(messageArray);
}
exports.batchCompileObject = batchCompileObject;
function getPslCls() {
    let messageArray = [
        'GETPSLCLS', '', ''
    ];
    return v2lvFormat(messageArray);
}
exports.getPslCls = getPslCls;
function getUserName() {
    return os.userInfo().username;
}
exports.getUserName = getUserName;
function getObjectType(fileName) {
    let elementBaseName = path.basename(fileName);
    let elementName = elementBaseName.substr(0, elementBaseName.lastIndexOf('.'));
    let elementExtension = elementBaseName.substr(elementBaseName.lastIndexOf('.') + 1, elementBaseName.length);
    if (elementName.includes('.'))
        elementName = elementBaseName;
    return {
        fileId: getFileDetails(elementExtension),
        fileName: elementName,
        fileBaseName: elementBaseName
    };
}
exports.getObjectType = getObjectType;
function getFileDetails(fileExtension) {
    if (fileExtension in exports.extensionToDescription)
        return exports.extensionToDescription[fileExtension];
    throw new Error(`Invalid file extension: ${fileExtension}`);
}
function getDbtblInfo(fileId) {
    switch (fileId) {
        case 'Batch': return 'DBTBL33';
        case 'Column': return 'DBTBL25';
        case 'Procedure': return 'DBTBL1';
        case 'Table': return 'DBTBL1';
        default: return 'Unknown Type';
    }
}
exports.getDbtblInfo = getDbtblInfo;
function pack(totalLength) {
    // For ING we use Big Endian !h which is 2 bytes
    let quotient = Math.floor(totalLength / 256);
    let firstByte = String.fromCharCode(quotient);
    let secondByte = String.fromCharCode(totalLength - (quotient * 256));
    return (firstByte + secondByte);
}
exports.pack = pack;
function unpack(message) {
    // For ING we use Big Endian !h which is 2 bytes
    if (!message.readUInt8(0) && !message.readUInt8(1))
        return longMessageLength(message);
    return { totalBytes: (message.readUInt8(0) * 256) + message.readUInt8(1), startByte: 3 };
}
exports.unpack = unpack;
function longMessageLength(message) {
    // the third byte of the message tells us how many bytes are used to encode the length
    const numberOfBytes = message.readUInt8(2);
    const lastLengthByte = 3 + numberOfBytes;
    // slice the message to just use the bytes that encode message length
    const messageLengthBytes = message.slice(3, lastLengthByte);
    let totalBytes = 0;
    for (let index = 0; index < messageLengthBytes.length; index++) {
        const byte = messageLengthBytes.readUInt8(index);
        totalBytes += byte * Math.pow(256, (messageLengthBytes.length - 1 - index));
    }
    return { totalBytes, startByte: lastLengthByte + 1 };
}
function mrpcConnMessage() {
    return String.fromCharCode(4, 3, 2) + '1';
}
function netConnMessage() {
    return String.fromCharCode(21, 2) + '5' + String.fromCharCode(6) + 'ICODE' + String.fromCharCode(2) + '1' + String.fromCharCode(8) + 'PREPARE' + String.fromCharCode(2) + '3';
}
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoidXRpbHMuanMiLCJzb3VyY2VSb290IjoiIiwic291cmNlcyI6WyIuLi8uLi8uLi9zcmMvbXRtL3V0aWxzLnRzIl0sIm5hbWVzIjpbXSwibWFwcGluZ3MiOiI7OztBQUFBLHlCQUF5QjtBQUN6Qiw2QkFBNkI7QUFRaEIsUUFBQSxzQkFBc0IsR0FBOEI7SUFDaEUsT0FBTyxFQUFFLE9BQU87SUFDaEIsS0FBSyxFQUFFLFFBQVE7SUFDZixLQUFLLEVBQUUsTUFBTTtJQUNiLEtBQUssRUFBRSxhQUFhO0lBQ3BCLGlCQUFpQjtJQUNqQixLQUFLLEVBQUUsT0FBTztJQUNkLEtBQUssRUFBRSxTQUFTO0lBQ2hCLEdBQUcsRUFBRSxXQUFXO0lBQ2hCLEtBQUssRUFBRSxjQUFjO0lBQ3JCLE1BQU0sRUFBRSxXQUFXO0lBQ25CLFlBQVksRUFBRSxpQkFBaUI7SUFDL0IsS0FBSyxFQUFFLFVBQVU7SUFDakIsS0FBSyxFQUFFLFVBQVU7SUFDakIsTUFBTSxFQUFFLFdBQVc7SUFDbkIsU0FBUyxFQUFFLGNBQWM7SUFDekIsTUFBTSxFQUFFLFlBQVk7SUFDcEIsUUFBUSxFQUFFLGlCQUFpQjtJQUMzQixLQUFLLEVBQUUsT0FBTztJQUNkLEtBQUssRUFBRSxRQUFRO0lBQ2YsS0FBSyxFQUFFLFFBQVE7SUFDZixzQkFBc0I7SUFDdEIsNkJBQTZCO0lBQzdCLEtBQUssRUFBRSxPQUFPO0lBQ2QsTUFBTSxFQUFFLFNBQVM7Q0FDakIsQ0FBQTtBQUVELFNBQWdCLFVBQVUsQ0FBQyxZQUFzQjtJQUNoRCxJQUFJLFNBQVMsR0FBRyxFQUFFLENBQUE7SUFDbEIsSUFBSSxZQUFZLENBQUMsTUFBTSxLQUFLLENBQUMsRUFBRTtRQUM5QixZQUFZLENBQUMsT0FBTyxDQUFDLGFBQWEsQ0FBQyxFQUFFO1lBQ3BDLFNBQVMsR0FBRyxTQUFTLEdBQUcsUUFBUSxDQUFDLGFBQWEsQ0FBQyxDQUFBO1FBQ2hELENBQUMsQ0FBQyxDQUFBO0tBQ0Y7SUFDRCxPQUFPLFNBQVMsQ0FBQTtBQUNqQixDQUFDO0FB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LE9BQU8sQ0FBQyxTQUFTLENBQUMsQ0FBQyxDQUFDLEdBQUcsR0FBRyxDQUFDLEdBQUcsT0FBTyxDQUFDLFNBQVMsQ0FBQyxDQUFDLENBQUMsRUFBRSxTQUFTLEVBQUUsQ0FBQyxFQUFFLENBQUM7QUFDMUYsQ0FBQztBQUpELHdCQUlDO0FBRUQsU0FBUyxpQkFBaUIsQ0FBQyxPQUFlO0lBQ3pDLHNGQUFzRjtJQUN0RixNQUFNLGFBQWEsR0FBRyxPQUFPLENBQUMsU0FBUyxDQUFDLENBQUMsQ0FBQyxDQUFDO0lBQzNDLE1BQU0sY0FBYyxHQUFHLENBQUMsR0FBRyxhQUFhLENBQUM7SUFFekMscUVBQXFFO0lBQ3JFLE1BQU0sa0JBQWtCLEdBQUcsT0FBTyxDQUFDLEtBQUssQ0FBQyxDQUFDLEVBQUUsY0FBYyxDQUFDLENBQUM7SUFDNUQsSUFBSSxVQUFVLEdBQUcsQ0FBQyxDQUFDO0lBQ25CLEtBQUssSUFBSSxLQUFLLEdBQUcsQ0FBQyxFQUFFLEtBQUssR0FBRyxrQkFBa0IsQ0FBQyxNQUFNLEVBQUUsS0FBSyxFQUFFLEVBQUU7UUFDL0QsTUFBTSxJQUFJLEdBQUcsa0JBQWtCLENBQUMsU0FBUyxDQUFDLEtBQUssQ0FBQyxDQUFDO1FBQ2pELFVBQVUsSUFBSSxJQUFJLEdBQUcsU0FBQSxHQUFHLEVBQUksQ0FBQyxrQkFBa0IsQ0FBQyxNQUFNLEdBQUcsQ0FBQyxHQUFHLEtBQUssQ0FBQyxDQUFBLENBQUM7S0FDcEU7SUFDRCxPQUFPLEVBQUUsVUFBVSxFQUFFLFNBQVMsRUFBRSxjQUFjLEdBQUcsQ0FBQyxFQUFDLENBQUM7QUFDckQsQ0FBQztBQUVELFNBQVMsZUFBZTtJQUN2QixPQUFPLE1BQU0sQ0FBQyxZQUFZLENBQUMsQ0FBQyxFQUFFLENBQUMsRUFBRSxDQUFDLENBQUMsR0FBRyxHQUFHLENBQUM7QUFDM0MsQ0FBQztBQUVELFNBQVMsY0FBYztJQUN0QixPQUFPLE1BQU0sQ0FBQyxZQUFZLENBQUMsRUFBRSxFQUFFLENBQUMsQ0FBQyxHQUFHLEdBQUcsR0FBRyxNQUFNLENBQUMsWUFBWSxDQUFDLENBQUMsQ0FBQyxHQUFHLE9BQU8sR0FBRyxNQUFNLENBQUMsWUFBWSxDQUFDLENBQUMsQ0FBQyxHQUFHLEdBQUcsR0FBRyxNQUFNLENBQUMsWUFBWSxDQUFDLENBQUMsQ0FBQyxHQUFHLFNBQVMsR0FBRyxNQUFNLENBQUMsWUFBWSxDQUFDLENBQUMsQ0FBQyxHQUFHLEdBQUcsQ0FBQztBQUMvSyxDQUFDIn0=