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
exports.matchConfig = exports.getConfig = exports.removeConfig = exports.transform = exports.setConfig = exports.activeConfigs = void 0;
const fs = require("fs-extra");
const minimatch = require("minimatch");
const path = require("path");
exports.activeConfigs = new Map();
function setConfig(configPath) {
    return __awaiter(this, void 0, void 0, function* () {
        const configBaseDir = path.dirname(configPath);
        const config = yield fs.readFile(configPath).then(b => JSON.parse(b.toString()));
        exports.activeConfigs.set(configBaseDir, transform(config));
    });
}
exports.setConfig = setConfig;
function transform(config) {
    const includes = [];
    const excludes = [];
    for (const pattern in config.include) {
        if (config.include.hasOwnProperty(pattern)) {
            const rules = config.include[pattern];
            const regexpPattern = minimatch.makeRe(pattern);
            if (!regexpPattern)
                throw new Error(`Invalid regexp patter ${pattern}`);
            includes.push({ pattern: regexpPattern, rules });
        }
    }
    for (const pattern in config.exclude) {
        if (config.exclude.hasOwnProperty(pattern)) {
            const rules = config.exclude[pattern];
            const regexpPattern = minimatch.makeRe(pattern);
            if (!regexpPattern)
                throw new Error(`Invalid regexp patter ${pattern}`);
            excludes.push({ pattern: regexpPattern, rules });
        }
    }
    return { include: includes, exclude: excludes };
}
exports.transform = transform;
function removeConfig(configPath) {
    return __awaiter(this, void 0, void 0, function* () {
        const configBaseDir = path.dirname(configPath);
        exports.activeConfigs.delete(configBaseDir);
    });
}
exports.removeConfig = removeConfig;
function getConfig(fsPath) {
    for (const configBaseDir of exports.activeConfigs.keys()) {
        const relative = path.relative(configBaseDir, fsPath);
        if (!!relative && !relative.startsWith('..') && !path.isAbsolute(relative)) {
            return exports.activeConfigs.get(configBaseDir);
        }
    }
}
exports.getConfig = getConfig;
function matchConfig(fileName, ruleName, configObj) {
    let matches = false;
    const findMatch = (configSettings) => {
        for (const configSetting of configSettings) {
            if (!fileName.match(configSetting.pattern))
                continue;
            for (const rulePattern of configSetting.rules) {
                if (rulePattern === '*' || rulePattern === ruleName)
                    return true;
            }
        }
    };
    matches = findMatch(configObj.include) || false;
    if (!matches)
        return false;
    return !findMatch(configObj.exclude);
}
exports.matchConfig = matchConfig;
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoiY29uZmlnLmpzIiwic291cmNlUm9vdCI6IiIsInNvdXJjZXMiOlsiLi4vLi4vLi4vc3JjL3BzbExpbnQvY29uZmlnLnRzIl0sIm5hbWVzIjpbXSwibWFwcGluZ3MiOiI7Ozs7Ozs7Ozs7OztBQUFBLCtCQUErQjtBQUMvQix1Q0FBdUM7QUFDdkMsNkJBQTZCO0FBR2xCLFFBQUEsYUFBYSxHQUFvQyxJQUFJLEdBQUcsRUFBRSxDQUFDO0FBaUJ0RSxTQUFzQixTQUFTLENBQUMsVUFBa0I7O1FBQ2pELE1BQU0sYUFBYSxHQUFrQixJQUFJLENBQUMsT0FBTyxDQUFDLFVBQVUsQ0FBQyxDQUFDO1FBQzlELE1BQU0sTUFBTSxHQUFXLE1BQU0sRUFBRSxDQUFDLFFBQVEsQ0FBQyxVQUFVLENBQUMsQ0FBQyxJQUFJLENBQUMsQ0FBQyxDQUFDLEVBQUUsQ0FBQyxJQUFJLENBQUMsS0FBSyxDQUFDLENBQUMsQ0FBQyxRQUFRLEVBQUUsQ0FBQyxDQUFDLENBQUM7UUFDekYscUJBQWEsQ0FBQyxHQUFHLENBQUMsYUFBYSxFQUFFLFNBQVMsQ0FBQyxNQUFNLENBQUMsQ0FBQyxDQUFDO0lBQ3JELENBQUM7Q0FBQTtBQUpELDhCQUlDO0FBRUQsU0FBZ0IsU0FBUyxDQUFDLE1BQWM7SUFDdkMsTUFBTSxRQUFRLEdBQXFCLEVBQUUsQ0FBQztJQUN0QyxNQUFNLFFBQVEsR0FBcUIsRUFBRSxDQUFDO0lBQ3RDLEtBQUssTUFBTSxPQUFPLElBQUksTUFBTSxDQUFDLE9BQU8sRUFBRTtRQUNyQyxJQUFJLE1BQU0sQ0FBQyxPQUFPLENBQUMsY0FBYyxDQUFDLE9BQU8sQ0FBQyxFQUFFO1lBQzNDLE1BQU0sS0FBSyxHQUFHLE1BQU0sQ0FBQyxPQUFPLENBQUMsT0FBTyxDQUFDLENBQUM7WUFDdEMsTUFBTSxhQUFhLEdBQUcsU0FBUyxDQUFDLE1BQU0sQ0FBQyxPQUFPLENBQUMsQ0FBQTtZQUUvQyxJQUFJLENBQUMsYUFBYTtnQkFBRSxNQUFNLElBQUksS0FBSyxDQUFDLHlCQUF5QixPQUFPLEVB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