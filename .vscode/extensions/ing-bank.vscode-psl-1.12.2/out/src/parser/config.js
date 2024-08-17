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
exports.getFinderPaths = exports.removeConfig = exports.setConfig = void 0;
const fs = require("fs-extra");
const path = require("path");
const activeConfigs = new Map();
function setConfig(configPath, workspaces) {
    return __awaiter(this, void 0, void 0, function* () {
        const configBaseDir = path.dirname(configPath);
        const config = yield fs.readFile(configPath).then(b => JSON.parse(b.toString()));
        config.parentProjects = config.parentProjects.map(p => workspaces.get(p)).filter(x => x);
        activeConfigs.set(configBaseDir, config);
    });
}
exports.setConfig = setConfig;
function removeConfig(configPath) {
    return __awaiter(this, void 0, void 0, function* () {
        const configBaseDir = path.dirname(configPath);
        activeConfigs.delete(configBaseDir);
    });
}
exports.removeConfig = removeConfig;
function getFinderPaths(currentDir, activeRoutine) {
    const defaultPslSources = ['dataqwik/procedure/', 'psl/'];
    const defaultFileDefinitionSources = ['dataqwik/table/'];
    const config = activeConfigs.get(currentDir);
    const projectPsl = [];
    const tables = [];
    const loadPsl = (base, source) => projectPsl.push(path.join(base, source));
    const loadFileDefinition = (base, source) => tables.push(path.join(base, source));
    const relativePslSources = config && config.pslSources ? config.pslSources : defaultPslSources;
    const relativeFileDefinitionSource = config && config.fileDefinitionSources ?
        config.fileDefinitionSources : defaultFileDefinitionSources;
    const corePsl = path.join(currentDir, '.vscode/pslcls/');
    // load core first
    projectPsl.push(corePsl);
    // load base sources
    relativePslSources.forEach(source => loadPsl(currentDir, source));
    relativeFileDefinitionSource.forEach(source => loadFileDefinition(currentDir, source));
    // load parent sources
    if (config && config.parentProjects) {
        for (const parent of config.parentProjects) {
            relativePslSources.forEach(source => loadPsl(parent, source));
            relativeFileDefinitionSource.forEach(source => loadFileDefinition(parent, source));
        }
    }
    return {
        activeRoutine,
        corePsl,
        projectPsl,
        tables,
    };
}
exports.getFinderPaths = getFinderPaths;
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoiY29uZmlnLmpzIiwic291cmNlUm9vdCI6IiIsInNvdXJjZXMiOlsiLi4vLi4vLi4vc3JjL3BhcnNlci9jb25maWcudHMiXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6Ijs7Ozs7Ozs7Ozs7O0FBQUEsK0JBQStCO0FBQy9CLDZCQUE2QjtBQUk3QixNQUFNLGFBQWEsR0FBc0MsSUFBSSxHQUFHLEVBQUUsQ0FBQztBQVFuRSxTQUFzQixTQUFTLENBQUMsVUFBa0IsRUFBRSxVQUErQjs7UUFDbEYsTUFBTSxhQUFhLEdBQWtCLElBQUksQ0FBQyxPQUFPLENBQUMsVUFBVSxDQUFDLENBQUM7UUFDOUQsTUFBTSxNQUFNLEdBQWtCLE1BQU0sRUFBRSxDQUFDLFFBQVEsQ0FBQyxVQUFVLENBQUMsQ0FBQyxJQUFJLENBQUMsQ0FBQyxDQUFDLEVBQUUsQ0FBQyxJQUFJLENBQUMsS0FBSyxDQUFDLENBQUMsQ0FBQyxRQUFRLEVBQUUsQ0FBQyxDQUFDLENBQUM7UUFDaEcsTUFBTSxDQUFDLGNBQWMsR0FBRyxNQUFNLENBQUMsY0FBYyxDQUFDLEdBQUcsQ0FBQyxDQUFDLENBQUMsRUFBRSxDQUFDLFVBQVUsQ0FBQyxHQUFHLENBQUMsQ0FBQyxDQUFDLENBQUMsQ0FBQyxNQUFNLENBQUMsQ0FBQyxDQUFDLEVBQUUsQ0FBQyxDQUFDLENBQUMsQ0FBQztRQUN6RixhQUFhLENBQUMsR0FBRyxDQUFDLGFBQWEsRUFBRSxNQUFNLENBQUMsQ0FBQztJQUMxQyxDQUFDO0NBQUE7QUFMRCw4QkFLQztBQUVELFNBQXNCLFlBQVksQ0FBQyxVQUFrQjs7UUFDcEQsTUFBTSxhQUFhLEdBQWtCLElBQUksQ0FBQyxPQUFPLENBQUMsVUFBVSxDQUFDLENBQUM7UUFDOUQsYUFBYSxDQUFDLE1BQU0sQ0FBQyxhQUFhLENBQUMsQ0FBQztJQUNyQyxDQUFDO0NBQUE7QUFIRCxvQ0FHQztBQWdDRCxTQUFnQixjQUFjLENBQUMsVUFBa0IsRUFBRSxhQUFzQjtJQUV4RSxNQUFNLGlCQUFpQixHQUFHLENBQUMscUJBQXFCLEVBQUUsTUFBTSxDQUFDLENBQUM7SUFDMUQsTUFBTSw0QkFBNEIsR0FBRyxDQUFDLGlCQUFpQixDQUFDLENBQUM7SUFFekQsTUFBTSxNQUFNLEdBQThCLGFBQWEsQ0FBQyxHQUFHLENBQUMsVUFBVSxDQUFDLENBQUM7SUFFeEUsTUFBTSxVQUFVLEdBQUcsRUFBRSxDQUFDO0lBQ3RCLE1BQU0sTUFBTSxHQUFHLEVBQUUsQ0FBQztJQUVsQixNQUFNLE9BQU8sR0FBRyxDQUFDLElBQUksRUFBRSxNQUFNLEVBQUUsRUFBRSxDQUFDLFVBQVUsQ0FBQyxJQUFJLENBQUMsSUFBSSxDQUFDLElBQUksQ0FBQyxJQUFJLEVBQUUsTUFBTSxDQUFDLENBQUMsQ0FBQztJQUMzRSxNQUFNLGtCQUFrQixHQUFHLENBQUMsSUFBSSxFQUFFLE1BQU0sRUFBRSxFQUFFLENBQUMsTUFBTSxDQUFDLElBQUksQ0FBQyxJQUFJLENBQUMsSUFBSSxDQUFDLElBQUksRUFBRSxNQUFNLENBQUMsQ0FBQyxDQUFDO0lBRWxGLE1BQU0sa0JBQWtCLEdBQUcsTUFBTSxJQUFJLE1BQU0sQ0FBQyxVQUFVLENBQUMsQ0FBQyxDQUFDLE1BQU0sQ0FBQyxVQUFVLENBQUMsQ0FBQyxDQUFDLGlCQUFpQixDQUFDO0lBQy9GLE1BQU0sNEJBQTRCLEdBQUcsTUFBTSxJQUFJLE1BQU0sQ0FBQyxxQkFBcUIsQ0FBQyxDQUFDO1FBQzVFLE1BQU0sQ0FBQyxxQkFBcUIsQ0FBQyxDQUFDLENBQUMsNEJBQTRCLENBQUM7SUFFN0QsTUFBTSxPQUFPLEdBQUcsSUFBSSxDQUFDLElBQUksQ0FBQyxVQUFVLEVBQUUsaUJBQWlCLENBQUMsQ0FBQztJQUN6RCxrQkFBa0I7SUFDbEIsVUFBVSxDQUFDLElBQUksQ0FBQyxPQUFPLENBQUMsQ0FBQztJQUV6QixvQkFBb0I7SUFDcEIsa0JBQWtCLENBQUMsT0FBTyxDQUFDLE1BQU0sQ0FBQyxFQUFFLENBQUMsT0FBTyxDQUFDLFVBQVUsRUFBRSxNQUFNLENBQUMsQ0FBQyxDQUFDO0lBQ2xFLDRCQUE0QixDQUFDLE9BQU8sQ0FBQyxNQUFNLENBQUMsRUFBRSxDQUFDLGtCQUFrQixDQUFDLFVBQVUsRUFBRSxNQUFNLENBQUMsQ0FBQyxDQUFDO0lBRXZGLHNCQUFzQjtJQUN0QixJQUFJLE1BQU0sSUFBSSxNQUFNLENBQUMsY0FBYyxFQUFFO1FBQ3BDLEtBQUssTUFBTSxNQUFNLElBQUksTUFBTSxDQUFDLGNBQWMsRUFBRTtZQUMzQyxrQkFBa0IsQ0FBQyxPQUFPLENBQUMsTUFBTSxDQUFDLEVBQUUsQ0FBQyxPQUFPLENBQUMsTUFBTSxFQUFFLE1BQU0sQ0FBQyxDQUFDLENBQUM7WUFDOUQsNEJBQTRCLENBQUMsT0FBTyxDQUFDLE1BQU0sQ0FBQyxFQUFFLENBQUMsa0JBQWtCLENBQUMsTUFBTSxFQUFFLE1BQU0sQ0FBQyxDQUFDLENBQUM7U0FDbkY7S0FFRDtJQUVELE9BQU87UUFDTixhQUFhO1FBQ2IsT0FBTztRQUNQLFVBQVU7UUFDVixNQUFNO0tBQ04sQ0FBQztBQUNILENBQUM7QUF4Q0Qsd0NBd0NDIn0=