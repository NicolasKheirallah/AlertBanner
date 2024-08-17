"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.getDiagnostics = void 0;
const path = require("path");
const api_1 = require("./api");
const config_1 = require("./config");
const elementsConventionChecker_1 = require("./elementsConventionChecker");
const methodDoc_1 = require("./methodDoc");
const multiLineDeclare_1 = require("./multiLineDeclare");
const parameters_1 = require("./parameters");
const runtime_1 = require("./runtime");
const tblcolDoc_1 = require("./tblcolDoc");
const todos_1 = require("./todos");
/**
 * Add new rules here to have them checked at the appropriate time.
 */
const componentRules = [];
const fileDefinitionRules = [
    new tblcolDoc_1.TblColDocumentation(),
];
const pslRules = [
    new todos_1.TodoInfo(),
];
const memberRules = [
    new elementsConventionChecker_1.MemberCamelCase(),
    new elementsConventionChecker_1.MemberLength(),
    new elementsConventionChecker_1.MemberStartsWithV(),
    new elementsConventionChecker_1.MemberLiteralCase(),
];
const methodRules = [
    new methodDoc_1.MethodDocumentation(),
    new methodDoc_1.MethodSeparator(),
    new parameters_1.MethodParametersOnNewLine(),
    new runtime_1.RuntimeStart(),
    new multiLineDeclare_1.MultiLineDeclare(),
    new methodDoc_1.TwoEmptyLines(),
];
const propertyRules = [
    new elementsConventionChecker_1.PropertyIsDummy(),
    new elementsConventionChecker_1.PropertyIsDuplicate(),
];
const declarationRules = [];
const parameterRules = [];
function getDiagnostics(profileComponent, parsedDocument, useConfig) {
    const subscription = new RuleSubscription(profileComponent, parsedDocument, useConfig);
    return subscription.reportRules();
}
exports.getDiagnostics = getDiagnostics;
/**
 * Interface for adding and executing rules.
 */
class RuleSubscription {
    constructor(profileComponent, parsedDocument, useConfig) {
        this.profileComponent = profileComponent;
        this.parsedDocument = parsedDocument;
        this.diagnostics = [];
        const config = useConfig ? (0, config_1.getConfig)(this.profileComponent.fsPath) : undefined;
        const initializeRules = (rules) => {
            return rules.filter(rule => {
                if (!config)
                    return true;
                return (0, config_1.matchConfig)(path.basename(this.profileComponent.fsPath), rule.ruleName, config);
            }).map(rule => {
                rule.profileComponent = this.profileComponent;
                return rule;
            });
        };
        const initializePslRules = (rules) => {
            const componentInitialized = initializeRules(rules);
            const pslParsedDocument = this.parsedDocument;
            return componentInitialized.map(rule => {
                rule.parsedDocument = pslParsedDocument;
                return rule;
            });
        };
        this.componentRules = initializeRules(componentRules);
        this.fileDefinitionRules = initializeRules(fileDefinitionRules);
        this.pslRules = initializePslRules(pslRules);
        this.methodRules = initializePslRules(methodRules);
        this.memberRules = initializePslRules(memberRules);
        this.propertyRules = initializePslRules(propertyRules);
        this.declarationRules = initializePslRules(declarationRules);
        this.parameterRules = initializePslRules(parameterRules);
    }
    reportRules() {
        const addDiagnostics = (rules, ...args) => {
            rules.forEach(rule => this.diagnostics.push(...rule.report(...args)));
        };
        addDiagnostics(this.componentRules);
        if (api_1.ProfileComponent.isFileDefinition(this.profileComponent.fsPath)) {
            addDiagnostics(this.fileDefinitionRules);
        }
        if (api_1.ProfileComponent.isPsl(this.profileComponent.fsPath)) {
            addDiagnostics(this.pslRules);
            const parsedDocument = this.parsedDocument;
            for (const property of parsedDocument.properties) {
                addDiagnostics(this.memberRules, property);
                addDiagnostics(this.propertyRules, property);
            }
            for (const declaration of parsedDocument.declarations) {
                addDiagnostics(this.memberRules, declaration);
                addDiagnostics(this.declarationRules, declaration);
            }
            for (const method of parsedDocument.methods) {
                addDiagnostics(this.memberRules, method);
                addDiagnostics(this.methodRules, method);
                for (const parameter of method.parameters) {
                    addDiagnostics(this.memberRules, parameter);
                    addDiagnostics(this.parameterRules, parameter, method);
                }
                for (const declaration of method.declarations) {
                    addDiagnostics(this.memberRules, declaration);
                    addDiagnostics(this.declarationRules, declaration, method);
                }
            }
        }
        return this.diagnostics;
    }
}
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoiYWN0aXZhdGUuanMiLCJzb3VyY2VSb290IjoiIiwic291cmNlcyI6WyIuLi8uLi8uLi9zcmMvcHNsTGludC9hY3RpdmF0ZS50cyJdLCJuYW1lcyI6W10sIm1hcHBpbmdzIjoiOzs7QUFBQSw2QkFBNkI7QUFDN0IsK0JBR2U7QUFDZixxQ0FBa0Q7QUFNbEQsMkVBR3FDO0FBQ3JDLDJDQUFrRjtBQUNsRix5REFBc0Q7QUFDdEQsNkNBQXlEO0FBQ3pELHVDQUF5QztBQUN6QywyQ0FBa0Q7QUFDbEQsbUNBQW1DO0FBRW5DOztHQUVHO0FBQ0gsTUFBTSxjQUFjLEdBQTJCLEVBQUUsQ0FBQztBQUNsRCxNQUFNLG1CQUFtQixHQUF5QjtJQUNqRCxJQUFJLCtCQUFtQixFQUFFO0NBQ3pCLENBQUM7QUFDRixNQUFNLFFBQVEsR0FBYztJQUMzQixJQUFJLGdCQUFRLEVBQUU7Q0FDZCxDQUFDO0FBQ0YsTUFBTSxXQUFXLEdBQWlCO0lBQ2pDLElBQUksMkNBQWUsRUFBRTtJQUNyQixJQUFJLHdDQUFZLEVBQUU7SUFDbEIsSUFBSSw2Q0FBaUIsRUFBRTtJQUN2QixJQUFJLDZDQUFpQixFQUFFO0NBQ3ZCLENBQUM7QUFDRixNQUFNLFdBQVcsR0FBaUI7SUFDakMsSUFBSSwrQkFBbUIsRUFBRTtJQUN6QixJQUFJLDJCQUFlLEVBQUU7SUFDckIsSUFBSSxzQ0FBeUIsRUFBRTtJQUMvQixJQUFJLHNCQUFZLEVBQUU7SUFDbEIsSUFBSSxtQ0FBZ0IsRUFBRTtJQUN0QixJQUFJLHlCQUFhLEVBQUU7Q0FDbkIsQ0FBQztBQUNGLE1BQU0sYUFBYSxHQUFtQjtJQUNyQyxJQUFJLDJDQUFlLEVBQUU7SUFDckIsSUFBSSwrQ0FBbUIsRUFBRTtDQUN6QixDQUFDO0FBQ0YsTUFBTSxnQkFBZ0IsR0FBc0IsRUFBRSxDQUFDO0FBQy9DLE1BQU0sY0FBYyxHQUFvQixFQUFFLENBQUM7QUFFM0MsU0FBZ0IsY0FBYyxDQUM3QixnQkFBa0MsRUFDbEMsY0FBK0IsRUFDL0IsU0FBbUI7SUFFbkIsTUFBTSxZ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