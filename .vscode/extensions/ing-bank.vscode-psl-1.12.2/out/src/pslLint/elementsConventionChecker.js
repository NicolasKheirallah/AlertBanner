"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.MemberStartsWithV = exports.MemberLength = exports.MemberCamelCase = exports.MemberLiteralCase = exports.PropertyIsDuplicate = exports.PropertyIsDummy = exports.PropertyStartsWithZ = exports.MethodStartsWithZ = void 0;
const parser_1 = require("../parser/parser");
const api_1 = require("./api");
class MethodStartsWithZ extends api_1.MethodRule {
    report(method) {
        const diagnostics = [];
        startsWithZ(method, diagnostics, this.ruleName);
        return diagnostics;
    }
}
exports.MethodStartsWithZ = MethodStartsWithZ;
class PropertyStartsWithZ extends api_1.PropertyRule {
    report(property) {
        const diagnostics = [];
        startsWithZ(property, diagnostics, this.ruleName);
        return diagnostics;
    }
}
exports.PropertyStartsWithZ = PropertyStartsWithZ;
class PropertyIsDummy extends api_1.PropertyRule {
    report(property) {
        const diagnostics = [];
        if (!this.parsedDocument.extending) {
            this.isCalledDummy(property, diagnostics);
        }
        return diagnostics;
    }
    isCalledDummy(member, diagnostics) {
        if (member.id.value.toLowerCase() === 'dummy') {
            diagnostics.push(createDiagnostic(member, 'Usage of "dummy" property is discouraged', api_1.DiagnosticSeverity.Information, this.ruleName));
        }
    }
}
exports.PropertyIsDummy = PropertyIsDummy;
class PropertyIsDuplicate extends api_1.PropertyRule {
    report(property) {
        const diagnostics = [];
        this.isDuplicateProperty(property, diagnostics);
        return diagnostics;
    }
    isDuplicateProperty(property, diagnostics) {
        const slicedProperty = this.parsedDocument.properties.slice(0, this.parsedDocument.properties.findIndex(x => x.id.position.line === property.id.position.line));
        for (const checkProperty of slicedProperty) {
            if (checkProperty.id.value === property.id.value) {
                const diagnostic = new api_1.Diagnostic(property.id.getRange(), `Property "${property.id.value}" is already declared.`, this.ruleName, api_1.DiagnosticSeverity.Warning);
                const aboveDuplicateProperty = new api_1.DiagnosticRelatedInformation(checkProperty.id.getRange(), `Reference to property "${checkProperty.id.value}".`);
                diagnostic.relatedInformation = [
                    aboveDuplicateProperty,
                ];
                diagnostic.source = 'lint';
                diagnostics.push(diagnostic);
                break;
            }
            if (checkProperty.id.value.toLowerCase() === property.id.value.toLowerCase()) {
                const diagnostic = new api_1.Diagnostic(property.id.getRange(), `Property "${property.id.value}" is already declared with different case.`, this.ruleName, api_1.DiagnosticSeverity.Warning);
                const aboveDuplicateProperty = new api_1.DiagnosticRelatedInformation(checkProperty.id.getRange(), `Reference to property "${checkProperty.id.value}".`);
                diagnostic.relatedInformation = [
                    aboveDuplicateProperty,
                ];
                diagnostic.source = 'lint';
                diagnostics.push(diagnostic);
                break;
            }
        }
    }
}
exports.PropertyIsDuplicate = PropertyIsDuplicate;
class MemberLiteralCase extends api_1.MemberRule {
    report(member) {
        const diagnostics = [];
        this.checkUpperCase(member, diagnostics);
        return diagnostics;
    }
    checkUpperCase(member, diagnostics) {
        if ((member.modifiers.findIndex(x => x.value === 'literal') > -1)) {
            if (member.id.value !== member.id.value.toUpperCase()) {
                diagnostics.push(createDiagnostic(member, 'is literal but not upper case.', api_1.DiagnosticSeverity.Warning, this.ruleName));
            }
        }
    }
}
exports.MemberLiteralCase = MemberLiteralCase;
class MemberCamelCase extends api_1.MemberRule {
    report(member) {
        const diagnostics = [];
        this.memberCase(member, diagnostics);
        return diagnostics;
    }
    memberCase(member, diagnostics) {
        const isLiteral = (member.modifiers.findIndex(x => x.value === 'literal') > -1);
        let isStaticDeclaration = false;
        member.types.forEach(type => {
            if (type.value === member.id.value) {
                isStaticDeclaration = true;
            }
        });
        // exception for variables starting with percentage
        if (member.id.value.charAt(0) === '%')
            return;
        // exception for literal properties
        if (isLiteral || isStaticDeclaration)
            return;
        if (member.memberClass === parser_1.MemberClass.method) {
            const method = member;
            if (method.batch)
                return;
        }
        if (member.id.value.charAt(0) > 'z' || member.id.value.charAt(0) < 'a') {
            if (isPublicDeclaration(member)) {
                const diagnostic = new api_1.Diagnostic(member.id.getRange(), `Declaration "${member.id.value}" is public and does not start with lower case.`, this.ruleName, api_1.DiagnosticSeverity.Information);
                diagnostic.source = 'lint';
                diagnostic.member = member;
                diagnostics.push(diagnostic);
            }
            else {
                diagnostics.push(createDiagnostic(member, 'does not start with lowercase.', api_1.DiagnosticSeverity.Warning, this.ruleName));
            }
        }
    }
}
exports.MemberCamelCase = MemberCamelCase;
class MemberLength extends api_1.MemberRule {
    report(member) {
        const diagnostics = [];
        this.checkMemberLength(member, diagnostics);
        return diagnostics;
    }
    checkMemberLength(member, diagnostics) {
        if (member.id.value.length > 25) {
            diagnostics.push(createDiagnostic(member, 'is longer than 25 characters.', api_1.DiagnosticSeverity.Warning, this.ruleName));
        }
    }
}
exports.MemberLength = MemberLength;
class MemberStartsWithV extends api_1.MemberRule {
    report(member) {
        const diagnostics = [];
        this.checkStartsWithV(member, diagnostics);
        return diagnostics;
    }
    checkStartsWithV(member, diagnostics) {
        if (member.id.value.charAt(0) !== 'v')
            return;
        if (isPublicDeclaration(member)) {
            diagnostics.push(createDiagnostic(member, `is public and starts with 'v'.`, api_1.DiagnosticSeverity.Information, this.ruleName));
        }
        else {
            diagnostics.push(createDiagnostic(member, `starts with 'v'.`, api_1.DiagnosticSeverity.Warning, this.ruleName));
        }
    }
}
exports.MemberStartsWithV = MemberStartsWithV;
function createDiagnostic(member, message, diagnosticSeverity, ruleName) {
    const diagnostic = new api_1.Diagnostic(member.id.getRange(), `${printEnum(member.memberClass)} "${member.id.value}" ${message}`, ruleName, diagnosticSeverity);
    diagnostic.source = 'lint';
    diagnostic.member = member;
    return diagnostic;
}
function startsWithZ(member, diagnostics, ruleName) {
    const firstChar = member.id.value.charAt(0);
    if (firstChar === 'z' || firstChar === 'Z') {
        diagnostics.push(createDiagnostic(member, `starts with '${firstChar}'.`, api_1.DiagnosticSeverity.Information, ruleName));
    }
}
function printEnum(memberClass) {
    const enumName = parser_1.MemberClass[memberClass];
    const capitalizedEnumName = enumName.charAt(0).toUpperCase() + enumName.slice(1);
    return enumName === 'method' ? 'Label' : capitalizedEnumName;
}
function isPublicDeclaration(member) {
    const isPublic = member.modifiers.findIndex(x => x.value === 'public') > -1;
    return member.memberClass === parser_1.MemberClass.declaration && isPublic;
}
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoiZWxlbWVudHNDb252ZW50aW9uQ2hlY2tlci5qcyIsInNvdXJjZVJvb3QiOiIiLCJzb3VyY2VzIjpbIi4uLy4uLy4uL3NyYy9wc2xMaW50L2VsZW1lbnRzQ29udmVudGlvbkNoZWNrZXIudHMiXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6Ijs7O0FBQUEsNkNBQXlFO0FBQ3pFLCtCQUdlO0FBRWYsTUFBYSxpQkFBa0IsU0FBUSxnQkFBVTtJQUVoRCxNQUFNLENBQUMsTUFBYztRQUNwQixNQUFNLFdBQVcsR0FBaUIsRUFBRSxDQUFDO1FBRXJDLFdBQVcsQ0FBQyxNQUFNLEVBQUUsV0FBVyxFQUFFLElBQUksQ0FBQyxRQUFRLENBQUMsQ0FBQztRQUVoRCxPQUFPLFdBQVcsQ0FBQztJQUNwQixDQUFDO0NBQ0Q7QUFURCw4Q0FTQztBQUNELE1BQWEsbUJBQW9CLFNBQVEsa0JBQVk7SUFFcEQsTUFBTSxDQUFDLFFBQWtCO1FBQ3hCLE1BQU0sV0FBVyxHQUFpQixFQUFFLENBQUM7UUFFckMsV0FBVyxDQUFDLFFBQVEsRUFBRSxXQUFXLEVBQUUsSUFBSSxDQUFDLFFBQVEsQ0FBQyxDQUFDO1FBRWxELE9BQU8sV0FBVyxDQUFDO0lBQ3BCLENBQUM7Q0FDRDtBQVRELGtEQVNDO0FBRUQsTUFBYSxlQUFnQixTQUFRLGtCQUFZO0lBRWhELE1BQU0sQ0FBQyxRQUFrQjtRQUN4QixNQUFNLFdBQVcsR0FBaUIsRUFBRSxDQUFDO1FBQ3JDLElBQUksQ0FBQyxJQUFJLENBQUMsY0FBYyxDQUFDLFNBQVMsRUFBRTtZQUNuQyxJQUFJLENBQUMsYUFBYSxDQUFDLFFBQVEsRUFBRSxXQUFXLENBQUMsQ0FBQztTQUMxQztRQUNELE9BQU8sV0FBVyxDQUFDO0lBQ3BCLENBQUM7SUFFRCxhQUFhLENBQUMsTUFBYyxFQUFFLFdBQXlCO1FBQ3RELElBQUksTUFBTSxDQUFDLEVBQUUsQ0FBQyxLQUFLLENBQUMsV0FBVyxFQUFFLEtBQUssT0FBTyxFQUFFO1lBQzlDLFdBQVcsQ0FBQyxJQUFJLENBQ2Ys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