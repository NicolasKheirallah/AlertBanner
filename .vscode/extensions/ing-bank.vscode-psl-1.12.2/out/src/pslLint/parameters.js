"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.MethodParametersOnNewLine = void 0;
const api_1 = require("./api");
/**
 * Checks if multiple parameters are written on the same line as the method declaration.
 */
class MethodParametersOnNewLine extends api_1.MethodRule {
    report(method) {
        if (method.batch)
            return [];
        const diagnostics = [];
        const methodLine = method.id.position.line;
        let previousParam;
        for (const param of method.parameters) {
            const paramPosition = param.id.position;
            if (previousParam && paramPosition.line === previousParam.id.position.line) {
                const message = `Parameter "${param.id.value}" on same line as parameter "${previousParam.id.value}".`;
                const diagnostic = new api_1.Diagnostic(param.id.getRange(), message, this.ruleName, api_1.DiagnosticSeverity.Warning);
                diagnostic.source = 'lint';
                diagnostics.push(diagnostic);
            }
            else if (method.parameters.length > 1 && paramPosition.line === methodLine) {
                const message = `Parameter "${param.id.value}" on same line as label "${method.id.value}".`;
                const diagnostic = new api_1.Diagnostic(param.id.getRange(), message, this.ruleName, api_1.DiagnosticSeverity.Warning);
                diagnostic.source = 'lint';
                diagnostics.push(diagnostic);
            }
            previousParam = param;
        }
        return diagnostics;
    }
}
exports.MethodParametersOnNewLine = MethodParametersOnNewLine;
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoicGFyYW1ldGVycy5qcyIsInNvdXJjZVJvb3QiOiIiLCJzb3VyY2VzIjpbIi4uLy4uLy4uL3NyYy9wc2xMaW50L3BhcmFtZXRlcnMudHMiXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6Ijs7O0FBQ0EsK0JBQW1FO0FBRW5FOztHQUVHO0FBQ0gsTUFBYSx5QkFBMEIsU0FBUSxnQkFBVTtJQUV4RCxNQUFNLENBQUMsTUFBYztRQUVwQixJQUFJLE1BQU0sQ0FBQyxLQUFLO1lBQUUsT0FBTyxFQUFFLENBQUM7UUFFNUIsTUFBTSxXQUFXLEdBQWlCLEVBQUUsQ0FBQztRQUNyQyxNQUFNLFVBQVUsR0FBRyxNQUFNLENBQUMsRUFBRSxDQUFDLFFBQVEsQ0FBQyxJQUFJLENBQUM7UUFFM0MsSUFBSSxhQUFvQyxDQUFDO1FBQ3pDLEtBQUssTUFBTSxLQUFLLElBQUksTUFBTSxDQUFDLFVBQVUsRUFBRTtZQUN0QyxNQUFNLGFBQWEsR0FBRyxLQUFLLENBQUMsRUFBRSxDQUFDLFFBQVEsQ0FBQztZQUN4QyxJQUFJLGFBQWEsSUFBSSxhQUFhLENBQUMsSUFBSSxLQUFLLGFBQWEsQ0FBQyxFQUFFLENBQUMsUUFBUSxDQUFDLElBQUksRUFBRTtnQkFDM0UsTUFBTSxPQUFPLEdBQUcsY0FBYyxLQUFLLENBQUMsRUFBRSxDQUFDLEtBQUssZ0NBQWdDLGFBQWEsQ0FBQyxFQUFFLENBQUMsS0FBSyxJQUFJLENBQUM7Z0JBQ3ZHLE1BQU0sVUFBVSxHQUFHLElBQUksZ0JBQVUsQ0FBQyxLQUFLLENBQUMsRUFBRSxDQUFDLFFBQVEsRUFBRSxFQUFFLE9BQU8sRUFBRSxJQUFJLENBQUMsUUFBUSxFQUFFLHdCQUFrQixDQUFDLE9BQU8sQ0FBQyxDQUFDO2dCQUMzRyxVQUFVLENBQUMsTUFBTSxHQUFHLE1BQU0sQ0FBQztnQkFDM0IsV0FBVyxDQUFDLElBQUksQ0FBQyxVQUFVLENBQUMsQ0FBQzthQUM3QjtpQkFDSSxJQUFJLE1BQU0sQ0FBQyxVQUFVLENBQUMsTUFBTSxHQUFHLENBQUMsSUFBSSxhQUFhLENBQUMsSUFBSSxLQUFLLFVBQVUsRUFBRTtnQkFDM0UsTUFBTSxPQUFPLEdBQUcsY0FBYyxLQUFLLENBQUMsRUFBRSxDQUFDLEtBQUssNEJBQTRCLE1BQU0sQ0FBQyxFQUFFLENBQUMsS0FBSyxJQUFJLENBQUM7Z0JBQzVGLE1BQU0sVUFBVSxHQUFHLElBQUksZ0JBQVUsQ0FBQyxLQUFLLENBQUMsRUFBRSxDQUFDLFFBQVEsRUFBRSxFQUFFLE9BQU8sRUFBRSxJQUFJLENBQUMsUUFBUSxFQUFFLHdCQUFrQixDQUFDLE9BQU8sQ0FBQyxDQUFDO2dCQUMzRyxVQUFVLENBQUMsTUFBTSxHQUFHLE1BQU0sQ0FBQztnQkFDM0IsV0FBVyxDQUFDLElBQUksQ0FBQyxVQUFVLENBQUMsQ0FBQzthQUM3QjtZQUNELGFBQWEsR0FBRyxLQUFLLENBQUM7U0FDdEI7UUFFRCxPQUFPLFdBQVcsQ0FBQztJQUNwQixDQUFDO0NBQ0Q7QUE3QkQsOERBNkJDIn0=