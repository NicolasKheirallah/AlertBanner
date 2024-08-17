"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const parser = require("../src/parser/parser");
function getMethod(methodString) {
    const d = parser.parseText(methodString);
    return d.methods[0];
}
function getParsedDoc(documentString) {
    return parser.parseText(documentString);
}
function toValues(tokens) {
    return tokens.map(t => t.value);
}
function argsToValues(args) {
    return args.map(a => a.types).map(ts => toValues(ts));
}
function argsToNames(args) {
    return toValues(args.map(a => a.id));
}
describe('Batch label', () => {
    const batchText = `---------- OPEN ------ Section marker

	type public Boolean ER
	type public Number BRCD
	type public String ET, RM

	do SOURCE^BCHSOURC("BOFF", "ACCUPD", .%UserID, .BRCD, .%UserClass)

	// ~p1 source not set up
	if ER set RM = $$^MSG(1184,"BOFF-ACCUPD"), %BatchExit = 1 do EXC quit`;
    const d = getParsedDoc(batchText);
    expect(d.methods).toHaveLength(1);
});
describe('Method Identifiers', () => {
    test('inline label statement symbol', () => {
        const methodString = 'label do something^SOMETHING';
        const result = getMethod(methodString);
        if (!result) {
            fail();
            return;
        }
        // let identifierValues = toValues(result.modifiers)
        expect(result.id.value).toEqual('label');
    });
    test('inline label statement keyword', () => {
        const methodString = 'label do something()';
        const result = getMethod(methodString);
        if (!result) {
            fail();
            return;
        }
        // let identifierValues = toValues(result.modifiers)
        expect(result.id.value).toEqual('label');
    });
    test('1 argument', () => {
        const methodString = 'public static void main(String args)';
        const result = getMethod(methodString);
        if (!result) {
            fail();
            return;
        }
        const identifierValues = toValues(result.modifiers);
        expect(identifierValues).toEqual(['public', 'static']);
    });
    test('2 arguments', () => {
        const methodString = 'public static void main(String arg1, String arg2)';
        const result = getMethod(methodString);
        if (!result) {
            fail();
            return;
        }
        const identifierValues = toValues(result.modifiers);
        expect(identifierValues).toEqual(['public', 'static']);
    });
    test('Label', () => {
        const methodString = 'main';
        const result = getMethod(methodString);
        if (!result) {
            fail();
            return;
        }
        // let identifierValues = toValues(result.modifiers)
        expect(result.id.value).toEqual('main');
    });
    test('Label from document', () => {
        const methodString = 'main\r\n';
        const result = getParsedDoc(methodString);
        if (!result) {
            fail();
            return;
        }
        expect(result.methods[0].id.value).toEqual('main');
    });
    test('Label with line comment', () => {
        const methodString = 'main // a comment';
        const result = getMethod(methodString);
        if (!result) {
            fail();
            return;
        }
        toValues(result.modifiers);
        expect(result.id.value).toEqual('main');
    });
    test('Label with parens', () => {
        const methodString = 'main()';
        const result = getMethod(methodString);
        if (!result) {
            fail();
            return;
        }
        toValues(result.modifiers);
        expect(result.id.value).toEqual('main');
    });
    test('Label with 1 argument', () => {
        const methodString = 'main(String x1)';
        const result = getMethod(methodString);
        if (!result) {
            fail();
            return;
        }
        toValues(result.modifiers);
        expect(result.id.value).toEqual('main');
    });
    test('Label with 2 arguments', () => {
        const methodString = 'main(String x1, String x2)';
        const result = getMethod(methodString);
        if (!result) {
            fail();
            return;
        }
        toValues(result.modifiers);
        expect(result.id.value).toEqual('main');
    });
    test('Label with 2 arguments multiline', () => {
        const methodString = 'main(String x1\n\t, String x2)';
        const result = getMethod(methodString);
        if (!result) {
            fail();
            return;
        }
        toValues(result.modifiers);
        expect(result.id.value).toEqual('main');
    });
    test('percent', () => {
        const methodString = 'public %main()';
        const method = getMethod(methodString);
        expect(method.id.value).toEqual('%main');
    });
});
describe('Argument Names', () => {
    test('1 argument', () => {
        const methodString = 'public static void main(String x1)';
        const result = getMethod(methodString);
        if (!result) {
            fail();
            return;
        }
        const argNameValues = argsToNames(result.parameters);
        expect(argNameValues).toEqual(['x1']);
    });
    test('2 arguments', () => {
        const methodString = 'public static void main(String x1, String x2)';
        const result = getMethod(methodString);
        if (!result) {
            fail();
            return;
        }
        const argNameValues = argsToNames(result.parameters);
        expect(argNameValues).toEqual(['x1', 'x2']);
    });
    test('1 argument multiline', () => {
        const methodString = 'public static void main(\n\tString x1)';
        const result = getMethod(methodString);
        if (!result) {
            fail();
            return;
        }
        const argNameValues = argsToNames(result.parameters);
        expect(argNameValues).toEqual(['x1']);
    });
    test('2 argument multiline', () => {
        const methodString = 'public static void main(String x1,\n\tString x2)';
        const result = getMethod(methodString);
        if (!result) {
            fail();
            return;
        }
        const argNameValues = argsToNames(result.parameters);
        expect(argNameValues).toEqual(['x1', 'x2']);
    });
    test('1 argument multitype', () => {
        const methodString = 'public static void main(void x1(Integer, Record))';
        const result = getMethod(methodString);
        if (!result) {
            fail();
            return;
        }
        const argNameValues = argsToNames(result.parameters);
        expect(argNameValues).toEqual(['x1']);
    });
    test('2 argument multitype', () => {
        const methodString = 'public static void main(void x1(Integer, Record), void x2(void, String))';
        const result = getMethod(methodString);
        if (!result) {
            fail();
            return;
        }
        const argNameValues = argsToNames(result.parameters);
        expect(argNameValues).toEqual(['x1', 'x2']);
    });
    test('2 argument multitype', () => {
        const methodString = 'public static void main(void x1(Integer, Record)\n\t, void x2(void, String))';
        const result = getMethod(methodString);
        if (!result) {
            fail();
            return;
        }
        const argNameValues = argsToNames(result.parameters);
        expect(argNameValues).toEqual(['x1', 'x2']);
    });
    test('test label with parens 1 arg', () => {
        const methodString = 'main(String x1)';
        const result = getMethod(methodString);
        if (!result) {
            fail();
            return;
        }
        const argNameValues = argsToNames(result.parameters);
        expect(argNameValues).toEqual(['x1']);
    });
    test('Label no args', () => {
        const methodString = 'main';
        const result = getMethod(methodString);
        if (!result) {
            fail();
            return;
        }
        const args = result.parameters;
        expect(args).toHaveLength(0);
    });
    test('Label with parens no args', () => {
        const methodString = 'main()';
        const result = getMethod(methodString);
        if (!result) {
            fail();
            return;
        }
        const args = result.parameters;
        expect(args).toHaveLength(0);
    });
    test('Label with multiline parens no args', () => {
        const methodString = 'main(\n\t)';
        const result = getMethod(methodString);
        if (!result) {
            fail();
            return;
        }
        const args = result.parameters;
        expect(args).toHaveLength(0);
    });
});
describe('Argument Types', () => {
    test('1 argument', () => {
        const methodString = 'public static void main(String x1)';
        const result = getMethod(methodString);
        if (!result) {
            fail();
            return;
        }
        const argValues = argsToValues(result.parameters);
        expect(argValues).toEqual([['String']]);
    });
    test('1 argument multitype', () => {
        const methodString = 'public static void main(String x1(Number))';
        const result = getMethod(methodString);
        if (!result) {
            fail();
            return;
        }
        const argValues = argsToValues(result.parameters);
        expect(argValues).toEqual([['String', 'Number']]);
    });
    test('test 2 argument types newline', () => {
        const methodString = 'public static void main(String x1 \n\t, Number x2)';
        const result = getMethod(methodString);
        if (!result) {
            fail();
            return;
        }
        const argValues = argsToValues(result.parameters);
        expect(argValues).toEqual([['String'], ['Number']]);
    });
    test('test 1 argument 3 types newline', () => {
        const methodString = 'public static void main(void x1(Integer, Record))';
        const result = getMethod(methodString);
        if (!result) {
            fail('Did not parse');
            return;
        }
        const argValues = argsToValues(result.parameters);
        expect(argValues).toEqual([['void', 'Integer', 'Record']]);
    });
    test('test 2 argument 3 types newline', () => {
        const methodString = 'public static void main(void x1(Integer, Record), void x2(void, String))';
        const result = getMethod(methodString);
        if (!result) {
            fail('Did not parse');
            return;
        }
        const argValues = argsToValues(result.parameters);
        expect(argValues).toEqual([['void', 'Integer', 'Record'], ['void', 'void', 'String']]);
    });
});
describe('Propertydefs', () => {
    test('empty propertydef', () => {
        const propertyString = '\t#PROPERTYDEF';
        const doc = getParsedDoc(propertyString);
        expect(doc.properties).toHaveLength(0);
    });
    test('one word propertydef', () => {
        const propertyString = '\t#PROPERTYDEF test';
        const doc = getParsedDoc(propertyString);
        expect(doc.properties).toHaveLength(1);
    });
});
test('parse document method count', () => {
    const documentString = `	#PACKAGE custom.core
	#CLASSDEF extends = Primitive public

	/*DOC -----------------------------------------------------------------
	Auto-generated by vscode-psl
	** ENDDOC */


	// --------------------------------------------------------------------
public final Integer toInteger()
	/*DOC -----------------------------------------------------------------
	convert Boolean to Integer
	** ENDDOC */
	do prim2prim^UCPRIM("Integer")
	quit


	// --------------------------------------------------------------------
public final Number toNumber()
	/*DOC -----------------------------------------------------------------
	convert Boolean to Number
	** ENDDOC */
	do prim2prim^UCPRIM("Number")
	quit


	// --------------------------------------------------------------------
public final String toString(String vMask)
	/*DOC -----------------------------------------------------------------
	convert Boolean to String
	** ENDDOC */
	do insMet^UCMETHOD("$$toString^PslNllBoolean(",1)
	quit
`;
    const doc = getParsedDoc(documentString);
    expect(doc.methods).toHaveLength(3);
});
test('parse extends', () => {
    const documentString = `	#PACKAGE custom.core
	#CLASSDEF extends = Primitive public

	/*DOC -----------------------------------------------------------------
	Auto-generated by vscode-psl
	** ENDDOC */


	// --------------------------------------------------------------------
public final Integer toInteger()
	/*DOC -----------------------------------------------------------------
	convert Boolean to Integer
	** ENDDOC */
	do prim2prim^UCPRIM("Integer")
	quit
`;
    const doc = getParsedDoc(documentString);
    expect(doc.extending.value).toBe('Primitive');
});
test('parse psl package', () => {
    const documentString = `	#PACKAGE custom.core
	#CLASSDEF extends = Primitive public

	/*DOC -----------------------------------------------------------------
	Auto-generated by vscode-psl
	** ENDDOC */


	// --------------------------------------------------------------------
public final Integer toInteger()
	/*DOC -----------------------------------------------------------------
	convert Boolean to Integer
	** ENDDOC */
	do prim2prim^UCPRIM("Integer")
	quit
`;
    const doc = getParsedDoc(documentString);
    expect(doc.pslPackage).toBe('custom.core');
});
test('parse numerical method', () => {
    const documentString = `	#PACKAGE custom.core
	#CLASSDEF extends = Primitive public

	/*DOC -----------------------------------------------------------------
	Auto-generated by vscode-psl
	** ENDDOC */


	// --------------------------------------------------------------------
public final Integer 900()
	/*DOC -----------------------------------------------------------------
	convert Boolean to Integer
	** ENDDOC */
	do prim2prim^UCPRIM("Integer")
	quit
`;
    const doc = getParsedDoc(documentString);
    expect(doc.methods[0].id.value).toBe('900');
});
test('parse document method location', () => {
    const documentString = `	#PACKAGE custom.core
	#CLASSDEF extends = Primitive public

	/*DOC -----------------------------------------------------------------
	Auto-generated by vscode-psl
	** ENDDOC */


	// --------------------------------------------------------------------
public final Integer toInteger()
	/*DOC -----------------------------------------------------------------
	convert Boolean to Integer
	** ENDDOC */
	do prim2prim^UCPRIM("Integer")
	quit


	// --------------------------------------------------------------------
public final Number toNumber()
	/*DOC -----------------------------------------------------------------
	convert Boolean to Number
	** ENDDOC */
	do prim2prim^UCPRIM("Number")
	quit


	// --------------------------------------------------------------------
public final String toString(String vMask)
	/*DOC -----------------------------------------------------------------
	convert Boolean to String
	** ENDDOC */
	do insMet^UCMETHOD("$$toString^PslNllBoolean(",1)
	quit
`;
    const doc = getParsedDoc(documentString);
    expect(doc.methods.map(method => method.line)).toEqual([9, 18, 27]);
});
test('labels in document', () => {
    const documentString = `	#PACKAGE custom.core
	#CLASSDEF extends = Primitive public

	/*DOC -----------------------------------------------------------------
	Auto-generated by vscode-psl
	** ENDDOC */



toInteger
	/*DOC -----------------------------------------------------------------
	convert Boolean to Integer
	** ENDDOC */
	do prim2prim^UCPRIM("Integer")
	quit


	// --------------------------------------------------------------------
toNumber
	/*DOC -----------------------------------------------------------------
	convert Boolean to Number
	** ENDDOC */
	do prim2prim^UCPRIM("Number")
	quit


	// --------------------------------------------------------------------
toString
	/*DOC -----------------------------------------------------------------
	convert Boolean to String
	** ENDDOC */
	do insMet^UCMETHOD("$$toString^PslNllBoolean(",1)
	quit
`;
    const doc = getParsedDoc(documentString);
    expect(doc.methods.map(method => method.id.value)).toEqual(['toInteger', 'toNumber', 'toString']);
    expect(doc.methods.map(method => method.line)).toEqual([9, 18, 27]);
});
test('parse methods with propertydef', () => {
    const documentString = `	#PACKAGE custom.core
	#CLASSDEF extends = Primitive public

	/*DOC -----------------------------------------------------------------
	Auto-generated by vscode-psl
	** ENDDOC */

	#PROPERTYDEF test class = String node = 1 public


	// --------------------------------------------------------------------
public final Integer toInteger()
	/*DOC -----------------------------------------------------------------
	convert Boolean to Integer
	** ENDDOC */
	do prim2prim^UCPRIM("Integer")
	quit


	// --------------------------------------------------------------------
public final Number toNumber()
	/*DOC -----------------------------------------------------------------
	convert Boolean to Number
	** ENDDOC */
	do prim2prim^UCPRIM("Number")
	quit


	// --------------------------------------------------------------------
public final String toString(String vMask)
	/*DOC -----------------------------------------------------------------
	convert Boolean to String
	** ENDDOC */
	do insMet^UCMETHOD("$$toString^PslNllBoolean(",1)
	quit
`;
    const doc = getParsedDoc(documentString);
    expect(doc.methods).toHaveLength(3);
});
test('parse methods with propertydef count', () => {
    const documentString = `	#PACKAGE custom.core
	#CLASSDEF extends = Primitive public

	/*DOC -----------------------------------------------------------------
	Auto-generated by vscode-psl
	** ENDDOC */

	#PROPERTYDEF test class = String node = 1 public


	// --------------------------------------------------------------------
public final Integer toInteger()
	/*DOC -----------------------------------------------------------------
	convert Boolean to Integer
	** ENDDOC */
	do prim2prim^UCPRIM("Integer")
	quit


	// --------------------------------------------------------------------
public final Number toNumber()
	/*DOC -----------------------------------------------------------------
	convert Boolean to Number
	** ENDDOC */
	do prim2prim^UCPRIM("Number")
	quit


	// --------------------------------------------------------------------
public final String toString(String vMask)
	/*DOC -----------------------------------------------------------------
	convert Boolean to String
	** ENDDOC */
	do insMet^UCMETHOD("$$toString^PslNllBoolean(",1)
	quit
`;
    const doc = getParsedDoc(documentString);
    expect(doc.properties).toHaveLength(1);
});
test('parse methods with propertydef count', () => {
    const documentString = `	#PACKAGE custom.core
	#CLASSDEF extends = Primitive public

	/*DOC -----------------------------------------------------------------
	Auto-generated by vscode-psl
	** ENDDOC */

	#PROPERTYDEF test class = String node = 1 public


	// --------------------------------------------------------------------
public final Integer toInteger()
	/*DOC -----------------------------------------------------------------
	convert Boolean to Integer
	** ENDDOC */
	do prim2prim^UCPRIM("Integer")
	quit


	// --------------------------------------------------------------------
public final Number toNumber()
	/*DOC -----------------------------------------------------------------
	convert Boolean to Number
	** ENDDOC */
	do prim2prim^UCPRIM("Number")
	quit


	// --------------------------------------------------------------------
public final String toString(String vMask)
	/*DOC -----------------------------------------------------------------
	convert Boolean to String
	** ENDDOC */
	do insMet^UCMETHOD("$$toString^PslNllBoolean(",1)
	quit
`;
    const doc = getParsedDoc(documentString);
    expect(toValues(doc.properties[0].modifiers)).toEqual(['public']);
    expect(doc.properties[0].id.value).toEqual('test');
});
describe('type declarations', () => {
    test('basic type declaration', () => {
        const declarationString = '\ttype public literal String x = "hi there"';
        const doc = getParsedDoc(declarationString);
        expect(doc.declarations[0].types[0].value).toEqual('String');
        expect(doc.declarations[0].id.value).toEqual('x');
    });
    test('mutliple type declaration', () => {
        const declarationString = '\ttype public literal String x,y';
        const doc = getParsedDoc(declarationString);
        expect(doc.declarations[0].types[0].value).toEqual('String');
        expect(doc.declarations[0].id.value).toEqual('x');
        expect(doc.declarations[1].types[0].value).toEqual('String');
        expect(doc.declarations[1].id.value).toEqual('y');
    });
    test('mutliple multitype type declaration', () => {
        const declarationString = '\ttype public literal String x(Number,Boolean),y';
        const doc = getParsedDoc(declarationString);
        expect(doc.declarations[0].types[0].value).toEqual('String');
        expect(doc.declarations[0].types[1].value).toEqual('Number');
        expect(doc.declarations[0].types[2].value).toEqual('Boolean');
        expect(doc.declarations[0].id.value).toEqual('x');
        expect(doc.declarations[1].types[0].value).toEqual('String');
        expect(doc.declarations[1].id.value).toEqual('y');
    });
    test('mutliple type declaration equal sign', () => {
        const declarationString = '\ttype String x = "hi", y = "hi"';
        const doc = getParsedDoc(declarationString);
        expect(doc.declarations[0].types[0].value).toEqual('String');
        expect(doc.declarations[0].id.value).toEqual('x');
        expect(doc.declarations[1].types[0].value).toEqual('String');
        expect(doc.declarations[1].id.value).toEqual('y');
    });
    test('static type declaration', () => {
        const declarationString = '\ttype static x';
        const doc = getParsedDoc(declarationString);
        expect(doc.declarations[0].types[0].value).toEqual('x');
        expect(doc.declarations[0].id.value).toEqual('x');
    });
    test('type type declaration', () => {
        const declarationString = '\ttype String type';
        const doc = getParsedDoc(declarationString);
        expect(doc.declarations[0].types[0].value).toEqual('String');
        expect(doc.declarations[0].id.value).toEqual('type');
    });
    test('method declarations', () => {
        const documentString = `
public static void main()
	type String x
	quit

public static void main2()
	type Number y
	quit
`;
        const doc = getParsedDoc(documentString);
        expect(doc.methods[0].declarations[0].id.value).toEqual('x');
        expect(doc.methods[1].declarations[0].id.value).toEqual('y');
    });
});
//# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJmaWxlIjoicGFyc2VyLXRlc3QuanMiLCJzb3VyY2VSb290IjoiIiwic291cmNlcyI6WyIuLi8uLi9fX3Rlc3RzX18vcGFyc2VyLXRlc3QudHMiXSwibmFtZXMiOltdLCJtYXBwaW5ncyI6Ijs7QUFBQSwrQ0FBK0M7QUFHL0MsU0FBUyxTQUFTLENBQUMsWUFBb0I7SUFDdEMsTUFBTSxDQUFDLEdBQUcsTUFBTSxDQUFDLFNBQVMsQ0FBQyxZQUFZLENBQUMsQ0FBQztJQUN6QyxPQUFPLENBQUMsQ0FBQyxPQUFPLENBQUMsQ0FBQyxDQUFDLENBQUM7QUFDckIsQ0FBQztBQUNELFNBQVMsWUFBWSxDQUFDLGNBQXNCO0lBQzNDLE9BQU8sTUFBTSxDQUFDLFNBQVMsQ0FBQyxjQUFjLENBQUMsQ0FBQztBQUN6QyxDQUFDO0FBRUQsU0FBUyxRQUFRLENBQUMsTUFBeUI7SUFDMUMsT0FBTyxNQUFNLENBQUMsR0FBRyxDQUFDLENBQUMsQ0FBQyxFQUFFLENBQUMsQ0FBQyxDQUFDLEtBQUssQ0FBQyxDQUFDO0FBQ2pDLENBQUM7QUFFRCxTQUFTLFlBQVksQ0FBQyxJQUF3QjtJQUM3QyxPQUFPLElBQUksQ0FBQyxHQUFHLENBQUMsQ0FBQyxDQUFDLEVBQUUsQ0FBQyxDQUFDLENBQUMsS0FBSyxDQUFDLENBQUMsR0FBRyxDQUFDLEVBQUUsQ0FBQyxFQUFFLENBQUMsUUFBUSxDQUFDLEVBQUUsQ0FBQyxDQUFDLENBQUM7QUFDdkQsQ0FBQztBQUVELFNBQVMsV0FBVyxDQUFDLElBQXdCO0lBQzVDLE9BQU8sUUFBUSxDQUFDLElBQUksQ0FBQyxHQUFHLENBQUMsQ0FBQyxDQUFDLEVBQUUsQ0FBQyxDQUFDLENBQUMsRUFBRSxDQUFDLENBQUMsQ0FBQztBQUN0QyxDQUFDO0FBRUQsUUFBUSxDQUFDLGFBQWEsRUFBRSxHQUFHLEVBQUU7SUFDNUIsTUFBTSxTQUFTLEdBQUc7Ozs7Ozs7Ozt1RUFTb0QsQ0FBQztJQUV2RSxNQUFNLENBQUMsR0FBRyxZQUFZLENBQUMsU0FBUyxDQUFDLENBQUM7SUFDbEMs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LElBQUksRUFBRSxDQUFDO1lBQ1AsT0FBTztTQUNQO1FBQ0QsTUFBTSxTQUFTLEdBQUcsWUFBWSxDQUFDLE1BQU0sQ0FBQyxVQUFVLENBQUMsQ0FBQztRQUNsRCxNQUFNLENBQUMsU0FBUyxDQUFDLENBQUMsT0FBTyxDQUFDLENBQUMsQ0FBQyxRQUFRLENBQUMsRUFBRSxDQUFDLFFBQVEsQ0FBQyxDQUFDLENBQUMsQ0FBQztJQUNyRCxDQUFDLENBQUMsQ0FBQztJQUVILElBQUksQ0FBQyxpQ0FBaUMsRUFBRSxHQUFHLEVBQUU7UUFDNUMsTUFBTSxZQUFZLEdBQUcsbURBQW1ELENBQUM7UUFDekUsTUFBTSxNQUFNLEdBQUcsU0FBUyxDQUFDLFlBQVksQ0FBQyxDQUFDO1FBQ3ZDLElBQUksQ0FBQyxNQUFNLEVBQUU7WUFDWixJQUFJLENBQUMsZUFBZSxDQUFDLENBQUM7WUFDdEIsT0FBTztTQUNQO1FBQ0QsTUFBTSxTQUFTLEdBQUcsWUFBWSxDQUFDLE1BQU0sQ0FBQyxVQUFVLENBQUMsQ0FBQztRQUNsRCxNQUFNLENBQUMsU0FBUyxDQUFDLENBQUMsT0FBTyxDQUFDLENBQUMsQ0FBQyxNQUFNLEVBQUUsU0FBUyxFQUFFLFFBQVEsQ0FBQyxDQUFDLENBQUMsQ0FBQztJQUM1RCxDQUFDLENBQUMsQ0FBQztJQUVILElBQUksQ0FBQyxpQ0FBaUMsRUFBRSxHQUFHLEVBQUU7UUFDNUMsTUFBTSxZQUFZLEdBQUcsMEVBQTBFLENBQUM7UUFDaEcsTUFBTSxNQUFNLEdBQUcsU0FBUyxDQUFDLFlBQVksQ0FBQyxDQUFDO1FBQ3ZDLElBQUksQ0FBQyxNQUFNLEVBQUU7WUFDWixJQUFJLENBQUMsZUFBZSxDQUFDLENBQUM7WUFDdEIsT0FBTztTQUNQO1FBQ0QsTUFBTSxTQUFTLEdBQUcsWUFBWSxDQUFDLE1BQU0sQ0FBQyxVQUFVLENBQUMsQ0FBQztRQUNsRCxNQUFNLENBQUMsU0FBUyxDQUFDLENBQUMsT0FBTyxDQUFDLENBQUMsQ0FBQyxNQUFNLEVBQUUsU0FBUyxFQUFFLFFBQVEsQ0FBQyxFQUFFLENBQUMsTUFB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