// What-if, tracing, formula evaluation, linting, and preview APIs
// Witan xlsx-code-mode exec API reference. All functions are async and take `wb` first.

type CellRef=string|{sheet:string;row:number;col:number|string}
type RangeRef=string|{sheet:string}|{sheet:string;from:{row?:number;col?:number|string};to:{row?:number;col?:number|string}}
type Visibility="visible"|"outsidePrintArea"|"collapsed"|"hidden"
interface Diag {code:string;detail?:string;address:string;formula?:string}

declare function sweepInputs(wb,args:{
	inputs:{address:CellRef;values:(number|string|boolean|null)[]}[];
	outputs:(string|CellRef)[];
	mode?:"cartesian"|"parallel";
	includeStats?:boolean;
}):Promise<{
	tsv:string;
	sweeps:{inputs:Record<string,string>;outputs:Record<string,string>;errors:Diag[];}[];
	stats?:Record<string,{min:number;max:number;mean:number;count:number}>;
	sweepCount:number;
	inputCount:number;
	outputCount:number;
}>;

interface Deps {cells:{address:string;depth:number;formula?:string;referenceType?:"direct"|"range"|"named"|"table"}[];warnings?:Diag[]}
declare function getCellPrecedents(wb,address:CellRef,depth?:number):Promise<Deps>;
declare function getCellDependents(wb,address:CellRef,depth?:number):Promise<Deps>;
declare function traceToInputs(wb,cell:CellRef):Promise<{
	address:string;
	referenceCount:number;
	text?:string;
	nearbyLabel?:string;
	context?:string;
}[]>;
declare function traceToOutputs(wb,cell:CellRef):Promise<{
	address:string;
	formula?:string;
	text?:string;
	visibility:Visibility;
	nearbyLabel?:string;
	context?:string;
}[]>;
interface FormulaEval {formula:string;value:number|string|boolean|null|unknown[][];error?:{code:string;detail?:string}}
declare function evaluateFormulas(wb,sheet:string,formulas:string[]):Promise<FormulaEval[]>;
declare function evaluateFormula(wb,sheet:string,formula:string):Promise<FormulaEval>;
declare function lint(wb,options?:{
	rangeAddresses?:string[];
	skipRuleIds?:string[];
	onlyRuleIds?:string[];
}):Promise<{
	diagnostics:{
		severity:"Info"|"Warning"|"Error";
		ruleId:string;
		message:string;
		location:string|null;
		visibility:Visibility|null;
	}[];
	total:number;
}>;
/** Generates a PNG and prints its path to stdout. */
declare function previewStyles(wb,range:RangeRef):Promise<void>;
