// Reading and search APIs
// Witan xlsx-code-mode exec API reference. All functions are async and take `wb` first.

type CellRef=string|{sheet:string;row:number;col:number|string}
type RangeRef=string|{sheet:string}|{sheet:string;from:{row?:number;col?:number|string};to:{row?:number;col?:number|string}}
type Visibility="visible"|"outsidePrintArea"|"collapsed"|"hidden"
interface Diag {code:string;detail?:string;address:string;formula?:string}

interface Value {
	address:string;
	sheet:string;
	row:number;
	col:number;
	colLetter:string;
	value:string|number|boolean|null;
	formula?:string;
	type:"string"|"number"|"bool"|"date"|"error"|"blank";
	text:string;
	format?:string;
	numberType?:"currency"|"percent"|"fraction"|"exponential"|"date"|"text"|"number";
	visibility:Visibility;
	context?:string;
	note?:{author:string;text:string};
	link?:{type:"internal"|"external";target:string;tooltip?:string};
	thread?:{
		resolved:boolean;
		comments:{authorId:string;text:string;createdAt:string}[];
	};
}
declare function readCell(wb,cell:CellRef,opts?:{
	context?:number;
}):Promise<Value>;
declare function readRange(wb,range:RangeRef):Promise<Value[][]>;
declare function readColumn(wb,sheetName:string,col:number|string,opts?:{
	startRow?:number;
	endRow?:number;
}):Promise<Value[]>;
declare function readRow(wb,sheetName:string,row:number,opts?:{
	startCol?:number;
	endCol?:number;
}):Promise<Value[]>;
declare function readRangeTsv(wb,range:RangeRef,opts?:{
	includeEmpty?:boolean;
	includeFormulas?:boolean;
}):Promise<string>;
declare function readColumnTsv(wb,sheetName:string,col:number|string,opts?:{
	startRow?:number;
	endRow?:number;
	includeEmpty?:boolean;
	includeFormulas?:boolean;
}):Promise<string>;
declare function readRowTsv(wb,sheetName:string,row:number,opts?:{
	startCol?:number;
	endCol?:number;
	includeEmpty?:boolean;
	includeFormulas?:boolean;
}):Promise<string>;
declare class SearchResults<T> extends Array<T> {truncated?:boolean}
type Matcher=string|string[]|number|boolean|RegExp|RegExp[]
/** Cell search by scalar, string list, or regex; `formulas:true` matches formulas only. */
declare function findCells(wb,matcher:Matcher,opts?:{in?:RangeRef|string;context?:number;limit?:number;offset?:number;formulas?:boolean}):Promise<{
	type:"cell";
	address:string;
	value:any;
	text:string;
	formula?:string;
	row:number;
	col:number;
	colLetter:string;
	sheet:string;
	visibility:Visibility;
	context?:string;
	role:string;
}[]>;
declare function findRows(wb,matcher:Matcher,opts?:{in?:RangeRef|string;context?:number;limit?:number;offset?:number}):Promise<{
	type:"row";
	row:number;
	sheet:string;
	matchedAt:string;
	range:string;
	tsv:string;
	visibility:Visibility;
	context?:string;
}[]>;
/** Replace text in values or formulas; use regex boundaries for formula edits. */
declare function findAndReplace(wb,find:string|RegExp,replace:string,opts?:{
		in?:RangeRef|string;
		matchCase?:boolean;
		wholeCell?:boolean;
		inFormulas?:boolean;
		limit?:number;
	}):Promise<{
	replaced:number;
	cells:string[];
	errors:Diag[];
}>;
declare function describeSheet(wb,sheetName:string):Promise<{
	tables:Record<string,{address:string;headerRows:string;headerCols:string|null;tableName?:string}>;
	structure:string; // Compact ASCII structure map
}>;
declare function tableLookup(wb,args:{
	table:string;
	rowLabel:string|number|boolean;
	columnLabel:string|number|boolean;
}):Promise<{
	address:string;
	value:any;
	text:string;
	row:number;
	col:number;
	colLetter:string;
	sheet:string;
	visibility:Visibility;
	rowLabelFoundAt:string;
	rowLabelFound:string;
	columnLabelFoundAt:string;
	columnLabelFound:string;
}[]>;
