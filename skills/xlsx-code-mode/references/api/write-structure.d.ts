// Cell, range, row, column, and sheet mutation APIs
// Witan xlsx-code-mode exec API reference. All functions are async and take `wb` first.

type CellRef=string|{sheet:string;row:number;col:number|string}
type RangeRef=string|{sheet:string}|{sheet:string;from:{row?:number;col?:number|string};to:{row?:number;col?:number|string}}
interface Diag {code:string;detail?:string;address:string;formula?:string}

interface WriteResult {
	touched:Record<string,string>;
	changed:string[];
	errors:Diag[];
	invalidatedTiles:{
		sheet:string;
		tileRow:number;
		tileCol:number;
	}[];
	updatedSheets:{
		name:string;
		usedRange:{
			startRow:number;
			startCol:number;
			endRow:number;
			endCol:number;
		}|null;
		tileRowCount:number;
		tileColCount:number;
	}[];
}
declare function setCells(wb,cells:Array<{
	address:CellRef;
	value?:unknown;
	formula?:string;
	format?:string;
	/** Legacy note payload; `null` clears. */
	note?:{text?:string;author?:string}|null;
	/** Hyperlink payload; `null` clears. */
	link?:{url?:string;ref?:string;tooltip?:string}|null;
	/** Thread payload: append via `add`, set `resolved`, or delete. */
	thread?:{add?:Array<{author?:string;text:string}>;resolved?:boolean;delete?:boolean}|null;
}>,opts?:{validationMode?:"ignore"|"reject"}):Promise<WriteResult>;

declare function scaleRange(wb,range:RangeRef,factor:number,opts?:{skipFormulas?:boolean}):Promise<WriteResult|null>;
declare function insertRowAfter(wb,sheetName:string,row:number,count?:number):Promise<void>;
declare function deleteRows(wb,sheetName:string,row:number,count?:number):Promise<void>;
declare function insertColumnAfter(wb,sheetName:string,column:number|string,count?:number):Promise<void>;
declare function deleteColumns(wb,sheetName:string,column:number|string,count?:number):Promise<void>;
declare function autoFitColumns(wb,sheetName:string,columns?:Array<number|string>,opts?:{
	minWidth?:number;
	maxWidth?:number;
	padding?:number;
}):Promise<Record<string,{width:number;previousWidth:number}>>;
declare function sortRange(wb,range:RangeRef,keys:Array<{
	col:number|string;
	order?:"asc"|"desc";
}>,opts?:{hasHeader?:boolean}):Promise<void>;
declare function copyRange(wb,source:RangeRef,destination:CellRef,opts?:{
	pasteType?:"all"|"values"|"formulas"|"formats";
}):Promise<{destination:string;cellsCopied:number;}>;
