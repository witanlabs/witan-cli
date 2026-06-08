// Workbook, sheet, and defined-name APIs
// Witan xlsx-code-mode exec API reference. All functions are async and take `wb` first.

type PartialDeep<T>={[K in keyof T]?:NonNullable<T[K]> extends object?PartialDeep<NonNullable<T[K]>>:T[K]}
interface WorkbookProperties {
	activeSheetIndex:number;
	defaultFont:{
		name:string;
		size:number;
	};
	iterativeCalculation:{
		enabled:boolean;
		maxIterations:number;
		maxChange:number;
	};
	metadata?:{
		author?:string;
		title?:string;
		subject?:string;
		company?:string;
		created?:string;
		modified?:string;
	};
	themeColors?:{
		dark1:string;
		light1:string;
		dark2:string;
		light2:string;
		accent1:string;
		accent2:string;
		accent3:string;
		accent4:string;
		accent5:string;
		accent6:string;
		hyperlink:string;
		followedHyperlink:string;
		majorFont?:string;
		minorFont?:string;
	};
}
declare function getWorkbookProperties(wb):Promise<WorkbookProperties>;
declare function setWorkbookProperties(wb,properties:PartialDeep<WorkbookProperties>):Promise<void>;
declare function listSheets(wb):Promise<Array<{
	address:string;
	rows:number;
	cols:number;
	sheet:string;
	hidden?:boolean;
	printArea?:string;
	listObjects?:string[];
	dataTables?:string[];
	precedents?:string[];
	dependents?:string[];
}>>;

interface NameDef {name:string;range:string;scope:string|null}
declare function listDefinedNames(wb):Promise<NameDef[]>;
declare function addDefinedName(wb,name:string,range:string,scope?:string):Promise<NameDef>;
declare function deleteDefinedName(wb,name:string,scope?:string):Promise<NameDef>;
declare function addSheet(wb,name:string):Promise<string>;
declare function deleteSheet(wb,name:string):Promise<void>;
declare function renameSheet(wb,oldName:string,newName:string):Promise<void>;

type SheetVisibility="visible"|"hidden"|"veryHidden";
interface SheetView {showGridLines:boolean;zoomScale:number}
interface SheetOutline {summaryRowsBelow:boolean;summaryColumnsRight:boolean;showSymbols:boolean}
interface SheetFormat {defaultRowHeight:number;defaultColWidth:number;font?:{name?:string;size?:number}|null}
interface RowProps {
	height?:number;
	hidden?:boolean;
	outlineLevel?:number;
	collapsed?:boolean;
}
interface ColProps {
	width?:number;
	hidden?:boolean;
	outlineLevel?:number;
	collapsed?:boolean;
}
interface SheetProperties {
	visibility:SheetVisibility;
	view:SheetView;
	outline:SheetOutline;
	format:SheetFormat;
	columns:Record<string,ColProps&{col:string;width:number}>;
	rows:Record<number,RowProps&{row:number;height:number}>;
	merges?:string[]|null;
}
declare function setSheetProperties(wb,sheetName:string,properties:PartialDeep<Pick<SheetProperties,"visibility"|"view"|"outline"|"format">>&{merges?:string[]}):Promise<void>;
declare function setRowProperties(wb,sheetName:string,fromRow:number,toRow:number,properties:RowProps):Promise<void>;
declare function setColumnProperties(wb,sheetName:string,fromCol:number|string,toCol:number|string,properties:ColProps):Promise<void>;
declare function getSheetProperties(wb,sheetName:string,filter?:{
	columns?:(number|string)[];
	rows?:number[];
}):Promise<SheetProperties>;
