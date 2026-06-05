// Workbook, sheet, and defined-name APIs
// Witan xlsx-code-mode exec API reference. All functions are async and take `wb` first.

declare function getWorkbookProperties(wb):Promise<{
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
}>;
declare function setWorkbookProperties(wb,properties:{
	activeSheetIndex?:number;
	defaultFont?:{
		name?:string;
		size?:number;
	};
	iterativeCalculation?:{
		enabled?:boolean;
		maxIterations?:number;
		maxChange?:number;
	};
	metadata?:{
		author?:string;
		title?:string;
		subject?:string;
		company?:string;
	};
	themeColors?:{
		dark1?:string;
		light1?:string;
		dark2?:string;
		light2?:string;
		accent1?:string;
		accent2?:string;
		accent3?:string;
		accent4?:string;
		accent5?:string;
		accent6?:string;
		hyperlink?:string;
		followedHyperlink?:string;
		majorFont?:string;
		minorFont?:string;
	};
}):Promise<void>;
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

declare function setSheetProperties(wb,sheetName:string,properties:{
	visibility?:"visible"|"hidden"|"veryHidden";
	view?:{
		showGridLines?:boolean;
		zoomScale?:number;
	};
	outline?:{
		summaryRowsBelow?:boolean;
		summaryColumnsRight?:boolean;
		showSymbols?:boolean;
	};
	format?:{
		defaultRowHeight?:number;
		defaultColWidth?:number;
		font?:{
			name?:string;
			size?:number;
		};
	};
	merges?:string[];
}):Promise<void>;
declare function setRowProperties(wb,sheetName:string,fromRow:number,toRow:number,properties:{
	height?:number;
	hidden?:boolean;
	outlineLevel?:number;
	collapsed?:boolean;
}):Promise<void>;
declare function setColumnProperties(wb,sheetName:string,fromCol:number|string,toCol:number|string,properties:{
	width?:number;
	hidden?:boolean;
	outlineLevel?:number;
	collapsed?:boolean;
}):Promise<void>;
declare function getSheetProperties(wb,sheetName:string,filter?:{
	columns?:(number|string)[];
	rows?:number[];
}):Promise<{
	visibility:"visible"|"hidden"|"veryHidden";
	view:{
		showGridLines:boolean;
		zoomScale:number;
	};
	outline:{
		summaryRowsBelow:boolean;
		summaryColumnsRight:boolean;
		showSymbols:boolean;
	};
	format:{
		defaultRowHeight:number;
		defaultColWidth:number;
		font?:{
			name?:string;
			size?:number;
		}|null;
	};
	columns:Record<
		string,
		{
			col:string;
			width:number;
			hidden?:boolean;
			outlineLevel?:number;
			collapsed?:boolean;
		}
	>;
	rows:Record<
		number,
		{
			row:number;
			height:number;
			hidden?:boolean;
			outlineLevel?:number;
			collapsed?:boolean;
		}
	>;
	merges?:string[]|null;
}>;
