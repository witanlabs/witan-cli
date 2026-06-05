// Witan xlsx API — type reference for `witan xlsx exec`.
//
// Inside an exec script two globals are provided for you:
//   xlsx  — this namespace; every function below is reached as `xlsx.<name>(...)`
//   wb    — an already-open workbook handle; pass it as the FIRST argument to every call
//
// Lifecycle is handled by the CLI, not your script:
//   - `--create` (flag) starts a new workbook session
//   - `--save`   (flag) persists changes to disk
//
// All functions are async; use top-level `await`. Imports are not available.

type CellRef=string|{sheet:string;row:number;col:number|string}
type RangeRef=string|{sheet:string}|{sheet:string;from:{row?:number;col?:number|string};to:{row?:number;col?:number|string}}
type Visibility="visible"|"outsidePrintArea"|"collapsed"|"hidden"
function getWorkbookProperties(wb):Promise<{
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
function setWorkbookProperties(wb,properties:{
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
function listSheets(wb):Promise<Array<{
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
type TimeUnit="days"|"months"|"years"
type ChartDataLabelPosition="bestFit"|"center"|"insideBase"|"insideEnd"|"outsideEnd"|"left"|"right"|"top"|"bottom"
/** Use exactly one of `text` or `ref`. */
interface ChartTextSource {text?:string;ref?:string}
interface ChartPositionAnchor {cell:string;xOffsetPts?:number;yOffsetPts?:number}
interface ChartPosIn {from:ChartPositionAnchor;to:ChartPositionAnchor}
interface ChartPos extends ChartPosIn {sheet:string}
interface ChartFillFormat {noFill?:boolean;color?:string}
interface ChartLineFormat {noLine?:boolean;color?:string;weight?:number;lineStyle?:string}
interface ChartFontFormat {bold?:boolean;color?:string;italic?:boolean;name?:string;size?:number;underline?:string}
interface ChartDataLabelFormat {fill?:ChartFillFormat;border?:ChartLineFormat;font?:ChartFontFormat}
interface ChartPlotAreaSpec {format?:{fill?:ChartFillFormat;border?:ChartLineFormat}}
interface ChartAxisSpec {
	title?:ChartTextSource;
	visible?:boolean;
	categoryType?:"category"|"date";
	min?:number;
	max?:number;
	majorUnit?:number;
	minorUnit?:number;
	baseTimeUnit?:TimeUnit;
	majorTimeUnit?:TimeUnit;
	minorTimeUnit?:TimeUnit;
	numberFormat?:string;
	numberFormatLinked?:boolean;
	reversed?:boolean;
	majorGridlines?:boolean;
	minorGridlines?:boolean;
	position?:"left"|"right"|"top"|"bottom";
}
interface ChartSpec {
	name:string;
	position:ChartPosIn;
	groups:{
		type:"column"|"bar"|"line"|"area"|"pie"|"doughnut"|"scatter"|"bubble"|"radar"|"stockHLC"|"stockOHLC"|"waterfall";
		scatterStyle?:"line"|"lineMarker"|"marker"|"smooth"|"smoothMarker"; /** scatter only */
		radarStyle?:"standard"|"marker"|"filled"; /** radar only */
		grouping?:"standard"|"stacked"|"percentStacked";
		axis?:"primary"|"secondary";
		gapWidth?:number;
		overlap?:number;
		varyColors?:boolean;
		smooth?:boolean;
		firstSliceAngle?:number;
		holeSize?:number;
		bubbleScale?:number; /** bubble only, 0-300 */
		showNegativeBubbles?:boolean; /** bubble only */
		sizeRepresents?:"area"|"width"; /** bubble only */
		dataLabels?:{
			showLegendKey?:boolean;
			showValue?:boolean;
			showCategory?:boolean;
			showSeriesName?:boolean;
			showPercent?:boolean;
			showBubbleSize?:boolean;
			showLeaderLines?:boolean;
			position?:ChartDataLabelPosition;
			numberFormat?:string;
			numberFormatLinked?:boolean;
			separator?:string;
			format?:ChartDataLabelFormat;
		};
		series:{
			name?:ChartTextSource;
			stockRole?:"volume"|"open"|"high"|"low"|"close"; /** stock charts only */
			categories?:string;
			categoriesRefType?:"string"|"number"|"multiLevelString";
			values?:string;
			xValues?:string;
			yValues?:string;
			bubbleSizes?:string; /** bubble only */
			fillColor?:string;
			lineColor?:string;
			lineWidth?:number;
			lineDashStyle?:string;
			smooth?:boolean;
			invertIfNegative?:boolean;
			totalIndexes?:number[]; /** waterfall only: zero-based subtotal/total point indexes */
			showConnectorLines?:boolean; /** waterfall only */
			marker?:{
				style?:"auto"|"none"|"circle"|"dash"|"diamond"|"dot"|"picture"|"plus"|"square"|"star"|"triangle"|"x";
				size?:number;
				fillColor?:string;
				borderColor?:string;
			};
			dataLabels?:{
				showLegendKey?:boolean;
				showValue?:boolean;
				showCategory?:boolean;
				showSeriesName?:boolean;
				showPercent?:boolean;
				showBubbleSize?:boolean;
				showLeaderLines?:boolean;
				position?:ChartDataLabelPosition; /** for bubble charts only center/left/right/top/bottom */
				numberFormat?:string;
				numberFormatLinked?:boolean;
				separator?:string;
				format?:ChartDataLabelFormat;
			};
		}[];
	}[];
	title?:ChartTextSource&{overlay?:boolean};
	legend?:{visible?:boolean;position?:"left"|"right"|"top"|"bottom"|"topRight";overlay?:boolean};
	axes?:{category?:ChartAxisSpec;value?:ChartAxisSpec;secondaryCategory?:ChartAxisSpec;secondaryValue?:ChartAxisSpec};
	format?:ChartDataLabelFormat; /** chart-area fill/border/font */
	plotArea?:ChartPlotAreaSpec;
	displayBlanksAs?:"gap"|"span"|"zero";
	plotVisibleOnly?:boolean;
	showDataLabelsOverMaximum?:boolean;
	roundedCorners?:boolean;
	styleId?:number; /** legacy styles 1-48, or modern catalog styles eg. 201,227,240,251,269,276. */
}
function listCharts(wb,options?:{ sheet?:string }):Promise<Array<{
	id?:number;
	sheet:string;
	name:string;
	type:string;
	groups:{type:string;axis?:string;seriesCount:number}[];
	groupCount:number;
	seriesCount:number;
	position:ChartPos;
}>>;
function getChart(wb,sheet:string,name:string):Promise<Omit<ChartSpec,"position">&{ position:ChartPos }>;
function addChart(wb,sheet:string,chart:ChartSpec):Promise<ChartSpec>;
function setChart(wb,sheet:string,name:string,chart:ChartSpec):Promise<ChartSpec>;
function deleteChart(wb,sheet:string,name:string):Promise<void>;
interface NameDef {name:string;range:string;scope:string|null}
function listDefinedNames(wb):Promise<NameDef[]>;
function addDefinedName(wb,name:string,range:string,scope?:string):Promise<NameDef>;
function deleteDefinedName(wb,name:string,scope?:string):Promise<NameDef>;
function addSheet(wb,name:string):Promise<string>;
function deleteSheet(wb,name:string):Promise<void>;
function renameSheet(wb,oldName:string,newName:string):Promise<void>;
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
function readCell(wb,cell:CellRef,opts?:{
	context?:number;
}):Promise<Value>;
function readRange(wb,range:RangeRef):Promise<Value[][]>;
function readColumn(wb,sheetName:string,col:number|string,opts?:{
	startRow?:number;
	endRow?:number;
}):Promise<Value[]>;
function readRow(wb,sheetName:string,row:number,opts?:{
	startCol?:number;
	endCol?:number;
}):Promise<Value[]>;
function readRangeTsv(wb,range:RangeRef,opts?:{
	includeEmpty?:boolean;
	includeFormulas?:boolean;
}):Promise<string>;
function readColumnTsv(wb,sheetName:string,col:number|string,opts?:{
	startRow?:number;
	endRow?:number;
	includeEmpty?:boolean;
	includeFormulas?:boolean;
}):Promise<string>;
function readRowTsv(wb,sheetName:string,row:number,opts?:{
	startCol?:number;
	endCol?:number;
	includeEmpty?:boolean;
	includeFormulas?:boolean;
}):Promise<string>;
declare class SearchResults<T> extends Array<T> {truncated?:boolean}
type Matcher=string|string[]|number|boolean|RegExp|RegExp[]
/** Cell search by scalar, string list, or regex; `formulas:true` matches formulas only. */
function findCells(wb,matcher:Matcher,opts?:{in?:RangeRef|string;context?:number;limit?:number;offset?:number;formulas?:boolean}):Promise<{
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
function findRows(wb,matcher:Matcher,opts?:{in?:RangeRef|string;context?:number;limit?:number;offset?:number}):Promise<{
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
function findAndReplace(wb,find:string|RegExp,replace:string,opts?:{
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
function describeSheet(wb,sheetName:string):Promise<{
	tables:Record<string,{address:string;headerRows:string;headerCols:string|null;tableName?:string}>;
	structure:string; // Compact ASCII structure map
}>;
function tableLookup(wb,args:{
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
function setCells(wb,cells:Array<{
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
function sweepInputs(wb,args:{
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
function scaleRange(wb,range:RangeRef,factor:number,opts?:{skipFormulas?:boolean}):Promise<WriteResult|null>;
function insertRowAfter(wb,sheetName:string,row:number,count?:number):Promise<void>;
function deleteRows(wb,sheetName:string,row:number,count?:number):Promise<void>;
function insertColumnAfter(wb,sheetName:string,column:number|string,count?:number):Promise<void>;
function deleteColumns(wb,sheetName:string,column:number|string,count?:number):Promise<void>;
function autoFitColumns(wb,sheetName:string,columns?:Array<number|string>,opts?:{
	minWidth?:number;
	maxWidth?:number;
	padding?:number;
}):Promise<Record<string,{width:number;previousWidth:number}>>;
function sortRange(wb,range:RangeRef,keys:Array<{
	col:number|string;
	order?:"asc"|"desc";
}>,opts?:{hasHeader?:boolean}):Promise<void>;
function copyRange(wb,source:RangeRef,destination:CellRef,opts?:{
	pasteType?:"all"|"values"|"formulas"|"formats";
}):Promise<{destination:string;cellsCopied:number;}>;
type StyleObj={
	fill?:{
		color?:string;
		pattern?:string;
		patternColor?:string;
		gradient?:{
			type:string;
			degree?:number;
			color1:string;
			color2:string;
			top?:number;
			bottom?:number;
			left?:number;
			right?:number;
		};
	};
	font?:{
		name?:string;
		size?:number;
		color?:string;
		bold?:boolean;
		italic?:boolean;
		strike?:boolean;
		underline?:string;
		verticalAlign?:string;
	};
	alignment?:{
		horizontal?:string;
		vertical?:string;
		rotation?:number;
		wrapText?:boolean;
		shrinkToFit?:boolean;
		indent?:number;
	};
	border?:{
		top?:{
			style:string;
			color:string;
		};
		bottom?:{
			style:string;
			color:string;
		};
		left?:{
			style:string;
			color:string;
		};
		right?:{
			style:string;
			color:string;
		};
		diagonal?:{
			style:string;
			color?:string;
			up?:boolean;
			down?:boolean;
		};
	};
	numberFormat?:string;
	centerContinuousSpan?:number;
	richText?:{
		text:string;
		style?:{
			name?:string;
			size?:number;
			color?:string;
			bold?:boolean;
			italic?:boolean;
			strike?:boolean;
			underline?:string;
			verticalAlign?:string;
		};
	}[];
};
function getStyle(wb,cell:CellRef):Promise<StyleObj>;
function setStyle(wb,target:CellRef|RangeRef,style:StyleObj):Promise<void>;
interface CfColorScalePoint {type:"formula"|"max"|"min"|"num"|"percent"|"percentile"|"autoMin"|"autoMax";value?:number;formula?:string;color?:string}
interface CfDataBarThreshold {type:"formula"|"max"|"min"|"num"|"percent"|"percentile"|"autoMin"|"autoMax";value?:number;formula?:string}
type CfWritableRuleType="cellValue"|"containsText"|"notContainsText"|"beginsWith"|"endsWith"|"containsBlanks"|"notContainsBlanks"|"containsErrors"|"notContainsErrors"|"expression"|"timePeriod"|"top"|"bottom"|"aboveAverage"|"belowAverage"|"duplicateValues"|"uniqueValues"|"twoColorScale"|"threeColorScale"|"dataBar"
interface CfRuleShared {
	address:string;
	priority?:number;
	stopIfTrue?:boolean;
	style?:StyleObj;
	operator?:"equal"|"notEqual"|"greaterThan"|"greaterThanOrEqual"|"lessThan"|"lessThanOrEqual"|"between"|"notBetween"|"above"|"aboveOrEqual"|"below"|"belowOrEqual";
	formula?:string;
	formula2?:string;
	text?:string;
	rank?:number;
	percent?:boolean;
	bottom?:boolean;
	stdDev?:number;
	lowValue?:CfColorScalePoint;
	midValue?:CfColorScalePoint;
	highValue?:CfColorScalePoint;
	dataBar?:{
		showValue?:boolean;
		gradient?:boolean;
		border?:boolean;
		negativeBarColorSameAsPositive?:boolean;
		negativeBarBorderColorSameAsPositive?:boolean;
		axisPosition?:"automatic"|"middle"|"none";
		direction?:"context"|"leftToRight"|"rightToLeft";
		fillColor?:string;
		borderColor?:string;
		negativeFillColor?:string;
		negativeBorderColor?:string;
		axisColor?:string;
		lowValue?:CfDataBarThreshold;
		highValue?:CfDataBarThreshold;
	};
	timePeriod?:"today"|"yesterday"|"tomorrow"|"last7Days"|"thisWeek"|"lastWeek"|"nextWeek"|"thisMonth"|"lastMonth"|"nextMonth";
}
function getConditionalFormatting(wb,sheetName:string):Promise<Array<CfRuleShared&{
	index?:number;
	type:CfWritableRuleType|"iconSet";
}>>;
/** Add conditional formatting rules; `opts.clear` replaces existing rules; `iconSet` is read-only. */
function setConditionalFormatting(wb,sheetName:string,rules:Array<CfRuleShared&{
	type:CfWritableRuleType;
}>,opts?:{clear?:boolean}):Promise<void>;
function removeConditionalFormatting(wb,sheetName:string,indices:number[]):Promise<void>;
type DvOperator="Between"|"NotBetween"|"EqualTo"|"NotEqualTo"|"GreaterThan"|"LessThan"|"GreaterThanOrEqualTo"|"LessThanOrEqualTo";
type DvBasic={operator:DvOperator;formula1:string|number;formula2?:string|number|null};
type DvRule={
	wholeNumber?:DvBasic|null;
	decimal?:DvBasic|null;
	list?:{source:string;inCellDropDown?:boolean}|null;
	date?:DvBasic|null;
	time?:DvBasic|null;
	textLength?:DvBasic|null;
	custom?:{formula:string}|null;
};
type DvSpec={
	address:string;
	rule:DvRule;
	ignoreBlanks?:boolean;
	prompt?:{showPrompt?:boolean;title?:string|null;message?:string|null}|null;
	errorAlert?:{showAlert?:boolean;style?:"Stop"|"Warning"|"Information";title?:string|null;message?:string|null}|null;
};
function getDataValidations(wb,opts?:{sheet?:string;address?:string}):Promise<Array<DvSpec&{
	index:number;
	sheet:string;
	type:"None"|"WholeNumber"|"Decimal"|"List"|"Date"|"Time"|"TextLength"|"Custom";
}>>;
function validateCells(wb,address:RangeRef,opts?:{
	maxCellsToScan?:number;
	maxInvalidCells?:number;
	treatUnsupportedAsInvalid?:boolean;
}):Promise<{
	status:"Valid"|"Invalid"|"NoValidation"|"Mixed"|"Unknown";
	invalidCells:string[];
	truncated:boolean;
	diagnostics:{code:string;message:string;details?:Record<string,string>|null}[];
}>;
function setDataValidations(wb,sheetName:string,rules:DvSpec[],opts?:{clear?:boolean}):Promise<void>;
function removeDataValidations(wb,sheetName:string,target:{indices:number[]}|{address:string}):Promise<void>;
function setSheetProperties(wb,sheetName:string,properties:{
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
function setRowProperties(wb,sheetName:string,fromRow:number,toRow:number,properties:{
	height?:number;
	hidden?:boolean;
	outlineLevel?:number;
	collapsed?:boolean;
}):Promise<void>;
function setColumnProperties(wb,sheetName:string,fromCol:number|string,toCol:number|string,properties:{
	width?:number;
	hidden?:boolean;
	outlineLevel?:number;
	collapsed?:boolean;
}):Promise<void>;
function getSheetProperties(wb,sheetName:string,filter?:{
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
interface Deps {cells:{address:string;depth:number;formula?:string;referenceType?:"direct"|"range"|"named"|"table"}[];warnings?:Diag[]}
function getCellPrecedents(wb,address:CellRef,depth?:number):Promise<Deps>;
function getCellDependents(wb,address:CellRef,depth?:number):Promise<Deps>;
function traceToInputs(wb,cell:CellRef):Promise<{
	address:string;
	referenceCount:number;
	text?:string;
	nearbyLabel?:string;
	context?:string;
}[]>;
function traceToOutputs(wb,cell:CellRef):Promise<{
	address:string;
	formula?:string;
	text?:string;
	visibility:Visibility;
	nearbyLabel?:string;
	context?:string;
}[]>;
interface FormulaEval {formula:string;value:number|string|boolean|null|unknown[][];error?:{code:string;detail?:string}}
function evaluateFormulas(wb,sheet:string,formulas:string[]):Promise<FormulaEval[]>;
function evaluateFormula(wb,sheet:string,formula:string):Promise<FormulaEval>;
function lint(wb,options?:{
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
function previewStyles(wb,range:RangeRef):Promise<void>;
