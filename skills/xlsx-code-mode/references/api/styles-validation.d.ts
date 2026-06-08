// Style, conditional formatting, and data validation APIs
// Witan xlsx-code-mode exec API reference. All functions are async and take `wb` first.

type CellRef=string|{sheet:string;row:number;col:number|string}
type RangeRef=string|{sheet:string}|{sheet:string;from:{row?:number;col?:number|string};to:{row?:number;col?:number|string}}

type FontStyle={
	name?:string;
	size?:number;
	color?:string;
	bold?:boolean;
	italic?:boolean;
	strike?:boolean;
	underline?:string;
	verticalAlign?:string;
}
type BorderSide={style:string;color:string}
type DiagonalBorder={style:string;color?:string;up?:boolean;down?:boolean}
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
	font?:FontStyle;
	alignment?:{
		horizontal?:string;
		vertical?:string;
		rotation?:number;
		wrapText?:boolean;
		shrinkToFit?:boolean;
		indent?:number;
	};
	border?:{
		top?:BorderSide;
		bottom?:BorderSide;
		left?:BorderSide;
		right?:BorderSide;
		diagonal?:DiagonalBorder;
	};
	numberFormat?:string;
	centerContinuousSpan?:number;
	richText?:{
		text:string;
		style?:FontStyle;
	}[];
};
declare function getStyle(wb,cell:CellRef):Promise<StyleObj>;
declare function setStyle(wb,target:CellRef|RangeRef,style:StyleObj):Promise<void>;
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
declare function getConditionalFormatting(wb,sheetName:string):Promise<Array<CfRuleShared&{
	index?:number;
	type:CfWritableRuleType|"iconSet";
}>>;
/** Add conditional formatting rules; `opts.clear` replaces existing rules; `iconSet` is read-only. */
declare function setConditionalFormatting(wb,sheetName:string,rules:Array<CfRuleShared&{
	type:CfWritableRuleType;
}>,opts?:{clear?:boolean}):Promise<void>;
declare function removeConditionalFormatting(wb,sheetName:string,indices:number[]):Promise<void>;
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
declare function getDataValidations(wb,opts?:{sheet?:string;address?:string}):Promise<Array<DvSpec&{
	index:number;
	sheet:string;
	type:"None"|"WholeNumber"|"Decimal"|"List"|"Date"|"Time"|"TextLength"|"Custom";
}>>;
declare function validateCells(wb,address:RangeRef,opts?:{
	maxCellsToScan?:number;
	maxInvalidCells?:number;
	treatUnsupportedAsInvalid?:boolean;
}):Promise<{
	status:"Valid"|"Invalid"|"NoValidation"|"Mixed"|"Unknown";
	invalidCells:string[];
	truncated:boolean;
	diagnostics:{code:string;message:string;details?:Record<string,string>|null}[];
}>;
declare function setDataValidations(wb,sheetName:string,rules:DvSpec[],opts?:{clear?:boolean}):Promise<void>;
declare function removeDataValidations(wb,sheetName:string,target:{indices:number[]}|{address:string}):Promise<void>;
