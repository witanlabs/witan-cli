// Chart APIs
// Witan xlsx-code-mode exec API reference. All functions are async and take `wb` first.

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
		type:"column"|"bar"|"line"|"area"|"pie"|"doughnut"|"scatter"|"bubble"|"radar"|"surface"|"stockHLC"|"stockOHLC"|"waterfall"|"histogram"|"pareto"|"funnel";
		scatterStyle?:"line"|"lineMarker"|"marker"|"smooth"|"smoothMarker"; /** scatter only */
		radarStyle?:"standard"|"marker"|"filled"; /** radar only */
		surfaceVariant?:"topView"|"topViewWireframe"; /** surface only */
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
			binOptions?:{
				type?:"auto"|"binCount"|"binWidth"|"category"; /** histogram/Pareto only */
				count?:number; /** required when type is "binCount" */
				width?:number; /** required when type is "binWidth" */
				allowOverflow?:boolean;
				overflowValue?:number;
				allowUnderflow?:boolean;
				underflowValue?:number;
			};
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
declare function listCharts(wb,options?:{ sheet?:string }):Promise<Array<{
	id?:number;
	sheet:string;
	name:string;
	type:string;
	groups:{type:string;axis?:string;seriesCount:number}[];
	groupCount:number;
	seriesCount:number;
	position:ChartPos;
}>>;
declare function getChart(wb,sheet:string,name:string):Promise<Omit<ChartSpec,"position">&{ position:ChartPos }>;
declare function addChart(wb,sheet:string,chart:ChartSpec):Promise<ChartSpec>;
declare function setChart(wb,sheet:string,name:string,chart:ChartSpec):Promise<ChartSpec>;
declare function deleteChart(wb,sheet:string,name:string):Promise<void>;
