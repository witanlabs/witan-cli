// Witan PPTX chart extension reference.
//
// This file documents the chart APIs implemented by Witan PPTX on top of the
// Office.js-compatible PowerPoint runtime. These APIs are Excel-style chart
// APIs adapted to PPTX; they are not present in upstream Office.js PowerPoint
// declarations.

declare namespace PowerPoint {
  /** A 2D data table used to create or replace a chart's embedded workbook data. */
  type ChartValues = Array<Array<string | number | boolean | null>>;

  /** A local A1 reference inside the chart's embedded workbook, for example `Sheet1!A1:B4`. */
  type ChartLocalRange = string;

  /** Chart data orientation. Use `"columns"` when series are stored by columns and `"rows"` when series are stored by rows. */
  type ChartSeriesBy = "columns" | "rows";

  /** Chart blank-cell display behavior. */
  type ChartDisplayBlanksAs = "NotPlotted" | "Zero" | "Interplotted";

  /** Chart legend positions. `"Corner"` maps to the top-right legend placement. */
  type ChartLegendPosition = "Top" | "Bottom" | "Left" | "Right" | "Corner";

  /** Data label positions. Some positions are chart-type-specific and invalid on bubble charts. */
  type ChartDataLabelPosition =
    | "BestFit"
    | "Center"
    | "InsideBase"
    | "InsideEnd"
    | "OutsideEnd"
    | "Left"
    | "Right"
    | "Top"
    | "Bottom";

  /** Axis kind accepted by `ChartAxes.getItem`. */
  type ChartAxisType = "Category" | "Value" | "Invalid";

  /** Chart series or axis group. */
  type ChartAxisGroupValue = "Primary" | "Secondary";

  /** Chart axis category interpretation. */
  type ChartAxisCategoryTypeValue = "Automatic" | "DateAxis" | "TextAxis";

  /** Time unit used by date category axes. */
  type ChartAxisTimeUnitValue = "Days" | "Months" | "Years";

  /** Histogram/Pareto binning mode. */
  type ChartBinTypeValue = "Auto" | "BinCount" | "BinWidth" | "Category";

  /** Box and whisker quartile calculation mode. */
  type ChartQuartileCalculation = "Inclusive" | "Exclusive";

  /** Chart line and border style. */
  type ChartLineStyleValue =
    | "Automatic"
    | "Continuous"
    | "Dash"
    | "DashDot"
    | "DashDotDot"
    | "Dot"
    | "None"
    | "RoundDot";

  /** Marker style values returned by the runtime. */
  type ChartMarkerStyleValue =
    | "Automatic"
    | "Circle"
    | "Dash"
    | "Diamond"
    | "Dot"
    | "None"
    | "Picture"
    | "Plus"
    | "Square"
    | "Star"
    | "Triangle"
    | "X";

  /** Marker style enum input values accepted by setters. */
  type ChartMarkerStyleInput =
    | ChartMarkerStyleValue
    | "auto"
    | "circle"
    | "dash"
    | "diamond"
    | "dot"
    | "none"
    | "picture"
    | "plus"
    | "square"
    | "star"
    | "triangle"
    | "x";

  /** Supported chart types. Values match `PowerPoint.ChartType.*`. */
  type ChartTypeValue =
    | "Area"
    | "AreaStacked"
    | "AreaStacked100"
    | "BarClustered"
    | "BarStacked"
    | "BarStacked100"
    | "Bubble"
    | "Boxwhisker"
    | "ColumnClustered"
    | "ColumnStacked"
    | "ColumnStacked100"
    | "Doughnut"
    | "Funnel"
    | "Histogram"
    | "Line"
    | "LineMarkers"
    | "LineMarkersStacked"
    | "LineMarkersStacked100"
    | "LineStacked"
    | "LineStacked100"
    | "Pareto"
    | "Pie"
    | "Radar"
    | "RadarFilled"
    | "RadarMarkers"
    | "StockHLC"
    | "StockOHLC"
    | "StockVHLC"
    | "StockVOHLC"
    | "SurfaceTopView"
    | "SurfaceTopViewWireframe"
    | "Waterfall"
    | "XYScatter"
    | "XYScatterLines"
    | "XYScatterLinesNoMarkers"
    | "XYScatterSmooth"
    | "XYScatterSmoothNoMarkers";

  /** Witan PPTX chart type enum object. */
  const ChartType: {
    area: "Area";
    areaStacked: "AreaStacked";
    areaStacked100: "AreaStacked100";
    barClustered: "BarClustered";
    barStacked: "BarStacked";
    barStacked100: "BarStacked100";
    bubble: "Bubble";
    boxwhisker: "Boxwhisker";
    columnClustered: "ColumnClustered";
    columnStacked: "ColumnStacked";
    columnStacked100: "ColumnStacked100";
    doughnut: "Doughnut";
    funnel: "Funnel";
    histogram: "Histogram";
    line: "Line";
    lineMarkers: "LineMarkers";
    lineMarkersStacked: "LineMarkersStacked";
    lineMarkersStacked100: "LineMarkersStacked100";
    lineStacked: "LineStacked";
    lineStacked100: "LineStacked100";
    pareto: "Pareto";
    pie: "Pie";
    radar: "Radar";
    radarFilled: "RadarFilled";
    radarMarkers: "RadarMarkers";
    stockHLC: "StockHLC";
    stockOHLC: "StockOHLC";
    stockVHLC: "StockVHLC";
    stockVOHLC: "StockVOHLC";
    surfaceTopView: "SurfaceTopView";
    surfaceTopViewWireframe: "SurfaceTopViewWireframe";
    waterfall: "Waterfall";
    xyScatter: "XYScatter";
    xyScatterLines: "XYScatterLines";
    xyScatterLinesNoMarkers: "XYScatterLinesNoMarkers";
    xyScatterSmooth: "XYScatterSmooth";
    xyScatterSmoothNoMarkers: "XYScatterSmoothNoMarkers";
  };

  /** Witan PPTX chart axis category enum object. */
  const ChartAxisCategoryType: {
    automatic: "Automatic";
    dateAxis: "DateAxis";
    textAxis: "TextAxis";
  };

  /** Witan PPTX chart axis group enum object. */
  const ChartAxisGroup: {
    primary: "Primary";
    secondary: "Secondary";
  };

  /** Witan PPTX chart axis time unit enum object. */
  const ChartAxisTimeUnit: {
    days: "Days";
    months: "Months";
    years: "Years";
  };

  /** Witan PPTX histogram/Pareto bin type enum object. */
  const ChartBinType: {
    auto: "Auto";
    binCount: "BinCount";
    binWidth: "BinWidth";
    category: "Category";
  };

  /** Witan PPTX chart line style enum object. */
  const ChartLineStyle: {
    automatic: "Automatic";
    continuous: "Continuous";
    dash: "Dash";
    dashDot: "DashDot";
    dashDotDot: "DashDotDot";
    dot: "Dot";
    none: "None";
    roundDot: "RoundDot";
  };

  /** Witan PPTX chart marker style enum object. */
  const ChartMarkerStyle: {
    automatic: "Automatic";
    circle: "Circle";
    dash: "Dash";
    diamond: "Diamond";
    dot: "Dot";
    none: "None";
    picture: "Picture";
    plus: "Plus";
    square: "Square";
    star: "Star";
    triangle: "Triangle";
    x: "X";
  };

  /** Witan PPTX chart underline style enum object. */
  const ChartUnderlineStyle: {
    none: "None";
    single: "Single";
  };

  /** Options for `ShapeCollection.addChart`. */
  interface AddChartOptions {
    /** Left position in points. Default: 0. */
    left?: number;
    /** Top position in points. Default: 0. */
    top?: number;
    /** Width in points. Default: 432. */
    width?: number;
    /** Height in points. Default: 252. */
    height?: number;
    /** Shape and chart name. If omitted, Witan generates `Chart 1`, `Chart 2`, etc. */
    name?: string;
    /** Whether series are read by columns or rows. */
    seriesBy?: ChartSeriesBy;
    /** Embedded workbook worksheet name. Default: `Sheet1`. */
    sheetName?: string;
    /** Top-left embedded workbook cell used for generated data. Default: `A1`. */
    topLeftCell?: string;
    /** Chart style id. Use a PowerPoint/Excel chart style number such as `201` or `227`. */
    styleId?: number;
  }

  /** Options for `Chart.setData`. */
  interface ChartSetDataOptions {
    /** Optional replacement chart type. Defaults to the chart's current type. */
    chartType?: ChartTypeValue;
    /** Whether series are read by columns or rows. */
    seriesBy?: ChartSeriesBy;
    /** Embedded workbook worksheet name. Default: `Sheet1`. */
    sheetName?: string;
    /** Top-left embedded workbook cell used for generated data. Default: `A1`. */
    topLeftCell?: string;
    /** Optional replacement chart style id. */
    styleId?: number;
  }

  /** Options for `Chart.setDataRange`. */
  interface ChartSetDataRangeOptions {
    /** Whether series are read by columns or rows. */
    seriesBy?: ChartSeriesBy;
  }

  interface ShapeCollection {
    /**
     * Creates a chart shape backed by an embedded workbook.
     *
     * The first row/column supplies labels depending on `seriesBy`. The returned
     * shape is a normal `PowerPoint.Shape` whose `type` is `"Chart"` and whose
     * chart can be accessed with `shape.getChart()`.
     */
    addChart(chartType: ChartTypeValue, values: ChartValues, options?: AddChartOptions): Shape;
  }

  interface Shape {
    /**
     * Returns the chart hosted by this shape.
     *
     * Throws an OfficeExtension invalid-argument error when the shape is not a
     * chart shape.
     */
    getChart(): Chart;

    /**
     * Returns the chart hosted by this shape, or a null object when the shape is
     * not a chart shape. Check `.isNullObject`.
     */
    getChartOrNullObject(): Chart;
  }

  /** Embedded-workbook-backed chart hosted inside a PPTX shape. */
  interface Chart extends OfficeExtension.ClientObject {
    /** Axis accessors for category/value and primary/secondary axes. */
    readonly axes: ChartAxes;
    /** Chart type. Can be changed to convert the chart. */
    chartType: ChartTypeValue;
    /** Chart-wide data labels for the first chart group. */
    readonly dataLabels: ChartDataLabels;
    /** Data source kind. Witan-authored charts currently report `EmbeddedWorkbook`. */
    readonly dataSourceKind: string;
    /** Blank-cell display behavior. ChartEx charts reject writes to this property. */
    displayBlanksAs: ChartDisplayBlanksAs;
    /** Chart area format. */
    readonly format: ChartAreaFormat;
    /** Chart frame height in points. */
    height: number;
    /** Shape id for the chart frame. */
    readonly id: string;
    /** Chart frame left position in points. */
    left: number;
    /** Legend formatting and visibility. */
    readonly legend: ChartLegend;
    /** Chart and shape name. */
    name: string;
    /** Plot area formatting. */
    readonly plotArea: ChartPlotArea;
    /** Whether hidden source rows/columns are excluded. ChartEx charts reject writes to this property. */
    plotVisibleOnly: boolean;
    /** Series collection. */
    readonly series: ChartSeriesCollection;
    /** Whether data labels may be shown over the maximum value. ChartEx charts reject writes to this property. */
    showDataLabelsOverMaximum: boolean;
    /** Chart style id. */
    style: number;
    /** Chart title. */
    readonly title: ChartTitle;
    /** Chart frame top position in points. */
    top: number;
    /** Chart frame width in points. */
    width: number;

    /** Deletes the chart shape from its parent shape collection. */
    delete(): void;

    /** Replaces embedded workbook data, optionally changing chart type/style. */
    setData(values: ChartValues, options?: ChartSetDataOptions): void;

    /** Repoints the chart to a local range in the embedded workbook without replacing the embedded workbook package. */
    setDataRange(rangeAddress: ChartLocalRange, options?: ChartSetDataRangeOptions): void;
  }

  /** Chart axes collection. */
  interface ChartAxes extends OfficeExtension.ClientObject {
    /** Primary category axis. */
    readonly categoryAxis: ChartAxis;
    /** Primary value axis. */
    readonly valueAxis: ChartAxis;

    /**
     * Gets an axis by kind and optional group.
     *
     * `type` accepts `"Category"`, `"Value"`, or `"Invalid"`; `"Invalid"` maps
     * to the value axis for Excel compatibility. `axisGroup` defaults to
     * `"Primary"`.
     */
    getItem(type: ChartAxisType, axisGroup?: ChartAxisGroupValue): ChartAxis;
  }

  /** A chart axis. */
  interface ChartAxis extends OfficeExtension.ClientObject {
    /** Axis group. */
    readonly axisGroup: ChartAxisGroupValue;
    /** Date-axis base time unit. Null unless this is a date category axis. */
    baseTimeUnit: ChartAxisTimeUnitValue | null;
    /** Category axis mode. Only valid on category axes. */
    categoryType: ChartAxisCategoryTypeValue;
    /** Whether number format is linked to source data. */
    linkNumberFormat: boolean;
    /** Major gridline settings. */
    readonly majorGridlines: ChartGridlines;
    /** Date-axis major time unit. Null unless this is a date category axis. */
    majorTimeUnitScale: ChartAxisTimeUnitValue | null;
    /** Major numeric unit. Set `""` or `null` for automatic. */
    majorUnit: number | null | "";
    /** Maximum axis value. Set `""` or `null` for automatic. */
    maximum: number | null | "";
    /** Minor gridline settings. */
    readonly minorGridlines: ChartGridlines;
    /** Date-axis minor time unit. Null unless this is a date category axis. */
    minorTimeUnitScale: ChartAxisTimeUnitValue | null;
    /** Minimum axis value. Set `""` or `null` for automatic. */
    minimum: number | null | "";
    /** Minor numeric unit. Set `""` or `null` for automatic. */
    minorUnit: number | null | "";
    /** Axis number format. Setting this also sets `linkNumberFormat` to false. */
    numberFormat: string;
    /** Whether plot order is reversed. */
    reversePlotOrder: boolean;
    /** Axis title. */
    readonly title: ChartAxisTitle;
    /** Axis type. */
    readonly type: "Category" | "Value";
    /** Axis visibility. */
    visible: boolean;
  }

  /** Major or minor gridline visibility. */
  interface ChartGridlines extends OfficeExtension.ClientObject {
    /** Whether these gridlines are visible. */
    visible: boolean;
  }

  /** Axis title. */
  interface ChartAxisTitle extends OfficeExtension.ClientObject {
    /** Axis title text. Null when no title exists. Setting text creates the title. */
    text: string | null;
    /** Whether the axis title exists. Setting false removes it. */
    visible: boolean;
  }

  /** Chart title. */
  interface ChartTitle extends OfficeExtension.ClientObject {
    /** Whether the title overlays the plot area. */
    overlay: boolean;
    /** Title text. Null when the title is hidden or unset. Setting text makes the title visible. */
    text: string | null;
    /** Title visibility. */
    visible: boolean;
  }

  /** Chart legend. */
  interface ChartLegend extends OfficeExtension.ClientObject {
    /** Whether the legend overlays the plot area. */
    overlay: boolean;
    /** Legend position. Null when the legend is hidden. */
    position: ChartLegendPosition | null;
    /** Legend visibility. */
    visible: boolean;
  }

  /** Chart series collection. */
  interface ChartSeriesCollection extends OfficeExtension.ClientObject {
    /** Series count as an immediate property. */
    readonly count: number;
    /** Materialized series list. */
    readonly items: ChartSeries[];
    /** Series count as an Office.js-style ClientResult. Read `.value` after `context.sync()`. */
    getCount(): OfficeExtension.ClientResult<number>;
    /** Gets a series by zero-based index. */
    getItemAt(index: number): ChartSeries;
  }

  /** A chart series. */
  interface ChartSeries extends OfficeExtension.ClientObject {
    /** Primary or secondary axis group. Changing this can create combo chart groups. */
    axisGroup: ChartAxisGroupValue;
    /** Bubble scale, 0-300. Only valid on bubble charts. */
    bubbleScale: number;
    /** Series-level data labels. */
    readonly dataLabels: ChartDataLabels;
    /** Doughnut hole size, 10-90. Only valid on doughnut charts. */
    doughnutHoleSize: number;
    /** First slice angle, 0-360. Only valid on pie and doughnut charts. */
    firstSliceAngle: number;
    /** Series format. */
    readonly format: ChartSeriesFormat;
    /** Gap width, 0-500. Only valid on bar, column, waterfall, and box-and-whisker charts. */
    gapWidth: number;
    /** Per-series chart type. Changing this can create combo chart groups. */
    chartType: ChartTypeValue;
    /** Whether this series has displayed data labels. */
    hasDataLabels: boolean;
    /** Whether negative values invert fill. Only valid on bar and column charts. */
    invertIfNegative: boolean;
    /** Marker fill color. Null when markers are unsupported or unset. */
    markerBackgroundColor: string | null;
    /** Marker border color. Null when markers are unsupported or unset. */
    markerForegroundColor: string | null;
    /** Marker size, rounded to an integer, 2-72. */
    markerSize: number;
    /** Marker style. Marker setters are valid for line charts and scatter charts with markers. */
    markerStyle: ChartMarkerStyleValue;
    /** Series name. */
    name: string;
    /** Bar/column overlap, -100 to 100. Only valid on bar and column charts. */
    overlap: number;
    /** Whether connector lines are visible. Only valid on waterfall charts. */
    showConnectorLines: boolean;
    /** Whether line/scatter curves are smoothed. */
    smooth: boolean;
    /** Whether categories vary by color. */
    varyByCategories: boolean;

    /** Deletes the series from the chart. */
    delete(): void;
    /** Gets histogram/Pareto binning options. Only valid on histogram and Pareto charts. */
    getBinOptions(): ChartBinOptions;
    /** Gets box-and-whisker options. Only valid on box-and-whisker charts. */
    getBoxwhiskerOptions(): ChartBoxwhiskerOptions;
    /** Returns the local embedded-workbook source range for a dimension. */
    getDimensionDataSourceString(dimension: "categories" | "category" | "values" | "value" | "xValues" | "xValue" | "xAxisValues" | "yValues" | "yValue" | "yAxisValues" | "bubbleSizes" | "bubbleSize"): string;
    /** Returns `"LocalRange"` when the dimension has a local embedded-workbook source. */
    getDimensionDataSourceType(dimension: string): "LocalRange";
    /** Sets bubble size source range. Only valid on bubble charts. */
    setBubbleSizes(sourceData: ChartLocalRange): void;
    /** Sets value/Y source range. For scatter and bubble charts this sets Y values. */
    setValues(sourceData: ChartLocalRange): void;
    /** Sets category/X source range. For scatter and bubble charts this sets X values. */
    setXAxisValues(sourceData: ChartLocalRange): void;
  }

  /** Histogram and Pareto binning options. */
  interface ChartBinOptions extends OfficeExtension.ClientObject {
    /** Returns whether an overflow bin is enabled. */
    getAllowOverflow(): boolean;
    /** Enables or disables an overflow bin. Enabling creates a default overflow value when absent. */
    setAllowOverflow(value: boolean): void;
    /** Returns whether an underflow bin is enabled. */
    getAllowUnderflow(): boolean;
    /** Enables or disables an underflow bin. Enabling creates a default underflow value when absent. */
    setAllowUnderflow(value: boolean): void;
    /** Returns the number of bins, or 0 when not set. */
    getCount(): number;
    /** Sets the number of bins and changes the bin type to `"BinCount"`. */
    setCount(value: number): void;
    /** Returns the overflow value, or 0 when not set. */
    getOverflowValue(): number;
    /** Sets the overflow value and enables overflow bins. */
    setOverflowValue(value: number): void;
    /** Returns the binning mode. */
    getType(): ChartBinTypeValue;
    /** Sets the binning mode. `"Category"` is only valid when the series has categories. */
    setType(value: ChartBinTypeValue): void;
    /** Returns the underflow value, or 0 when not set. */
    getUnderflowValue(): number;
    /** Sets the underflow value and enables underflow bins. */
    setUnderflowValue(value: number): void;
    /** Returns the bin width, or 0 when not set. */
    getWidth(): number;
    /** Sets the bin width and changes the bin type to `"BinWidth"`. */
    setWidth(value: number): void;
  }

  /** Box-and-whisker chart display options. */
  interface ChartBoxwhiskerOptions extends OfficeExtension.ClientObject {
    /** Returns the quartile calculation mode. */
    getQuartileCalculation(): ChartQuartileCalculation;
    /** Sets the quartile calculation mode. */
    setQuartileCalculation(quartileCalculation: ChartQuartileCalculation): void;
    /** Returns whether inner/non-outlier points are shown. */
    getShowInnerPoints(): boolean;
    /** Sets whether inner/non-outlier points are shown. */
    setShowInnerPoints(value: boolean): void;
    /** Returns whether the mean line is shown. */
    getShowMeanLine(): boolean;
    /** Sets whether the mean line is shown. */
    setShowMeanLine(value: boolean): void;
    /** Returns whether the mean marker is shown. */
    getShowMeanMarker(): boolean;
    /** Sets whether the mean marker is shown. */
    setShowMeanMarker(value: boolean): void;
    /** Returns whether outlier points are shown. */
    getShowOutlierPoints(): boolean;
    /** Sets whether outlier points are shown. */
    setShowOutlierPoints(value: boolean): void;
  }

  /** Chart-wide or series-level data labels. */
  interface ChartDataLabels extends OfficeExtension.ClientObject {
    /** Label format. */
    readonly format: ChartDataLabelFormat;
    /** Whether number format is linked to source data. */
    linkNumberFormat: boolean;
    /** Label number format. Setting this also sets `linkNumberFormat` to false. */
    numberFormat: string;
    /** Label position. Some positions are invalid for some chart types. */
    position: ChartDataLabelPosition | null;
    /** Separator between displayed label fields. */
    separator: string;
    /** Show bubble size. Only valid on bubble charts when true. */
    showBubbleSize: boolean;
    /** Show category name. */
    showCategoryName: boolean;
    /** Show leader lines. */
    showLeaderLines: boolean;
    /** Show legend key. */
    showLegendKey: boolean;
    /** Show percentage. Only valid on pie/doughnut charts when true. */
    showPercentage: boolean;
    /** Show series name. */
    showSeriesName: boolean;
    /** Show value. */
    showValue: boolean;
  }

  /** Chart area formatting. */
  interface ChartAreaFormat extends OfficeExtension.ClientObject {
    readonly border: ChartBorder;
    readonly fill: ChartFill;
    readonly font: ChartFont;
    /** Whether the chart area has rounded corners. */
    roundedCorners: boolean;
  }

  /** Plot area wrapper. */
  interface ChartPlotArea extends OfficeExtension.ClientObject {
    readonly format: ChartPlotAreaFormat;
  }

  /** Plot area formatting. */
  interface ChartPlotAreaFormat extends OfficeExtension.ClientObject {
    readonly border: ChartBorder;
    readonly fill: ChartFill;
  }

  /** Series formatting. */
  interface ChartSeriesFormat extends OfficeExtension.ClientObject {
    readonly fill: ChartFill;
    readonly line: ChartLineFormat;
  }

  /** Data label formatting. */
  interface ChartDataLabelFormat extends OfficeExtension.ClientObject {
    readonly border: ChartBorder;
    readonly fill: ChartFill;
    readonly font: ChartFont;
  }

  /** Solid chart fill formatting. */
  interface ChartFill extends OfficeExtension.ClientObject {
    /** Clears the fill. */
    clear(): void;
    /** Returns the solid fill color as a ClientResult. Read `.value` after `context.sync()`. */
    getSolidColor(): OfficeExtension.ClientResult<string>;
    /** Sets a solid fill color. Named colors and hex colors are normalized by Witan. */
    setSolidColor(color: string): void;
  }

  /** Chart line formatting. */
  interface ChartLineFormat extends OfficeExtension.ClientObject {
    /** Line color, or empty string when not visible. */
    color: string;
    /** Line style. */
    lineStyle: ChartLineStyleValue;
    /** Line weight in points, or null when unset/not visible. */
    weight: number | null;
    /** Clears the line. */
    clear(): void;
  }

  /** Chart border formatting. */
  interface ChartBorder extends OfficeExtension.ClientObject {
    /** Border color, or empty string when not visible. */
    color: string;
    /** Border line style. */
    lineStyle: ChartLineStyleValue;
    /** Border weight in points, or null when unset/not visible. */
    weight: number | null;
    /** Clears the border. */
    clear(): void;
  }

  /** Chart font formatting. */
  interface ChartFont extends OfficeExtension.ClientObject {
    bold: boolean;
    /** Font color, or empty string when unset. */
    color: string;
    italic: boolean;
    /** Font family name, or empty string when unset. */
    name: string;
    /** Font size in points, or null when unset. */
    size: number | null;
    /** Underline style. */
    underline: "None" | "Single";
  }
}
