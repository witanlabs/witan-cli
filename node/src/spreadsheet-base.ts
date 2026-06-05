import { WitanProcessError } from './errors.js';
import { dropUndefined, serializeMatcher } from './helpers.js';
import type {
  AutoFitColumnResult,
  AutoFitRowResult,
  CellAssignment,
  ChartInfo,
  ChartSpec,
  ChartSummary,
  ConditionalFormattingRule,
  CopyRangeResult,
  DataTable,
  DataTableMutationResult,
  DataTableSpec,
  DefinedName,
  DependencyResult,
  FindAndReplaceResult,
  FormulaResult,
  LintResult,
  ListObject,
  ListObjectMutationResult,
  ListObjectSpec,
  ListObjectUpdate,
  PasteType,
  SearchCell,
  SearchRow,
  SetCellsValidationMode,
  SheetInfo,
  SheetProperties,
  SheetPropertiesUpdate,
  RowProperties,
  ColumnProperties,
  Style,
  SheetDescription,
  SweepInput,
  SweepMode,
  SweepResult,
  TableLookupResult,
  TraceInput,
  TraceOutput,
  Value,
  WorkbookProperties,
  WorkbookPropertiesUpdate,
  WriteResult,
} from './generated-types.js';

// SDK-specific utility types (not part of RPC wire format)
export type Matcher = string | number | boolean | RegExp | (string | RegExp)[];
export type ReplaceMatcher = string | RegExp;

export interface FindCellsOptions {
  /** Range to search within (e.g., "Sheet1!A:Z") */
  in?: string;
  /** Number of context rows/cols (default: 2) */
  context?: number;
  /** Maximum results (default: 20) */
  limit?: number;
  /** Skip first N results (default: 0) */
  offset?: number;
  /** Search in formulas instead of values */
  formulas?: boolean;
}

export interface FindRowsOptions {
  /** Range to search within */
  in?: string;
  /** Number of context rows */
  context?: number;
  /** Maximum results (default: 20) */
  limit?: number;
  /** Skip first N results (default: 0) */
  offset?: number;
}

export interface FindAndReplaceOptions {
  /** Range to search within */
  in?: string;
  /** Case-sensitive matching */
  matchCase?: boolean;
  /** Match entire cell content only */
  wholeCell?: boolean;
  /** Replace in formulas */
  inFormulas?: boolean;
  /** Maximum replacements */
  limit?: number;
}

export interface SweepOptions {
  /** Sweep mode: "cartesian" (all combinations) or "parallel" (zip) */
  mode?: SweepMode;
  /** Include min/max/mean statistics */
  includeStats?: boolean;
}

export interface LintOptions {
  /** Specific ranges to lint */
  rangeAddresses?: string[];
  /** Rule IDs to skip */
  skipRuleIds?: string[];
  /** Only run these rule IDs */
  onlyRuleIds?: string[];
}

export interface AutoFitColumnsOptions {
  /** Specific columns to fit (default: all) */
  columns?: (number | string)[];
  /** Minimum width */
  minWidth?: number;
  /** Maximum width */
  maxWidth?: number;
  /** Extra padding */
  padding?: number;
}

export interface AutoFitRowsOptions {
  /** Specific rows to fit (default: all) */
  rows?: number[];
  /** Minimum height */
  minHeight?: number;
  /** Maximum height */
  maxHeight?: number;
}

export interface SortKey {
  /** Column index or letter */
  column: number | string;
  /** Sort direction */
  descending?: boolean;
}

/**
 * Shared spreadsheet RPC operations for XLSX workbooks and Google Sheets sessions.
 */
export abstract class SpreadsheetSessionBase {
  protected abstract request(
    method: string,
    op: string,
    args?: Record<string, unknown>
  ): Promise<unknown>;

  /**
   * Get workbook-level properties.
   */
  async getWorkbookProperties(): Promise<WorkbookProperties> {
    return (await this.request('getWorkbookProperties', 'getWorkbookProperties', {})) as WorkbookProperties;
  }

  /**
   * Set workbook-level properties.
   */
  async setWorkbookProperties(properties: WorkbookPropertiesUpdate): Promise<void> {
    await this.request('setWorkbookProperties', 'setWorkbookProperties', properties as Record<string, unknown>);
  }

  /**
   * List all sheets in the workbook.
   */
  async listSheets(): Promise<SheetInfo[]> {
    const result = (await this.request('listSheets', 'listSheets', {})) as { sheets?: SheetInfo[] };
    return result.sheets ?? [];
  }

  /**
   * Add a new sheet to the workbook.
   * @param name - Name for the new sheet
   * @returns The name of the created sheet
   */
  async addSheet(name: string): Promise<string> {
    await this.request('addSheet', 'addSheet', { name });
    return name;
  }

  /**
   * Delete a sheet from the workbook.
   * @param name - Name of the sheet to delete
   */
  async deleteSheet(name: string): Promise<void> {
    await this.request('deleteSheet', 'deleteSheet', { name });
  }

  /**
   * Rename a sheet.
   * @param oldName - Current name of the sheet
   * @param newName - New name for the sheet
   */
  async renameSheet(oldName: string, newName: string): Promise<void> {
    await this.request('renameSheet', 'renameSheet', { oldName, newName });
  }

  /**
   * Get properties of a sheet.
   * @param sheet - Sheet name
   * @param options - Filter options for columns/rows
   */
  async getSheetProperties(
    sheet: string,
    options: { columns?: (number | string)[]; rows?: number[] } = {}
  ): Promise<SheetProperties> {
    const filter = dropUndefined({
      columns: options.columns,
      rows: options.rows,
    });

    const args: Record<string, unknown> = { sheet };
    if (Object.keys(filter).length > 0) {
      args['filter'] = filter;
    }

    const result = (await this.request('getSheetProperties', 'getSheetProperties', args)) as SheetProperties;

    result.columns ??= {};
    result.rows ??= {};

    return result;
  }

  /**
   * Set properties of a sheet.
   * @param sheet - Sheet name
   * @param properties - Properties to set
   */
  async setSheetProperties(sheet: string, properties: SheetPropertiesUpdate): Promise<void> {
    await this.request('setSheetProperties', 'setSheetProperties', {
      sheet,
      properties,
    });
  }

  /**
   * Read a single cell value.
   * This is a composite operation that calls readRange and extracts the first cell.
   *
   * @param cell - Cell address (e.g., "Sheet1!A1" or "A1")
   * @param options - Read options
   * @returns The cell value
   */
  async readCell(cell: string, options: { context?: number } = {}): Promise<Value> {
    const data = (await this.request(
      'readCell',
      'readRange',
      dropUndefined({ address: cell, context: options.context })
    )) as Value[][];
    return data[0]![0]!;
  }

  /**
   * Read a range of cells.
   *
   * @param range - Range address (e.g., "Sheet1!A1:B10")
   * @returns 2D array of cell values
   */
  async readRange(range: string): Promise<Value[][]> {
    return (await this.request('readRange', 'readRange', { address: range })) as Value[][];
  }

  /**
   * Read a row of cells.
   *
   * @param sheet - Sheet name
   * @param row - Row number (1-based)
   * @param options - Read options
   * @returns Array of cell values
   */
  async readRow(
    sheet: string,
    row: number,
    options: { startCol?: number; endCol?: number } = {}
  ): Promise<Value[]> {
    return (await this.request(
      'readRow',
      'readRow',
      dropUndefined({
        sheet,
        row,
        startCol: options.startCol,
        endCol: options.endCol,
      })
    )) as Value[];
  }

  /**
   * Read a column of cells.
   *
   * @param sheet - Sheet name
   * @param col - Column number (1-based) or letter (e.g., "A")
   * @param options - Read options
   * @returns Array of cell values
   */
  async readColumn(
    sheet: string,
    col: number | string,
    options: { startRow?: number; endRow?: number } = {}
  ): Promise<Value[]> {
    return (await this.request(
      'readColumn',
      'readColumn',
      dropUndefined({
        sheet,
        col,
        startRow: options.startRow,
        endRow: options.endRow,
      })
    )) as Value[];
  }

  /**
   * Read a range as tab-separated values.
   *
   * @param range - Range address
   * @param options - Read options
   * @returns TSV string
   */
  async readRangeTsv(
    range: string,
    options: { includeEmpty?: boolean; includeFormulas?: boolean } = {}
  ): Promise<string> {
    return (await this.request(
      'readRangeTsv',
      'readRangeTsv',
      dropUndefined({
        address: range,
        includeEmpty: options.includeEmpty,
        includeFormulas: options.includeFormulas,
      })
    )) as string;
  }

  /**
   * Read a row as tab-separated values.
   *
   * @param sheet - Sheet name
   * @param row - Row number (1-based)
   * @param options - Read options
   * @returns TSV string
   */
  async readRowTsv(
    sheet: string,
    row: number,
    options: {
      startCol?: number;
      endCol?: number;
      includeEmpty?: boolean;
      includeFormulas?: boolean;
    } = {}
  ): Promise<string> {
    return (await this.request(
      'readRowTsv',
      'readRowTsv',
      dropUndefined({
        sheet,
        row,
        startCol: options.startCol,
        endCol: options.endCol,
        includeEmpty: options.includeEmpty,
        includeFormulas: options.includeFormulas,
      })
    )) as string;
  }

  /**
   * Read a column as tab-separated values.
   *
   * @param sheet - Sheet name
   * @param col - Column number (1-based) or letter
   * @param options - Read options
   * @returns TSV string
   */
  async readColumnTsv(
    sheet: string,
    col: number | string,
    options: {
      startRow?: number;
      endRow?: number;
      includeEmpty?: boolean;
      includeFormulas?: boolean;
    } = {}
  ): Promise<string> {
    return (await this.request(
      'readColumnTsv',
      'readColumnTsv',
      dropUndefined({
        sheet,
        col,
        startRow: options.startRow,
        endRow: options.endRow,
        includeEmpty: options.includeEmpty,
        includeFormulas: options.includeFormulas,
      })
    )) as string;
  }

  /**
   * Set values in multiple cells.
   *
   * @param cells - Array of cell assignments
   * @returns Write result with changed cells and any errors
   */
  async setCells(
    cells: CellAssignment[],
    options: { validationMode?: SetCellsValidationMode } = {}
  ): Promise<WriteResult> {
    return (await this.request(
      'setCells',
      'setCells',
      dropUndefined({ cells, validationMode: options.validationMode })
    )) as WriteResult;
  }

  /**
   * Set the style of a cell or range.
   *
   * @param target - Cell or range address
   * @param style - Style properties to apply
   */
  async setStyle(target: string, style: Style): Promise<void> {
    await this.request('setStyle', 'setStyle', { address: target, style });
  }

  /**
   * Copy a range to another location.
   *
   * @param source - Source range address
   * @param destination - Destination cell address
   * @param options - Copy options
   * @returns Copy result with destination and cell count
   */
  async copyRange(
    source: string,
    destination: string,
    options: { pasteType?: PasteType } = {}
  ): Promise<CopyRangeResult> {
    return (await this.request(
      'copyRange',
      'copyRange',
      dropUndefined({
        source,
        destination,
        pasteType: options.pasteType,
      })
    )) as CopyRangeResult;
  }

  /**
   * Scale numeric values in a range by a factor.
   * This is a composite operation that reads the range, filters numeric cells,
   * and calls setCells with the scaled values.
   *
   * @param range - Range address
   * @param factor - Multiplication factor
   * @param options - Scale options
   * @returns Write result, or null if no numeric cells were found
   */
  async scaleRange(
    range: string,
    factor: number,
    options: { skipFormulas?: boolean } = {}
  ): Promise<WriteResult | null> {
    const skipFormulas = options.skipFormulas ?? true;
    const data = await this.readRange(range);

    const assignments: CellAssignment[] = [];
    for (const row of data) {
      for (const cell of row) {
        const value = cell.value;
        const hasFormula = Boolean(cell.formula);

        if (typeof value === 'number' && (!hasFormula || !skipFormulas)) {
          assignments.push({ address: cell.address, value: value * factor });
        }
      }
    }

    if (assignments.length === 0) {
      return null;
    }

    return this.setCells(assignments);
  }

  /**
   * Find cells matching a pattern.
   *
   * @param matcher - String, number, boolean, RegExp, or array of strings/RegExps
   * @param options - Search options
   * @returns Array of matching cells
   */
  async findCells(
    matcher: Matcher,
    options: {
      in?: string;
      context?: number;
      limit?: number;
      offset?: number;
      formulas?: boolean;
    } = {}
  ): Promise<SearchCell[]> {
    const result = (await this.request(
      'findCells',
      'findCells',
      dropUndefined({
        matcher: serializeMatcher(matcher),
        in: options.in,
        context: options.context ?? 2,
        limit: options.limit ?? 20,
        offset: options.offset ?? 0,
        formulas: options.formulas,
      })
    )) as { matches?: SearchCell[] };
    return result.matches ?? [];
  }

  /**
   * Find rows containing cells matching a pattern.
   *
   * @param matcher - String, number, boolean, RegExp, or array of strings/RegExps
   * @param options - Search options
   * @returns Array of matching rows
   */
  async findRows(
    matcher: Matcher,
    options: {
      in?: string;
      context?: number;
      limit?: number;
      offset?: number;
    } = {}
  ): Promise<SearchRow[]> {
    const result = (await this.request(
      'findRows',
      'findRows',
      dropUndefined({
        matcher: serializeMatcher(matcher),
        in: options.in,
        context: options.context,
        limit: options.limit ?? 20,
        offset: options.offset ?? 0,
      })
    )) as { matches?: SearchRow[] };
    return result.matches ?? [];
  }

  /**
   * Find and replace text in cells.
   *
   * @param find - String or RegExp to find
   * @param replace - Replacement string
   * @param options - Find and replace options
   * @returns Result with count of replacements made
   */
  async findAndReplace(
    find: ReplaceMatcher,
    replace: string,
    options: {
      in?: string;
      matchCase?: boolean;
      wholeCell?: boolean;
      inFormulas?: boolean;
      limit?: number;
    } = {}
  ): Promise<FindAndReplaceResult> {
    return (await this.request(
      'findAndReplace',
      'findAndReplace',
      dropUndefined({
        find: serializeMatcher(find),
        replace,
        in: options.in,
        matchCase: options.matchCase,
        wholeCell: options.wholeCell,
        inFormulas: options.inFormulas,
        limit: options.limit,
      })
    )) as FindAndReplaceResult;
  }

  /**
   * Insert rows after a specified row.
   *
   * @param sheet - Sheet name
   * @param row - Row number after which to insert (1-based)
   * @param count - Number of rows to insert (default: 1)
   */
  async insertRowAfter(sheet: string, row: number, count: number = 1): Promise<void> {
    await this.request('insertRowAfter', 'insertRowAfter', { sheet, row, count });
  }

  /**
   * Delete rows from a sheet.
   *
   * @param sheet - Sheet name
   * @param row - Starting row number (1-based)
   * @param count - Number of rows to delete (default: 1)
   */
  async deleteRows(sheet: string, row: number, count: number = 1): Promise<void> {
    await this.request('deleteRows', 'deleteRows', { sheet, row, count });
  }

  /**
   * Insert columns after a specified column.
   *
   * @param sheet - Sheet name
   * @param column - Column number (1-based) or letter after which to insert
   * @param count - Number of columns to insert (default: 1)
   */
  async insertColumnAfter(sheet: string, column: number | string, count: number = 1): Promise<void> {
    await this.request('insertColumnAfter', 'insertColumnAfter', { sheet, column, count });
  }

  /**
   * Delete columns from a sheet.
   *
   * @param sheet - Sheet name
   * @param column - Starting column number (1-based) or letter
   * @param count - Number of columns to delete (default: 1)
   */
  async deleteColumns(sheet: string, column: number | string, count: number = 1): Promise<void> {
    await this.request('deleteColumns', 'deleteColumns', { sheet, column, count });
  }

  /**
   * Set properties for a range of rows.
   *
   * @param sheet - Sheet name
   * @param fromRow - Starting row number (1-based)
   * @param toRow - Ending row number (1-based)
   * @param properties - Properties to set (e.g., height, hidden)
   */
  async setRowProperties(
    sheet: string,
    fromRow: number,
    toRow: number,
    properties: RowProperties
  ): Promise<void> {
    await this.request('setRowProperties', 'setRowProperties', {
      sheet,
      fromRow,
      toRow,
      properties,
    });
  }

  /**
   * Set properties for a range of columns.
   *
   * @param sheet - Sheet name
   * @param fromCol - Starting column number (1-based) or letter
   * @param toCol - Ending column number (1-based) or letter
   * @param properties - Properties to set (e.g., width, hidden)
   */
  async setColumnProperties(
    sheet: string,
    fromCol: number | string,
    toCol: number | string,
    properties: ColumnProperties
  ): Promise<void> {
    await this.request('setColumnProperties', 'setColumnProperties', {
      sheet,
      fromCol,
      toCol,
      properties,
    });
  }

  /**
   * Auto-fit column widths to content.
   *
   * @param sheet - Sheet name
   * @param columns - Specific columns to fit (default: all)
   * @param options - Auto-fit options
   * @returns Map of column letters to width results
   */
  async autoFitColumns(
    sheet: string,
    columns?: (number | string)[],
    options: { minWidth?: number; maxWidth?: number; padding?: number } = {}
  ): Promise<Record<string, AutoFitColumnResult>> {
    const result = (await this.request(
      'autoFitColumns',
      'autoFitColumns',
      dropUndefined({
        sheet,
        columns,
        minWidth: options.minWidth,
        maxWidth: options.maxWidth,
        padding: options.padding,
      })
    )) as { columns?: Record<string, AutoFitColumnResult> };
    return result.columns ?? {};
  }

  /**
   * Auto-fit row heights to content.
   *
   * @param sheet - Sheet name
   * @param rows - Specific rows to fit (default: all)
   * @param options - Auto-fit options
   * @returns Map of row numbers to height results
   */
  async autoFitRows(
    sheet: string,
    rows?: number[],
    options: { minHeight?: number; maxHeight?: number } = {}
  ): Promise<Record<string, AutoFitRowResult>> {
    const result = (await this.request(
      'autoFitRows',
      'autoFitRows',
      dropUndefined({
        sheet,
        rows,
        minHeight: options.minHeight,
        maxHeight: options.maxHeight,
      })
    )) as { rows?: Record<string, AutoFitRowResult> };
    return result.rows ?? {};
  }

  /**
   * Sort a range by specified columns.
   *
   * @param range - Range address to sort
   * @param keys - Sort keys specifying columns and order
   * @param options - Sort options
   */
  async sortRange(
    range: string,
    keys: SortKey[],
    options: { hasHeader?: boolean } = {}
  ): Promise<void> {
    await this.request(
      'sortRange',
      'sortRange',
      dropUndefined({
        range,
        keys,
        hasHeader: options.hasHeader,
      })
    );
  }

  /**
   * List all defined names in the workbook.
   */
  async listDefinedNames(): Promise<DefinedName[]> {
    return (await this.request('listDefinedNames', 'listDefinedNames', {})) as DefinedName[];
  }

  /**
   * Add a defined name to the workbook.
   *
   * @param name - Name to define
   * @param range - Range address the name refers to
   * @param options - Options including optional scope
   * @returns The created defined name
   */
  async addDefinedName(
    name: string,
    range: string,
    options: { scope?: string } = {}
  ): Promise<DefinedName> {
    return (await this.request(
      'addDefinedName',
      'addDefinedName',
      dropUndefined({ name, range, scope: options.scope })
    )) as DefinedName;
  }

  /**
   * Delete a defined name from the workbook.
   *
   * @param name - Name to delete
   * @param options - Options including optional scope
   * @returns The deleted defined name
   */
  async deleteDefinedName(name: string, options: { scope?: string } = {}): Promise<DefinedName> {
    return (await this.request(
      'deleteDefinedName',
      'deleteDefinedName',
      dropUndefined({ name, scope: options.scope })
    )) as DefinedName;
  }

  /**
   * Get a list object (table) by name.
   *
   * @param name - Name of the list object
   * @returns The list object
   */
  async getListObject(name: string): Promise<ListObject> {
    return (await this.request('getListObject', 'getListObject', { name })) as ListObject;
  }

  /**
   * Add a list object (table) to a sheet.
   *
   * @param sheet - Sheet name
   * @param listObject - List object specification
   * @returns Mutation result with the created list object
   */
  async addListObject(sheet: string, listObject: ListObjectSpec): Promise<ListObjectMutationResult> {
    return (await this.request('addListObject', 'addListObject', {
      sheet,
      listObject,
    })) as ListObjectMutationResult;
  }

  /**
   * Update a list object (table).
   *
   * @param name - Name of the list object to update
   * @param listObject - Updated list object properties
   * @returns Mutation result with the updated list object
   */
  async setListObject(name: string, listObject: ListObjectUpdate): Promise<ListObjectMutationResult> {
    return (await this.request('setListObject', 'setListObject', {
      name,
      listObject,
    })) as ListObjectMutationResult;
  }

  /**
   * Delete a list object (table).
   *
   * @param name - Name of the list object to delete
   * @returns Write result
   */
  async deleteListObject(name: string): Promise<WriteResult> {
    return (await this.request('deleteListObject', 'deleteListObject', { name })) as WriteResult;
  }

  /**
   * Get a data table by address.
   *
   * @param address - Address of the data table
   * @returns The data table
   */
  async getDataTable(address: string): Promise<DataTable> {
    return (await this.request('getDataTable', 'getDataTable', { address })) as DataTable;
  }

  /**
   * Add a data table to a sheet.
   *
   * @param sheet - Sheet name
   * @param dataTable - Data table specification
   * @returns Mutation result with the created data table
   */
  async addDataTable(sheet: string, dataTable: DataTableSpec): Promise<DataTableMutationResult> {
    return (await this.request('addDataTable', 'addDataTable', {
      sheet,
      dataTable,
    })) as DataTableMutationResult;
  }

  /**
   * Delete a data table.
   *
   * @param address - Address of the data table to delete
   * @returns Write result
   */
  async deleteDataTable(address: string): Promise<WriteResult> {
    return (await this.request('deleteDataTable', 'deleteDataTable', { address })) as WriteResult;
  }

  /**
   * List charts in the workbook.
   *
   * @param options - Options to filter by sheet
   * @returns Array of chart summaries
   */
  async listCharts(options: { sheet?: string } = {}): Promise<ChartSummary[]> {
    const result = (await this.request(
      'listCharts',
      'listCharts',
      dropUndefined({ sheet: options.sheet })
    )) as { charts?: ChartSummary[] };
    return result.charts ?? [];
  }

  /**
   * Get a chart by name.
   *
   * @param sheet - Sheet containing the chart
   * @param name - Chart name
   * @returns The chart info
   */
  async getChart(sheet: string, name: string): Promise<ChartInfo> {
    const result = (await this.request('getChart', 'getChart', { sheet, name })) as {
      chart?: ChartInfo;
    };
    return result.chart ?? ({} as ChartInfo);
  }

  /**
   * Add a chart to a sheet.
   *
   * @param sheet - Sheet name
   * @param chart - Chart specification
   * @returns The created chart
   */
  async addChart(sheet: string, chart: ChartSpec): Promise<ChartSpec> {
    const result = (await this.request('addChart', 'addChart', { sheet, chart })) as {
      chart?: ChartSpec;
    };
    return result.chart ?? ({} as ChartSpec);
  }

  /**
   * Update a chart.
   *
   * @param sheet - Sheet containing the chart
   * @param name - Chart name
   * @param chart - Updated chart specification
   * @returns The updated chart
   */
  async setChart(sheet: string, name: string, chart: ChartSpec): Promise<ChartSpec> {
    const result = (await this.request('setChart', 'setChart', { sheet, name, chart })) as {
      chart?: ChartSpec;
    };
    return result.chart ?? ({} as ChartSpec);
  }

  /**
   * Delete a chart.
   *
   * @param sheet - Sheet containing the chart
   * @param name - Chart name
   */
  async deleteChart(sheet: string, name: string): Promise<void> {
    await this.request('deleteChart', 'deleteChart', { sheet, name });
  }

  /**
   * Get conditional formatting rules for a sheet.
   *
   * @param sheet - Sheet name
   * @returns Array of conditional formatting rules
   */
  async getConditionalFormatting(sheet: string): Promise<ConditionalFormattingRule[]> {
    const result = (await this.request('getConditionalFormatting', 'getConditionalFormatting', {
      sheet,
    })) as { rules?: ConditionalFormattingRule[] };
    return result.rules ?? [];
  }

  /**
   * Set conditional formatting rules for a sheet.
   *
   * @param sheet - Sheet name
   * @param rules - Array of conditional formatting rules
   * @param options - Options including whether to clear existing rules
   */
  async setConditionalFormatting(
    sheet: string,
    rules: ConditionalFormattingRule[],
    options: { clear?: boolean } = {}
  ): Promise<void> {
    await this.request(
      'setConditionalFormatting',
      'setConditionalFormatting',
      dropUndefined({ sheet, rules, clear: options.clear })
    );
  }

  /**
   * Remove conditional formatting rules by index.
   *
   * @param sheet - Sheet name
   * @param indices - Indices of rules to remove
   */
  async removeConditionalFormatting(sheet: string, indices: number[]): Promise<void> {
    await this.request('removeConditionalFormatting', 'removeConditionalFormatting', {
      sheet,
      indices,
    });
  }

  /**
   * Evaluate multiple formulas in a sheet context.
   *
   * @param sheet - Sheet name for formula context
   * @param formulas - Array of formula strings
   * @returns Array of formula results
   */
  async evaluateFormulas(sheet: string, formulas: string[]): Promise<FormulaResult[]> {
    return (await this.request('evaluateFormulas', 'evaluateFormulas', {
      sheet,
      formulas,
    })) as FormulaResult[];
  }

  /**
   * Evaluate a single formula in a sheet context.
   * This is a convenience wrapper around evaluateFormulas.
   *
   * @param sheet - Sheet name for formula context
   * @param formula - Formula string
   * @returns The formula result
   */
  async evaluateFormula(sheet: string, formula: string): Promise<FormulaResult> {
    const results = await this.evaluateFormulas(sheet, [formula]);
    return results[0]!;
  }

  /**
   * Get cells that a formula depends on (precedents).
   *
   * @param address - Cell address
   * @param depth - How many levels to trace (default: 1, use Infinity for all)
   * @returns Dependency result with cells and warnings
   */
  async getCellPrecedents(address: string, depth: number | typeof Infinity = 1): Promise<DependencyResult> {
    const rpcDepth = Number.isFinite(depth) ? Math.floor(depth) : -1;
    return (await this.request('getCellPrecedents', 'getCellPrecedents', {
      address,
      depth: rpcDepth,
    })) as DependencyResult;
  }

  /**
   * Get cells that depend on a cell (dependents).
   *
   * @param address - Cell address
   * @param depth - How many levels to trace (default: 1, use Infinity for all)
   * @returns Dependency result with cells and warnings
   */
  async getCellDependents(address: string, depth: number | typeof Infinity = 1): Promise<DependencyResult> {
    const rpcDepth = Number.isFinite(depth) ? Math.floor(depth) : -1;
    return (await this.request('getCellDependents', 'getCellDependents', {
      address,
      depth: rpcDepth,
    })) as DependencyResult;
  }

  /**
   * Trace a cell to its input sources.
   *
   * @param address - Cell address
   * @returns Array of input traces
   */
  async traceToInputs(address: string): Promise<TraceInput[]> {
    return (await this.request('traceToInputs', 'traceToInputs', { address })) as TraceInput[];
  }

  /**
   * Trace a cell to its output destinations.
   *
   * @param address - Cell address
   * @returns Array of output traces
   */
  async traceToOutputs(address: string): Promise<TraceOutput[]> {
    return (await this.request('traceToOutputs', 'traceToOutputs', { address })) as TraceOutput[];
  }

  /**
   * Run a sweep over input combinations and capture outputs.
   *
   * @param inputs - Array of input specifications with addresses and values
   * @param outputs - Array of output cell addresses
   * @param options - Sweep options
   * @returns Sweep result with all combinations
   */
  async sweepInputs(
    inputs: SweepInput[],
    outputs: string[],
    options: { mode?: SweepMode; includeStats?: boolean } = {}
  ): Promise<SweepResult> {
    return (await this.request(
      'sweepInputs',
      'sweepInputs',
      dropUndefined({
        inputs,
        outputs,
        mode: options.mode,
        includeStats: options.includeStats,
      })
    )) as SweepResult;
  }

  /**
   * Alias for sweepInputs.
   */
  async scenarios(
    inputs: SweepInput[],
    outputs: string[],
    options: { mode?: SweepMode; includeStats?: boolean } = {}
  ): Promise<SweepResult> {
    return this.sweepInputs(inputs, outputs, options);
  }

  /**
   * Get a description of a sheet's structure.
   *
   * @param sheet - Sheet name
   * @returns Sheet description
   */
  async describeSheet(sheet: string): Promise<SheetDescription> {
    return (await this.request('describeSheet', 'describeSheet', { sheet })) as SheetDescription;
  }

  /**
   * Get descriptions of all visible sheets.
   * This is a composite operation that calls listSheets and describeSheet.
   *
   * @returns Map of sheet names to descriptions
   */
  async describeSheets(): Promise<Record<string, SheetDescription>> {
    const sheets = await this.listSheets();
    const result: Record<string, SheetDescription> = {};
    for (const sheet of sheets) {
      if (!sheet.hidden) {
        result[sheet.sheet] = await this.describeSheet(sheet.sheet);
      }
    }
    return result;
  }

  /**
   * Look up a value in a table by row and column labels.
   *
   * @param table - Table name or address
   * @param rowLabel - Row label to find
   * @param columnLabel - Column label to find
   * @returns Array of lookup results
   */
  async tableLookup(
    table: string,
    rowLabel: string | number | boolean,
    columnLabel: string | number | boolean
  ): Promise<TableLookupResult[]> {
    return (await this.request('tableLookup', 'tableLookup', {
      table,
      rowLabel,
      columnLabel,
    })) as TableLookupResult[];
  }

  /**
   * Run linting rules on the workbook.
   *
   * @param options - Lint options
   * @returns Lint result with diagnostics
   */
  async lint(
    options: {
      rangeAddresses?: string[];
      skipRuleIds?: string[];
      onlyRuleIds?: string[];
    } = {}
  ): Promise<LintResult> {
    return (await this.request(
      'lint',
      'lint',
      dropUndefined({
        rangeAddresses: options.rangeAddresses,
        skipRuleIds: options.skipRuleIds,
        onlyRuleIds: options.onlyRuleIds,
      })
    )) as LintResult;
  }

  /**
   * Generate a preview image of cell styles.
   *
   * @param range - Range address to preview
   * @returns Data URL of the preview image
   */
  async previewStyles(range: string): Promise<string> {
    const result = (await this.request('previewStyles', 'previewStyles', { address: range })) as {
      contentType?: string;
      data?: string;
    };
    const contentType = result.contentType;
    const data = result.data;
    if (typeof contentType !== 'string' || typeof data !== 'string') {
      throw new WitanProcessError(`Invalid previewStyles result: ${JSON.stringify(result)}`);
    }
    return `data:${contentType};base64,${data}`;
  }

  /**
   * Reduce a list of addresses to minimal non-overlapping ranges.
   *
   * @param addresses - Array of cell or range addresses
   * @returns Array of reduced addresses
   */
  async reduceAddresses(addresses: string[]): Promise<string[]> {
    return (await this.request('reduceAddresses', 'reduceAddresses', { addresses })) as string[];
  }

  /**
   * Get the style of a cell.
   *
   * @param cell - Cell address
   * @returns Style properties
   */
  async getStyle(cell: string): Promise<Style> {
    return (await this.request('getStyle', 'getStyle', { address: cell })) as Style;
  }
}
