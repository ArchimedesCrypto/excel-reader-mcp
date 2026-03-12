// Read Operations Type Definitions

export interface ReadExcelArgs {
  filePath: string;
  sheetName?: string;
  startRow?: number;
  maxRows?: number;
}

export interface ExcelChunk {
  rowStart: number;
  rowEnd: number;
  columns: string[];
  data: Record<string, any>[];
}

export interface ExcelSheetData {
  name: string;
  totalRows: number;
  totalColumns: number;
  chunk: ExcelChunk;
  hasMore: boolean;
  nextChunk?: {
    rowStart: number;
    columns: string[];
  };
}

export interface ExcelData {
  fileName: string;
  totalSheets: number;
  currentSheet: ExcelSheetData;
}

export interface ListSheetsArgs {
  filePath: string;
}

export interface ListSheetsResponse {
  fileName: string;
  sheets: Array<{
    name: string;
    index: number;
    hidden: boolean;
  }>;
  totalSheets: number;
}

export interface GetCellArgs {
  filePath: string;
  sheetName?: string;
  cell: string;  // e.g., "A1", "B5"
}

export interface GetCellResponse {
  cell: string;
  value: any;
  type: 'string' | 'number' | 'boolean' | 'date' | 'formula' | 'empty';
  formula?: string;  // If cell contains a formula
  formatted?: string; // Formatted display value
}

export interface GetRangeArgs {
  filePath: string;
  sheetName?: string;
  range: string;  // e.g., "A1:C10"
  includeHeaders?: boolean;  // Treat first row as headers
}

export interface GetRangeResponse {
  range: string;
  dimensions: {
    startRow: number;
    endRow: number;
    startCol: string;
    endCol: string;
    rowCount: number;
    colCount: number;
  };
  headers?: string[];
  data: any[][];  // 2D array of values
  // Chunking support for large ranges
  hasMore: boolean;
  nextRange?: string;
}

export interface SearchTextArgs {
  filePath: string;
  sheetName?: string;  // If not provided, search all sheets
  searchText: string;
  matchCase?: boolean;
  matchWholeCell?: boolean;
  maxResults?: number;  // Default 100
}

export interface SearchResult {
  sheet: string;
  cell: string;
  value: string;
  row: number;
  column: string;
}

export interface SearchTextResponse {
  query: string;
  totalMatches: number;
  results: SearchResult[];
  hasMore: boolean;
}

export interface GetSheetInfoArgs {
  filePath: string;
  sheetName?: string;
}

export interface GetSheetInfoResponse {
  name: string;
  dimensions: {
    totalRows: number;
    totalColumns: number;
    usedRange: string;  // e.g., "A1:Z100"
  };
  mergedCells: string[];  // e.g., ["A1:B2", "D5:D10"]
  hasFormulas: boolean;
  columnInfo: Array<{
    letter: string;
    header?: string;  // First row value
    dataType: 'mixed' | 'string' | 'number' | 'date' | 'empty';
  }>;
}

export interface GetFormulasArgs {
  filePath: string;
  sheetName?: string;
  range?: string;  // If not provided, get all formulas
}

export interface FormulaCell {
  cell: string;
  formula: string;
  calculatedValue: any;
}

export interface GetFormulasResponse {
  sheet: string;
  range: string;
  formulas: FormulaCell[];
  totalFormulas: number;
}

export interface GetMergedCellsArgs {
  filePath: string;
  sheetName?: string;
}

export interface MergedRange {
  range: string;
  startCell: string;
  endCell: string;
  rowSpan: number;
  colSpan: number;
  value: any;  // Value in the merged cell
}

export interface GetMergedCellsResponse {
  sheet: string;
  mergedRanges: MergedRange[];
  totalMerged: number;
}
