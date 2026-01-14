// Data Analysis Operations Type Definitions

export interface FilterCondition {
  column: string;  // Column letter or header name
  operator: 'equals' | 'not_equals' | 'contains' | 'not_contains' | 
            'starts_with' | 'ends_with' | 'greater_than' | 'less_than' |
            'greater_equal' | 'less_equal' | 'is_empty' | 'is_not_empty';
  value?: any;
}

export interface FilterRowsArgs {
  filePath: string;
  sheetName?: string;
  conditions: FilterCondition[];
  logic?: 'AND' | 'OR';  // Default AND
  startRow?: number;
  maxRows?: number;
}

export interface FilterRowsResponse {
  totalRowsScanned: number;
  matchingRows: number;
  data: Record<string, any>[];
  columns: string[];
  hasMore: boolean;
  nextChunk?: {
    startRow: number;
  };
}

export interface SortCriteria {
  column: string;  // Column letter or header name
  order: 'asc' | 'desc';
}

export interface SortDataArgs {
  filePath: string;
  sheetName?: string;
  sortBy: SortCriteria[];
  hasHeaders?: boolean;  // Default true
  startRow?: number;
  maxRows?: number;
}

export interface SortDataResponse {
  totalRows: number;
  sortedBy: SortCriteria[];
  columns: string[];
  data: Record<string, any>[];
  hasMore: boolean;
  nextChunk?: {
    startRow: number;
  };
}

export interface GetUniqueValuesArgs {
  filePath: string;
  sheetName?: string;
  column: string;  // Column letter or header name
  includeCount?: boolean;  // Include occurrence count
  sortBy?: 'value' | 'count';
  sortOrder?: 'asc' | 'desc';
}

export interface UniqueValue {
  value: any;
  count?: number;
  firstOccurrence?: number;  // Row number
}

export interface GetUniqueValuesResponse {
  column: string;
  totalRows: number;
  uniqueCount: number;
  values: UniqueValue[];
}

export interface GetColumnStatsArgs {
  filePath: string;
  sheetName?: string;
  columns?: string[];  // If not provided, analyze all numeric columns
}

export interface ColumnStats {
  column: string;
  header?: string;
  dataType: string;
  totalValues: number;
  emptyCount: number;
  // Numeric stats - only for numeric columns
  numericStats?: {
    sum: number;
    average: number;
    min: number;
    max: number;
    median: number;
    standardDeviation: number;
  };
  // String stats - only for string columns
  stringStats?: {
    minLength: number;
    maxLength: number;
    avgLength: number;
  };
}

export interface GetColumnStatsResponse {
  sheet: string;
  totalRows: number;
  columns: ColumnStats[];
}
