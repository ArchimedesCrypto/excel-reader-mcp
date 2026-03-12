// Central export for all type definitions

export * from './read-operations.js';
export * from './analysis.js';
export * from './export.js';

// Common utility types
export interface ChunkingOptions {
  maxResponseSize?: number;  // Default 100KB
  maxRows?: number;
  startRow?: number;
}

export interface ChunkedResult<T> {
  data: T[];
  totalItems: number;
  hasMore: boolean;
  nextChunk?: {
    startRow: number;
    [key: string]: any;
  };
}

// Error codes
export enum ExcelMcpErrorCode {
  FILE_NOT_FOUND = 'FILE_NOT_FOUND',
  SHEET_NOT_FOUND = 'SHEET_NOT_FOUND',
  INVALID_CELL_ADDRESS = 'INVALID_CELL_ADDRESS',
  INVALID_RANGE = 'INVALID_RANGE',
  INVALID_COLUMN = 'INVALID_COLUMN',
  PARSE_ERROR = 'PARSE_ERROR',
  EXPORT_ERROR = 'EXPORT_ERROR',
}
