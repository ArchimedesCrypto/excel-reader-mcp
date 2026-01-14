// Export/Conversion Operations Type Definitions

export interface ExportToCsvArgs {
  filePath: string;
  sheetName?: string;
  outputPath?: string;  // If not provided, return as string
  delimiter?: string;   // Default comma
  includeHeaders?: boolean;  // Default true
  dateFormat?: string;  // Default ISO 8601
}

export interface ExportToCsvResponse {
  sheet: string;
  totalRows: number;
  totalColumns: number;
  outputPath?: string;  // If saved to file
  csvContent?: string;  // If returned as string - chunked if large
  hasMore?: boolean;
  nextChunk?: {
    startRow: number;
  };
}

export interface ExportToJsonArgs {
  filePath: string;
  sheetName?: string;
  outputPath?: string;  // If not provided, return as string
  format?: 'array' | 'objects';  // Default 'objects'
  includeMetadata?: boolean;
  dateFormat?: string;
}

export interface ExportToJsonResponse {
  sheet: string;
  totalRows: number;
  outputPath?: string;
  jsonContent?: string | object[];  // Chunked if large
  hasMore?: boolean;
  nextChunk?: {
    startRow: number;
  };
}
