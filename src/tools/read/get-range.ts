// Get values from a specific range

import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import { GetRangeArgs, GetRangeResponse } from '../../types/index.js';
import { loadWorkbook } from '../../utils/workbook-cache.js';
import { existsSync } from 'fs';
import * as XLSX from 'xlsx';
import { isValidRange, parseRange, getRangeDimensions } from '../../utils/cell-utils.js';
import { applyChunking } from '../../utils/chunking.js';

export function validateGetRangeArgs(args: any): args is GetRangeArgs {
  return (
    typeof args === 'object' &&
    args !== null &&
    typeof args.filePath === 'string' &&
    typeof args.range === 'string' &&
    (args.sheetName === undefined || typeof args.sheetName === 'string') &&
    (args.includeHeaders === undefined || typeof args.includeHeaders === 'boolean')
  );
}

export function getRange(args: GetRangeArgs): GetRangeResponse {
  const { filePath, sheetName, range, includeHeaders = false } = args;

  if (!existsSync(filePath)) {
    throw new McpError(
      ErrorCode.InvalidRequest,
      `File not found: ${filePath}`
    );
  }

  if (!isValidRange(range)) {
    throw new McpError(
      ErrorCode.InvalidRequest,
      `Invalid range: ${range}`
    );
  }

  try {
    const workbook = loadWorkbook(filePath);
    const selectedSheetName = sheetName || workbook.SheetNames[0];
    const worksheet = workbook.Sheets[selectedSheetName];

    if (!worksheet) {
      throw new McpError(
        ErrorCode.InvalidRequest,
        `Sheet not found: ${selectedSheetName}`
      );
    }

    const dimensions = getRangeDimensions(range);
    const { start, end } = parseRange(range);

    // Extract data from the range
    const data: any[][] = [];
    for (let row = start.row; row <= end.row; row++) {
      const rowData: any[] = [];
      for (let col = start.colIndex; col <= end.colIndex; col++) {
        const cellAddr = XLSX.utils.encode_cell({ r: row - 1, c: col });
        const cell = worksheet[cellAddr];
        rowData.push(cell ? cell.v : null);
      }
      data.push(rowData);
    }

    // Handle headers if requested
    let headers: string[] | undefined;
    let actualData = data;
    
    if (includeHeaders && data.length > 0) {
      headers = data[0].map((h: any) => String(h || ''));
      actualData = data.slice(1);
    }

    // For large ranges, we might need to chunk (though ranges are typically bounded)
    const hasMore = false; // Ranges are bounded, no pagination needed
    const nextRange = undefined;

    return {
      range,
      dimensions,
      headers,
      data: actualData,
      hasMore,
      nextRange,
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      `Error reading range: ${error instanceof Error ? error.message : String(error)}`
    );
  }
}
