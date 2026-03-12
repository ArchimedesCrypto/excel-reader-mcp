// Sort data by one or more columns

import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import { SortDataArgs, SortDataResponse, SortCriteria } from '../../types/index.js';
import { loadWorkbook } from '../../utils/workbook-cache.js';
import { existsSync } from 'fs';
import * as XLSX from 'xlsx';
import { applyChunking } from '../../utils/chunking.js';

export function validateSortDataArgs(args: any): args is SortDataArgs {
  return (
    typeof args === 'object' &&
    args !== null &&
    typeof args.filePath === 'string' &&
    Array.isArray(args.sortBy) &&
    (args.sheetName === undefined || typeof args.sheetName === 'string') &&
    (args.hasHeaders === undefined || typeof args.hasHeaders === 'boolean') &&
    (args.startRow === undefined || typeof args.startRow === 'number') &&
    (args.maxRows === undefined || typeof args.maxRows === 'number')
  );
}

export function sortData(args: SortDataArgs): SortDataResponse {
  const {
    filePath,
    sheetName,
    sortBy,
    hasHeaders = true,
    startRow = 0,
    maxRows,
  } = args;

  if (!existsSync(filePath)) {
    throw new McpError(
      ErrorCode.InvalidRequest,
      `File not found: ${filePath}`
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

    const allData = XLSX.utils.sheet_to_json(worksheet, {
      raw: true,
      dateNF: 'yyyy-mm-dd'
    }) as Record<string, any>[];

    const columns = allData.length > 0 ? Object.keys(allData[0]) : [];

    // Sort the data
    const sortedData = [...allData].sort((a, b) => {
      for (const criteria of sortBy) {
        const aVal = a[criteria.column];
        const bVal = b[criteria.column];

        let comparison = 0;
        if (aVal < bVal) comparison = -1;
        else if (aVal > bVal) comparison = 1;

        if (comparison !== 0) {
          return criteria.order === 'asc' ? comparison : -comparison;
        }
      }
      return 0;
    });

    // Apply chunking
    const chunked = applyChunking(sortedData, { startRow, maxRows });

    return {
      totalRows: sortedData.length,
      sortedBy: sortBy,
      columns,
      data: chunked.data,
      hasMore: chunked.hasMore,
      nextChunk: chunked.nextChunk,
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      `Error sorting data: ${error instanceof Error ? error.message : String(error)}`
    );
  }
}
