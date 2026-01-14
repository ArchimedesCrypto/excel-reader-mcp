// Filter rows based on column criteria

import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import { FilterRowsArgs, FilterRowsResponse, FilterCondition } from '../../types/index.js';
import { loadWorkbook } from '../../utils/workbook-cache.js';
import { existsSync } from 'fs';
import * as XLSX from 'xlsx';
import { applyChunking } from '../../utils/chunking.js';

export function validateFilterRowsArgs(args: any): args is FilterRowsArgs {
  return (
    typeof args === 'object' &&
    args !== null &&
    typeof args.filePath === 'string' &&
    Array.isArray(args.conditions) &&
    (args.sheetName === undefined || typeof args.sheetName === 'string') &&
    (args.logic === undefined || args.logic === 'AND' || args.logic === 'OR') &&
    (args.startRow === undefined || typeof args.startRow === 'number') &&
    (args.maxRows === undefined || typeof args.maxRows === 'number')
  );
}

function evaluateCondition(value: any, condition: FilterCondition): boolean {
  const { operator, value: condValue } = condition;

  switch (operator) {
    case 'equals':
      return value == condValue;
    case 'not_equals':
      return value != condValue;
    case 'contains':
      return String(value).includes(String(condValue));
    case 'not_contains':
      return !String(value).includes(String(condValue));
    case 'starts_with':
      return String(value).startsWith(String(condValue));
    case 'ends_with':
      return String(value).endsWith(String(condValue));
    case 'greater_than':
      return Number(value) > Number(condValue);
    case 'less_than':
      return Number(value) < Number(condValue);
    case 'greater_equal':
      return Number(value) >= Number(condValue);
    case 'less_equal':
      return Number(value) <= Number(condValue);
    case 'is_empty':
      return value === null || value === undefined || value === '';
    case 'is_not_empty':
      return value !== null && value !== undefined && value !== '';
    default:
      return false;
  }
}

export function filterRows(args: FilterRowsArgs): FilterRowsResponse {
  const {
    filePath,
    sheetName,
    conditions,
    logic = 'AND',
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

    // Filter data based on conditions
    const filteredData = allData.filter((row) => {
      const results = conditions.map((condition) => {
        const columnValue = row[condition.column];
        return evaluateCondition(columnValue, condition);
      });

      if (logic === 'AND') {
        return results.every((r) => r);
      } else {
        return results.some((r) => r);
      }
    });

    const totalRowsScanned = allData.length;
    const matchingRows = filteredData.length;

    // Apply chunking
    const chunked = applyChunking(filteredData, { startRow, maxRows });

    return {
      totalRowsScanned,
      matchingRows,
      data: chunked.data,
      columns,
      hasMore: chunked.hasMore,
      nextChunk: chunked.nextChunk,
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      `Error filtering rows: ${error instanceof Error ? error.message : String(error)}`
    );
  }
}
