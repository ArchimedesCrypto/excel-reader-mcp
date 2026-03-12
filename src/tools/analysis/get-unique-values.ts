// Get distinct values in a column

import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import { GetUniqueValuesArgs, GetUniqueValuesResponse, UniqueValue } from '../../types/index.js';
import { loadWorkbook } from '../../utils/workbook-cache.js';
import { existsSync } from 'fs';
import * as XLSX from 'xlsx';

export function validateGetUniqueValuesArgs(args: any): args is GetUniqueValuesArgs {
  return (
    typeof args === 'object' &&
    args !== null &&
    typeof args.filePath === 'string' &&
    typeof args.column === 'string' &&
    (args.sheetName === undefined || typeof args.sheetName === 'string') &&
    (args.includeCount === undefined || typeof args.includeCount === 'boolean') &&
    (args.sortBy === undefined || args.sortBy === 'value' || args.sortBy === 'count') &&
    (args.sortOrder === undefined || args.sortOrder === 'asc' || args.sortOrder === 'desc')
  );
}

export function getUniqueValues(args: GetUniqueValuesArgs): GetUniqueValuesResponse {
  const {
    filePath,
    sheetName,
    column,
    includeCount = false,
    sortBy = 'value',
    sortOrder = 'asc',
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

    // Check if column exists
    if (allData.length > 0 && !(column in allData[0])) {
      throw new McpError(
        ErrorCode.InvalidRequest,
        `Column not found: ${column}`
      );
    }

    // Build unique values map
    const valueMap = new Map<any, { count: number; firstOccurrence: number }>();
    
    allData.forEach((row, index) => {
      const value = row[column];
      if (valueMap.has(value)) {
        const entry = valueMap.get(value)!;
        entry.count++;
      } else {
        valueMap.set(value, { count: 1, firstOccurrence: index + 1 });
      }
    });

    // Convert to array
    let values: UniqueValue[] = Array.from(valueMap.entries()).map(([value, info]) => ({
      value,
      count: includeCount ? info.count : undefined,
      firstOccurrence: info.firstOccurrence,
    }));

    // Sort values
    values.sort((a, b) => {
      let comparison = 0;
      
      if (sortBy === 'value') {
        if (a.value < b.value) comparison = -1;
        else if (a.value > b.value) comparison = 1;
      } else if (sortBy === 'count') {
        comparison = (a.count || 0) - (b.count || 0);
      }
      
      return sortOrder === 'asc' ? comparison : -comparison;
    });

    return {
      column,
      totalRows: allData.length,
      uniqueCount: values.length,
      values,
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      `Error getting unique values: ${error instanceof Error ? error.message : String(error)}`
    );
  }
}
