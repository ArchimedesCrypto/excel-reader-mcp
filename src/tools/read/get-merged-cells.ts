// Get all merged cell ranges in a sheet

import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import { GetMergedCellsArgs, GetMergedCellsResponse, MergedRange } from '../../types/index.js';
import { loadWorkbook } from '../../utils/workbook-cache.js';
import { existsSync } from 'fs';
import * as XLSX from 'xlsx';

export function validateGetMergedCellsArgs(args: any): args is GetMergedCellsArgs {
  return (
    typeof args === 'object' &&
    args !== null &&
    typeof args.filePath === 'string' &&
    (args.sheetName === undefined || typeof args.sheetName === 'string')
  );
}

export function getMergedCells(args: GetMergedCellsArgs): GetMergedCellsResponse {
  const { filePath, sheetName } = args;

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

    const mergedRanges: MergedRange[] = [];

    if (worksheet['!merges']) {
      for (const merge of worksheet['!merges']) {
        const range = XLSX.utils.encode_range(merge);
        const startCell = XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c });
        const endCell = XLSX.utils.encode_cell({ r: merge.e.r, c: merge.e.c });
        const rowSpan = merge.e.r - merge.s.r + 1;
        const colSpan = merge.e.c - merge.s.c + 1;

        // Get the value from the first cell in the merged range
        const cell = worksheet[startCell];
        const value = cell ? cell.v : null;

        mergedRanges.push({
          range,
          startCell,
          endCell,
          rowSpan,
          colSpan,
          value,
        });
      }
    }

    return {
      sheet: selectedSheetName,
      mergedRanges,
      totalMerged: mergedRanges.length,
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      `Error reading merged cells: ${error instanceof Error ? error.message : String(error)}`
    );
  }
}
