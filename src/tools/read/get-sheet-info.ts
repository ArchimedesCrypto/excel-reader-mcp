// Get detailed metadata about a sheet

import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import { GetSheetInfoArgs, GetSheetInfoResponse } from '../../types/index.js';
import { loadWorkbook } from '../../utils/workbook-cache.js';
import { existsSync } from 'fs';
import * as XLSX from 'xlsx';
import { indexToColumnLetter } from '../../utils/cell-utils.js';

export function validateGetSheetInfoArgs(args: any): args is GetSheetInfoArgs {
  return (
    typeof args === 'object' &&
    args !== null &&
    typeof args.filePath === 'string' &&
    (args.sheetName === undefined || typeof args.sheetName === 'string')
  );
}

export function getSheetInfo(args: GetSheetInfoArgs): GetSheetInfoResponse {
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

    // Get the range of the sheet
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    const totalRows = range.e.r - range.s.r + 1;
    const totalColumns = range.e.c - range.s.c + 1;
    const usedRange = XLSX.utils.encode_range(range);

    // Get merged cells
    const mergedCells = worksheet['!merges']
      ? worksheet['!merges'].map((merge: any) => XLSX.utils.encode_range(merge))
      : [];

    // Check for formulas
    let hasFormulas = false;
    for (const cellAddr in worksheet) {
      if (cellAddr[0] === '!') continue;
      const cell = worksheet[cellAddr];
      if (cell.f) {
        hasFormulas = true;
        break;
      }
    }

    // Get column info by examining first row and data types
    const columnInfo = [];
    const firstRowData = XLSX.utils.sheet_to_json(worksheet, { 
      header: 1,
      raw: false 
    })[0] as any[] || [];

    for (let col = range.s.c; col <= range.e.c; col++) {
      const letter = indexToColumnLetter(col);
      const header = firstRowData[col];

      // Determine data type by sampling column values
      const columnValues = [];
      for (let row = range.s.r + 1; row <= Math.min(range.s.r + 100, range.e.r); row++) {
        const cellAddr = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = worksheet[cellAddr];
        if (cell && cell.v !== undefined && cell.v !== null && cell.v !== '') {
          columnValues.push(cell);
        }
      }

      let dataType: 'mixed' | 'string' | 'number' | 'date' | 'empty' = 'empty';
      if (columnValues.length === 0) {
        dataType = 'empty';
      } else {
        const types = new Set(columnValues.map((cell: any) => cell.t));
        if (types.size > 1) {
          dataType = 'mixed';
        } else {
          const type = Array.from(types)[0];
          switch (type) {
            case 'n':
              dataType = 'number';
              break;
            case 's':
              dataType = 'string';
              break;
            case 'd':
              dataType = 'date';
              break;
            default:
              dataType = 'mixed';
          }
        }
      }

      columnInfo.push({
        letter,
        header: header !== undefined ? String(header) : undefined,
        dataType,
      });
    }

    return {
      name: selectedSheetName,
      dimensions: {
        totalRows,
        totalColumns,
        usedRange,
      },
      mergedCells,
      hasFormulas,
      columnInfo,
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      `Error reading sheet info: ${error instanceof Error ? error.message : String(error)}`
    );
  }
}
