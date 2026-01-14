// Get formulas from cells in a range

import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import { GetFormulasArgs, GetFormulasResponse, FormulaCell } from '../../types/index.js';
import { loadWorkbook } from '../../utils/workbook-cache.js';
import { existsSync } from 'fs';
import * as XLSX from 'xlsx';
import { isValidRange, parseRange } from '../../utils/cell-utils.js';

export function validateGetFormulasArgs(args: any): args is GetFormulasArgs {
  return (
    typeof args === 'object' &&
    args !== null &&
    typeof args.filePath === 'string' &&
    (args.sheetName === undefined || typeof args.sheetName === 'string') &&
    (args.range === undefined || typeof args.range === 'string')
  );
}

export function getFormulas(args: GetFormulasArgs): GetFormulasResponse {
  const { filePath, sheetName, range } = args;

  if (!existsSync(filePath)) {
    throw new McpError(
      ErrorCode.InvalidRequest,
      `File not found: ${filePath}`
    );
  }

  if (range && !isValidRange(range)) {
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

    const formulas: FormulaCell[] = [];
    let searchRange: { start: any; end: any } | null = null;

    if (range) {
      searchRange = parseRange(range);
    }

    // Iterate through cells to find formulas
    for (const cellAddr in worksheet) {
      if (cellAddr[0] === '!') continue; // Skip special properties

      const cell = worksheet[cellAddr];
      if (!cell || !cell.f) continue; // Skip non-formula cells

      // If range is specified, check if cell is within range
      if (searchRange) {
        const decoded = XLSX.utils.decode_cell(cellAddr);
        const row = decoded.r + 1; // Convert to 1-based
        const col = decoded.c;

        if (
          row < searchRange.start.row ||
          row > searchRange.end.row ||
          col < searchRange.start.colIndex ||
          col > searchRange.end.colIndex
        ) {
          continue; // Cell is outside the specified range
        }
      }

      formulas.push({
        cell: cellAddr,
        formula: cell.f,
        calculatedValue: cell.v,
      });
    }

    return {
      sheet: selectedSheetName,
      range: range || worksheet['!ref'] || 'A1',
      formulas,
      totalFormulas: formulas.length,
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      `Error reading formulas: ${error instanceof Error ? error.message : String(error)}`
    );
  }
}
