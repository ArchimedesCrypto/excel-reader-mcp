// Get the value of a specific cell

import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import { GetCellArgs, GetCellResponse } from '../../types/index.js';
import { loadWorkbook } from '../../utils/workbook-cache.js';
import { existsSync } from 'fs';
import * as XLSX from 'xlsx';
import { isValidCellAddress } from '../../utils/validation.js';

export function validateGetCellArgs(args: any): args is GetCellArgs {
  return (
    typeof args === 'object' &&
    args !== null &&
    typeof args.filePath === 'string' &&
    typeof args.cell === 'string' &&
    (args.sheetName === undefined || typeof args.sheetName === 'string')
  );
}

export function getCell(args: GetCellArgs): GetCellResponse {
  const { filePath, sheetName, cell } = args;

  if (!existsSync(filePath)) {
    throw new McpError(
      ErrorCode.InvalidRequest,
      `File not found: ${filePath}`
    );
  }

  if (!isValidCellAddress(cell)) {
    throw new McpError(
      ErrorCode.InvalidRequest,
      `Invalid cell address: ${cell}`
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

    const cellData = worksheet[cell];

    if (!cellData) {
      return {
        cell,
        value: null,
        type: 'empty',
      };
    }

    // Determine cell type
    let type: 'string' | 'number' | 'boolean' | 'date' | 'formula' | 'empty' = 'empty';
    let value: any = cellData.v;
    let formula: string | undefined;
    let formatted: string | undefined;

    if (cellData.f) {
      type = 'formula';
      formula = cellData.f;
    } else if (cellData.t === 'n') {
      type = 'number';
    } else if (cellData.t === 's') {
      type = 'string';
    } else if (cellData.t === 'b') {
      type = 'boolean';
    } else if (cellData.t === 'd') {
      type = 'date';
    }

    // Get formatted value if available
    if (cellData.w) {
      formatted = cellData.w;
    }

    return {
      cell,
      value,
      type,
      formula,
      formatted,
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      `Error reading cell: ${error instanceof Error ? error.message : String(error)}`
    );
  }
}
