// List all sheets in an Excel workbook

import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import { ListSheetsArgs, ListSheetsResponse } from '../../types/index.js';
import { loadWorkbook } from '../../utils/workbook-cache.js';
import { existsSync } from 'fs';
import path from 'path';

export function validateListSheetsArgs(args: any): args is ListSheetsArgs {
  return (
    typeof args === 'object' &&
    args !== null &&
    typeof args.filePath === 'string'
  );
}

export function listSheets(args: ListSheetsArgs): ListSheetsResponse {
  const { filePath } = args;

  if (!existsSync(filePath)) {
    throw new McpError(
      ErrorCode.InvalidRequest,
      `File not found: ${filePath}`
    );
  }

  try {
    const workbook = loadWorkbook(filePath);
    const fileName = path.basename(filePath);

    const sheets = workbook.SheetNames.map((name, index) => {
      const sheet = workbook.Sheets[name];
      // Check if sheet is hidden (SheetJS stores this in the Workbook property)
      const hidden = workbook.Workbook?.Sheets?.[index]?.Hidden === 1;

      return {
        name,
        index,
        hidden,
      };
    });

    return {
      fileName,
      sheets,
      totalSheets: sheets.length,
    };
  } catch (error) {
    throw new McpError(
      ErrorCode.InternalError,
      `Error reading Excel file: ${error instanceof Error ? error.message : String(error)}`
    );
  }
}
