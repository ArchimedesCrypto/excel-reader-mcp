// Export a sheet to CSV format

import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import { ExportToCsvArgs, ExportToCsvResponse } from '../../types/index.js';
import { loadWorkbook } from '../../utils/workbook-cache.js';
import { existsSync, writeFileSync } from 'fs';
import * as XLSX from 'xlsx';

export function validateExportToCsvArgs(args: any): args is ExportToCsvArgs {
  return (
    typeof args === 'object' &&
    args !== null &&
    typeof args.filePath === 'string' &&
    (args.sheetName === undefined || typeof args.sheetName === 'string') &&
    (args.outputPath === undefined || typeof args.outputPath === 'string') &&
    (args.delimiter === undefined || typeof args.delimiter === 'string') &&
    (args.includeHeaders === undefined || typeof args.includeHeaders === 'boolean') &&
    (args.dateFormat === undefined || typeof args.dateFormat === 'string')
  );
}

export function exportToCsv(args: ExportToCsvArgs): ExportToCsvResponse {
  const {
    filePath,
    sheetName,
    outputPath,
    delimiter = ',',
    includeHeaders = true,
    dateFormat = 'yyyy-mm-dd',
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

    // Get the data as array of arrays
    const data = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      raw: false,
      dateNF: dateFormat
    }) as any[][];

    let totalRows = data.length;
    let totalColumns = data.length > 0 ? data[0].length : 0;

    // Build CSV content
    let csvContent = '';
    
    if (includeHeaders && data.length > 0) {
      // First row as headers
      csvContent += data[0].map((cell: any) => {
        const str = String(cell || '');
        // Escape cells containing delimiter, quotes, or newlines
        if (str.includes(delimiter) || str.includes('"') || str.includes('\n')) {
          return `"${str.replace(/"/g, '""')}"`;
        }
        return str;
      }).join(delimiter) + '\n';
      
      // Rest of the data
      for (let i = 1; i < data.length; i++) {
        csvContent += data[i].map((cell: any) => {
          const str = String(cell || '');
          if (str.includes(delimiter) || str.includes('"') || str.includes('\n')) {
            return `"${str.replace(/"/g, '""')}"`;
          }
          return str;
        }).join(delimiter) + '\n';
      }
    } else {
      // All rows as data
      csvContent = data.map((row: any[]) =>
        row.map((cell: any) => {
          const str = String(cell || '');
          if (str.includes(delimiter) || str.includes('"') || str.includes('\n')) {
            return `"${str.replace(/"/g, '""')}"`;
          }
          return str;
        }).join(delimiter)
      ).join('\n');
    }

    // Save to file or return content
    if (outputPath) {
      writeFileSync(outputPath, csvContent, 'utf-8');
      return {
        sheet: selectedSheetName,
        totalRows,
        totalColumns,
        outputPath,
      };
    } else {
      return {
        sheet: selectedSheetName,
        totalRows,
        totalColumns,
        csvContent,
      };
    }
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      `Error exporting to CSV: ${error instanceof Error ? error.message : String(error)}`
    );
  }
}
