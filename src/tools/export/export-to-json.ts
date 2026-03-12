// Export a sheet to JSON format

import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import { ExportToJsonArgs, ExportToJsonResponse } from '../../types/index.js';
import { loadWorkbook } from '../../utils/workbook-cache.js';
import { existsSync, writeFileSync } from 'fs';
import * as XLSX from 'xlsx';

export function validateExportToJsonArgs(args: any): args is ExportToJsonArgs {
  return (
    typeof args === 'object' &&
    args !== null &&
    typeof args.filePath === 'string' &&
    (args.sheetName === undefined || typeof args.sheetName === 'string') &&
    (args.outputPath === undefined || typeof args.outputPath === 'string') &&
    (args.format === undefined || args.format === 'array' || args.format === 'objects') &&
    (args.includeMetadata === undefined || typeof args.includeMetadata === 'boolean') &&
    (args.dateFormat === undefined || typeof args.dateFormat === 'string')
  );
}

export function exportToJson(args: ExportToJsonArgs): ExportToJsonResponse {
  const {
    filePath,
    sheetName,
    outputPath,
    format = 'objects',
    includeMetadata = false,
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

    let jsonData: any;

    if (format === 'array') {
      // Export as array of arrays
      jsonData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        raw: false,
        dateNF: dateFormat
      });
    } else {
      // Export as array of objects
      jsonData = XLSX.utils.sheet_to_json(worksheet, {
        raw: false,
        dateNF: dateFormat
      });
    }

    const totalRows = Array.isArray(jsonData) ? jsonData.length : 0;

    // Add metadata if requested
    let result: any = jsonData;
    if (includeMetadata) {
      const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
      result = {
        metadata: {
          sheet: selectedSheetName,
          totalRows,
          totalColumns: range.e.c - range.s.c + 1,
          dateFormat,
          format,
        },
        data: jsonData,
      };
    }

    // Save to file or return content
    const jsonContent = JSON.stringify(result, null, 2);

    if (outputPath) {
      writeFileSync(outputPath, jsonContent, 'utf-8');
      return {
        sheet: selectedSheetName,
        totalRows,
        outputPath,
      };
    } else {
      return {
        sheet: selectedSheetName,
        totalRows,
        jsonContent: result,
      };
    }
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      `Error exporting to JSON: ${error instanceof Error ? error.message : String(error)}`
    );
  }
}
