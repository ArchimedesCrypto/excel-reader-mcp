// Get statistics for numeric columns

import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import { GetColumnStatsArgs, GetColumnStatsResponse, ColumnStats } from '../../types/index.js';
import { loadWorkbook } from '../../utils/workbook-cache.js';
import { existsSync } from 'fs';
import * as XLSX from 'xlsx';

export function validateGetColumnStatsArgs(args: any): args is GetColumnStatsArgs {
  return (
    typeof args === 'object' &&
    args !== null &&
    typeof args.filePath === 'string' &&
    (args.sheetName === undefined || typeof args.sheetName === 'string') &&
    (args.columns === undefined || Array.isArray(args.columns))
  );
}

function calculateMedian(values: number[]): number {
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0
    ? (sorted[mid - 1] + sorted[mid]) / 2
    : sorted[mid];
}

function calculateStandardDeviation(values: number[], mean: number): number {
  const squareDiffs = values.map((value) => Math.pow(value - mean, 2));
  const avgSquareDiff = squareDiffs.reduce((a, b) => a + b, 0) / values.length;
  return Math.sqrt(avgSquareDiff);
}

export function getColumnStats(args: GetColumnStatsArgs): GetColumnStatsResponse {
  const { filePath, sheetName, columns } = args;

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

    const totalRows = allData.length;
    const allColumns = allData.length > 0 ? Object.keys(allData[0]) : [];
    const columnsToAnalyze = columns || allColumns;

    const columnStats: ColumnStats[] = columnsToAnalyze.map((column) => {
      const values = allData.map((row) => row[column]);
      const emptyCount = values.filter((v) => v === null || v === undefined || v === '').length;
      const nonEmptyValues = values.filter((v) => v !== null && v !== undefined && v !== '');

      // Determine data type
      const types = new Set(nonEmptyValues.map((v) => typeof v));
      let dataType = 'mixed';
      if (types.size === 1) {
        const type = Array.from(types)[0];
        if (type === 'number') dataType = 'number';
        else if (type === 'string') dataType = 'string';
      }

      const stats: ColumnStats = {
        column,
        header: column,
        dataType,
        totalValues: values.length,
        emptyCount,
      };

      // Calculate numeric stats
      if (dataType === 'number' || dataType === 'mixed') {
        const numericValues = nonEmptyValues
          .map((v) => Number(v))
          .filter((v) => !isNaN(v));

        if (numericValues.length > 0) {
          const sum = numericValues.reduce((a, b) => a + b, 0);
          const average = sum / numericValues.length;
          const min = Math.min(...numericValues);
          const max = Math.max(...numericValues);
          const median = calculateMedian(numericValues);
          const standardDeviation = calculateStandardDeviation(numericValues, average);

          stats.numericStats = {
            sum,
            average,
            min,
            max,
            median,
            standardDeviation,
          };

          if (dataType === 'mixed' && numericValues.length > 0) {
            dataType = 'number';
            stats.dataType = 'number';
          }
        }
      }

      // Calculate string stats
      if (dataType === 'string' || dataType === 'mixed') {
        const stringValues = nonEmptyValues.map((v) => String(v));

        if (stringValues.length > 0) {
          const lengths = stringValues.map((s) => s.length);
          const minLength = Math.min(...lengths);
          const maxLength = Math.max(...lengths);
          const avgLength = lengths.reduce((a, b) => a + b, 0) / lengths.length;

          stats.stringStats = {
            minLength,
            maxLength,
            avgLength,
          };

          if (dataType === 'mixed' && !stats.numericStats) {
            dataType = 'string';
            stats.dataType = 'string';
          }
        }
      }

      return stats;
    });

    return {
      sheet: selectedSheetName,
      totalRows,
      columns: columnStats,
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      `Error calculating column stats: ${error instanceof Error ? error.message : String(error)}`
    );
  }
}
