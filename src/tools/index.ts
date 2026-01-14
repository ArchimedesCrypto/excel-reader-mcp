// Central tool registry

import { listSheets, validateListSheetsArgs } from './read/list-sheets.js';
import { getCell, validateGetCellArgs } from './read/get-cell.js';
import { getRange, validateGetRangeArgs } from './read/get-range.js';
import { searchText, validateSearchTextArgs } from './read/search-text.js';
import { getSheetInfo, validateGetSheetInfoArgs } from './read/get-sheet-info.js';
import { getFormulas, validateGetFormulasArgs } from './read/get-formulas.js';
import { getMergedCells, validateGetMergedCellsArgs } from './read/get-merged-cells.js';
import { filterRows, validateFilterRowsArgs } from './analysis/filter-rows.js';
import { sortData, validateSortDataArgs } from './analysis/sort-data.js';
import { getUniqueValues, validateGetUniqueValuesArgs } from './analysis/get-unique-values.js';
import { getColumnStats, validateGetColumnStatsArgs } from './analysis/get-column-stats.js';
import { exportToCsv, validateExportToCsvArgs } from './export/export-to-csv.js';
import { exportToJson, validateExportToJsonArgs } from './export/export-to-json.js';

export interface ToolHandler {
  name: string;
  description: string;
  inputSchema: any;
  handler: (args: any) => any;
  validator: (args: any) => boolean;
}

export const toolHandlers: ToolHandler[] = [
  {
    name: 'list_sheets',
    description: 'List all sheet names in an Excel workbook',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: {
          type: 'string',
          description: 'Path to the Excel file',
        },
      },
      required: ['filePath'],
    },
    handler: listSheets,
    validator: validateListSheetsArgs,
  },
  {
    name: 'get_cell',
    description: 'Get the value of a specific cell',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: {
          type: 'string',
          description: 'Path to the Excel file',
        },
        sheetName: {
          type: 'string',
          description: 'Name of the sheet (optional, defaults to first sheet)',
        },
        cell: {
          type: 'string',
          description: 'Cell address (e.g., "A1", "B5")',
        },
      },
      required: ['filePath', 'cell'],
    },
    handler: getCell,
    validator: validateGetCellArgs,
  },
  {
    name: 'get_range',
    description: 'Get values from a specific cell range',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: {
          type: 'string',
          description: 'Path to the Excel file',
        },
        sheetName: {
          type: 'string',
          description: 'Name of the sheet (optional)',
        },
        range: {
          type: 'string',
          description: 'Cell range (e.g., "A1:C10")',
        },
        includeHeaders: {
          type: 'boolean',
          description: 'Treat first row as headers (optional)',
        },
      },
      required: ['filePath', 'range'],
    },
    handler: getRange,
    validator: validateGetRangeArgs,
  },
  {
    name: 'search_text',
    description: 'Find cells containing specific text',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: {
          type: 'string',
          description: 'Path to the Excel file',
        },
        sheetName: {
          type: 'string',
          description: 'Name of the sheet (optional, searches all sheets if not provided)',
        },
        searchText: {
          type: 'string',
          description: 'Text to search for',
        },
        matchCase: {
          type: 'boolean',
          description: 'Case-sensitive search (optional)',
        },
        matchWholeCell: {
          type: 'boolean',
          description: 'Match entire cell content (optional)',
        },
        maxResults: {
          type: 'number',
          description: 'Maximum number of results (optional, default 100)',
        },
      },
      required: ['filePath', 'searchText'],
    },
    handler: searchText,
    validator: validateSearchTextArgs,
  },
  {
    name: 'get_sheet_info',
    description: 'Get detailed metadata about a sheet',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: {
          type: 'string',
          description: 'Path to the Excel file',
        },
        sheetName: {
          type: 'string',
          description: 'Name of the sheet (optional)',
        },
      },
      required: ['filePath'],
    },
    handler: getSheetInfo,
    validator: validateGetSheetInfoArgs,
  },
  {
    name: 'get_formulas',
    description: 'Get formulas from cells in a range or entire sheet',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: {
          type: 'string',
          description: 'Path to the Excel file',
        },
        sheetName: {
          type: 'string',
          description: 'Name of the sheet (optional)',
        },
        range: {
          type: 'string',
          description: 'Cell range to search (optional, searches entire sheet if not provided)',
        },
      },
      required: ['filePath'],
    },
    handler: getFormulas,
    validator: validateGetFormulasArgs,
  },
  {
    name: 'get_merged_cells',
    description: 'Get all merged cell ranges in a sheet',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: {
          type: 'string',
          description: 'Path to the Excel file',
        },
        sheetName: {
          type: 'string',
          description: 'Name of the sheet (optional)',
        },
      },
      required: ['filePath'],
    },
    handler: getMergedCells,
    validator: validateGetMergedCellsArgs,
  },
  {
    name: 'filter_rows',
    description: 'Filter rows based on column criteria',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: {
          type: 'string',
          description: 'Path to the Excel file',
        },
        sheetName: {
          type: 'string',
          description: 'Name of the sheet (optional)',
        },
        conditions: {
          type: 'array',
          description: 'Array of filter conditions',
          items: {
            type: 'object',
            properties: {
              column: { type: 'string' },
              operator: { 
                type: 'string',
                enum: ['equals', 'not_equals', 'contains', 'not_contains', 'starts_with', 
                       'ends_with', 'greater_than', 'less_than', 'greater_equal', 
                       'less_equal', 'is_empty', 'is_not_empty']
              },
              value: { description: 'Value to compare (not needed for is_empty/is_not_empty)' },
            },
            required: ['column', 'operator'],
          },
        },
        logic: {
          type: 'string',
          enum: ['AND', 'OR'],
          description: 'Logic for combining conditions (optional, default AND)',
        },
        startRow: {
          type: 'number',
          description: 'Starting row for pagination (optional)',
        },
        maxRows: {
          type: 'number',
          description: 'Maximum rows to return (optional)',
        },
      },
      required: ['filePath', 'conditions'],
    },
    handler: filterRows,
    validator: validateFilterRowsArgs,
  },
  {
    name: 'sort_data',
    description: 'Sort data by one or more columns',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: {
          type: 'string',
          description: 'Path to the Excel file',
        },
        sheetName: {
          type: 'string',
          description: 'Name of the sheet (optional)',
        },
        sortBy: {
          type: 'array',
          description: 'Array of sort criteria',
          items: {
            type: 'object',
            properties: {
              column: { type: 'string' },
              order: { type: 'string', enum: ['asc', 'desc'] },
            },
            required: ['column', 'order'],
          },
        },
        hasHeaders: {
          type: 'boolean',
          description: 'Whether first row contains headers (optional, default true)',
        },
        startRow: {
          type: 'number',
          description: 'Starting row for pagination (optional)',
        },
        maxRows: {
          type: 'number',
          description: 'Maximum rows to return (optional)',
        },
      },
      required: ['filePath', 'sortBy'],
    },
    handler: sortData,
    validator: validateSortDataArgs,
  },
  {
    name: 'get_unique_values',
    description: 'Get distinct values in a column',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: {
          type: 'string',
          description: 'Path to the Excel file',
        },
        sheetName: {
          type: 'string',
          description: 'Name of the sheet (optional)',
        },
        column: {
          type: 'string',
          description: 'Column name or letter',
        },
        includeCount: {
          type: 'boolean',
          description: 'Include occurrence count (optional)',
        },
        sortBy: {
          type: 'string',
          enum: ['value', 'count'],
          description: 'Sort by value or count (optional)',
        },
        sortOrder: {
          type: 'string',
          enum: ['asc', 'desc'],
          description: 'Sort order (optional)',
        },
      },
      required: ['filePath', 'column'],
    },
    handler: getUniqueValues,
    validator: validateGetUniqueValuesArgs,
  },
  {
    name: 'get_column_stats',
    description: 'Get statistics for numeric columns',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: {
          type: 'string',
          description: 'Path to the Excel file',
        },
        sheetName: {
          type: 'string',
          description: 'Name of the sheet (optional)',
        },
        columns: {
          type: 'array',
          items: { type: 'string' },
          description: 'Columns to analyze (optional, analyzes all if not provided)',
        },
      },
      required: ['filePath'],
    },
    handler: getColumnStats,
    validator: validateGetColumnStatsArgs,
  },
  {
    name: 'export_to_csv',
    description: 'Export a sheet to CSV format',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: {
          type: 'string',
          description: 'Path to the Excel file',
        },
        sheetName: {
          type: 'string',
          description: 'Name of the sheet (optional)',
        },
        outputPath: {
          type: 'string',
          description: 'Path to save CSV file (optional, returns content if not provided)',
        },
        delimiter: {
          type: 'string',
          description: 'CSV delimiter (optional, default comma)',
        },
        includeHeaders: {
          type: 'boolean',
          description: 'Include header row (optional, default true)',
        },
        dateFormat: {
          type: 'string',
          description: 'Date format (optional, default ISO 8601)',
        },
      },
      required: ['filePath'],
    },
    handler: exportToCsv,
    validator: validateExportToCsvArgs,
  },
  {
    name: 'export_to_json',
    description: 'Export a sheet to JSON format',
    inputSchema: {
      type: 'object',
      properties: {
        filePath: {
          type: 'string',
          description: 'Path to the Excel file',
        },
        sheetName: {
          type: 'string',
          description: 'Name of the sheet (optional)',
        },
        outputPath: {
          type: 'string',
          description: 'Path to save JSON file (optional, returns content if not provided)',
        },
        format: {
          type: 'string',
          enum: ['array', 'objects'],
          description: 'Format as array of arrays or array of objects (optional, default objects)',
        },
        includeMetadata: {
          type: 'boolean',
          description: 'Include metadata in output (optional)',
        },
        dateFormat: {
          type: 'string',
          description: 'Date format (optional, default ISO 8601)',
        },
      },
      required: ['filePath'],
    },
    handler: exportToJson,
    validator: validateExportToJsonArgs,
  },
];

export function getToolByName(name: string): ToolHandler | undefined {
  return toolHandlers.find((tool) => tool.name === name);
}
