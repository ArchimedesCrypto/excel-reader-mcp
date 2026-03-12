// Find cells containing specific text

import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import { SearchTextArgs, SearchTextResponse, SearchResult } from '../../types/index.js';
import { loadWorkbook } from '../../utils/workbook-cache.js';
import { existsSync } from 'fs';
import * as XLSX from 'xlsx';
import { indexToColumnLetter } from '../../utils/cell-utils.js';

export function validateSearchTextArgs(args: any): args is SearchTextArgs {
  return (
    typeof args === 'object' &&
    args !== null &&
    typeof args.filePath === 'string' &&
    typeof args.searchText === 'string' &&
    (args.sheetName === undefined || typeof args.sheetName === 'string') &&
    (args.matchCase === undefined || typeof args.matchCase === 'boolean') &&
    (args.matchWholeCell === undefined || typeof args.matchWholeCell === 'boolean') &&
    (args.maxResults === undefined || typeof args.maxResults === 'number')
  );
}

export function searchText(args: SearchTextArgs): SearchTextResponse {
  const {
    filePath,
    sheetName,
    searchText,
    matchCase = false,
    matchWholeCell = false,
    maxResults = 100,
  } = args;

  if (!existsSync(filePath)) {
    throw new McpError(
      ErrorCode.InvalidRequest,
      `File not found: ${filePath}`
    );
  }

  try {
    const workbook = loadWorkbook(filePath);
    const results: SearchResult[] = [];
    let totalMatches = 0;

    // Determine which sheets to search
    const sheetsToSearch = sheetName ? [sheetName] : workbook.SheetNames;

    // Prepare search text
    const searchStr = matchCase ? searchText : searchText.toLowerCase();

    for (const sheet of sheetsToSearch) {
      const worksheet = workbook.Sheets[sheet];

      if (!worksheet) {
        if (sheetName) {
          throw new McpError(
            ErrorCode.InvalidRequest,
            `Sheet not found: ${sheetName}`
          );
        }
        continue;
      }

      // Iterate through all cells in the sheet
      for (const cellAddr in worksheet) {
        if (cellAddr[0] === '!') continue; // Skip special properties

        const cell = worksheet[cellAddr];
        if (!cell || cell.v === undefined || cell.v === null) continue;

        // Convert cell value to string for comparison
        const cellValue = String(cell.v);
        const compareValue = matchCase ? cellValue : cellValue.toLowerCase();

        let isMatch = false;
        if (matchWholeCell) {
          isMatch = compareValue === searchStr;
        } else {
          isMatch = compareValue.includes(searchStr);
        }

        if (isMatch) {
          totalMatches++;

          if (results.length < maxResults) {
            // Parse cell address
            const decoded = XLSX.utils.decode_cell(cellAddr);
            const column = indexToColumnLetter(decoded.c);
            const row = decoded.r + 1; // Convert to 1-based

            results.push({
              sheet,
              cell: cellAddr,
              value: cellValue,
              row,
              column,
            });
          }
        }
      }
    }

    return {
      query: searchText,
      totalMatches,
      results,
      hasMore: totalMatches > maxResults,
    };
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    throw new McpError(
      ErrorCode.InternalError,
      `Error searching text: ${error instanceof Error ? error.message : String(error)}`
    );
  }
}
