#!/usr/bin/env node
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ErrorCode,
  ListToolsRequestSchema,
  McpError,
} from '@modelcontextprotocol/sdk/types.js';
import * as XLSX from 'xlsx';
import { existsSync, readFileSync } from 'fs';

interface ExcelChunk {
  rowStart: number;
  rowEnd: number;
  columns: string[];
  data: Record<string, any>[];
}

interface ExcelSheetData {
  name: string;
  totalRows: number;
  totalColumns: number;
  chunk: ExcelChunk;
  hasMore: boolean;
  nextChunk?: {
    rowStart: number;
    columns: string[];
  };
}

interface ExcelData {
  fileName: string;
  totalSheets: number;
  currentSheet: ExcelSheetData;
}

interface ReadExcelArgs {
  filePath: string;
  sheetName?: string;
  startRow?: number;
  maxRows?: number;
}

const MAX_RESPONSE_SIZE = 100 * 1024; // 100KB default max response size

const isValidReadExcelArgs = (args: any): args is ReadExcelArgs =>
  typeof args === 'object' &&
  args !== null &&
  typeof args.filePath === 'string' &&
  (args.sheetName === undefined || typeof args.sheetName === 'string') &&
  (args.startRow === undefined || typeof args.startRow === 'number') &&
  (args.maxRows === undefined || typeof args.maxRows === 'number');

// Estimate size of stringified JSON
const estimateJsonSize = (obj: any): number => {
  const str = JSON.stringify(obj);
  return str.length * 2; // Rough estimate, multiply by 2 for unicode
};

// Calculate optimal chunk size
const calculateChunkSize = (data: any[], maxSize: number): number => {
  const singleRowSize = estimateJsonSize(data[0]);
  return Math.max(1, Math.floor(maxSize / singleRowSize));
};

class ExcelReaderServer {
  private server: Server;

  constructor() {
    this.server = new Server(
      {
        name: 'excel-reader',
        version: '1.0.0',
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

    this.setupToolHandlers();
    
    // Error handling
    this.server.onerror = (error) => console.error('[MCP Error]', error);
    process.on('SIGINT', async () => {
      await this.server.close();
      process.exit(0);
    });
  }

  private readExcelFile(args: ReadExcelArgs): ExcelData {
    const { filePath, sheetName, startRow = 0, maxRows } = args;
    if (!existsSync(filePath)) {
      throw new McpError(
        ErrorCode.InvalidRequest,
        `File not found: ${filePath}`
      );
    }

    try {
      // Read file as buffer first
      const data = readFileSync(filePath);
      const workbook = XLSX.read(data, {
        type: 'buffer',
        cellDates: true,
        cellNF: false,
        cellText: false,
        dateNF: 'yyyy-mm-dd'
      });
      const fileName = filePath.split(/[\\/]/).pop() || '';
      const selectedSheetName = sheetName || workbook.SheetNames[0];
      const worksheet = workbook.Sheets[selectedSheetName];
      const allData = XLSX.utils.sheet_to_json(worksheet, {
        raw: true,
        dateNF: 'yyyy-mm-dd'
      }) as Record<string, any>[];

      const totalRows = allData.length;
      const columns = totalRows > 0 ? Object.keys(allData[0] as object) : [];
      const totalColumns = columns.length;

      // Calculate chunk size based on data size
      let effectiveMaxRows = maxRows;
      if (!effectiveMaxRows) {
        const initialChunk = allData.slice(0, 100); // Sample first 100 rows
        if (initialChunk.length > 0) {
          effectiveMaxRows = calculateChunkSize(initialChunk, MAX_RESPONSE_SIZE);
        } else {
          effectiveMaxRows = 100; // Default if no data
        }
      }

      const endRow = Math.min(startRow + effectiveMaxRows, totalRows);
      const chunkData = allData.slice(startRow, endRow);
      
      const hasMore = endRow < totalRows;
      const nextChunk = hasMore ? {
        rowStart: endRow,
        columns
      } : undefined;

      return {
        fileName,
        totalSheets: workbook.SheetNames.length,
        currentSheet: {
          name: selectedSheetName,
          totalRows,
          totalColumns,
          chunk: {
            rowStart: startRow,
            rowEnd: endRow,
            columns,
            data: chunkData
          },
          hasMore,
          nextChunk
        }
      };
    } catch (error) {
      throw new McpError(
        ErrorCode.InternalError,
        `Error reading Excel file: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  private setupToolHandlers() {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: [
        {
          name: 'read_excel',
          description: 'Read an Excel file and return its contents as structured data',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: {
                type: 'string',
                description: 'Path to the Excel file to read',
              },
              sheetName: {
                type: 'string',
                description: 'Name of the sheet to read (optional)',
              },
              startRow: {
                type: 'number',
                description: 'Starting row index (optional)',
              },
              maxRows: {
                type: 'number',
                description: 'Maximum number of rows to read (optional)',
              },
            },
            required: ['filePath'],
          },
        },
      ],
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      if (request.params.name !== 'read_excel') {
        throw new McpError(
          ErrorCode.MethodNotFound,
          `Unknown tool: ${request.params.name}`
        );
      }

      if (!isValidReadExcelArgs(request.params.arguments)) {
        throw new McpError(
          ErrorCode.InvalidParams,
          'Invalid read_excel arguments'
        );
      }

      try {
        const data = this.readExcelFile(request.params.arguments);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(data, null, 2),
            },
          ],
        };
      } catch (error) {
        if (error instanceof McpError) {
          throw error;
        }
        throw new McpError(
          ErrorCode.InternalError,
          `Unexpected error: ${error instanceof Error ? error.message : String(error)}`
        );
      }
    });
  }

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error('Excel Reader MCP server running on stdio');
  }
}

const server = new ExcelReaderServer();
server.run().catch(console.error);
