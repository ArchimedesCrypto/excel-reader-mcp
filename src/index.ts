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
import { toolHandlers, getToolByName } from './tools/index.js';
import { loadWorkbook } from './utils/workbook-cache.js';
import { calculateChunkSize, MAX_RESPONSE_SIZE } from './utils/chunking.js';

// Legacy types for read_excel tool
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

const isValidReadExcelArgs = (args: any): args is ReadExcelArgs =>
  typeof args === 'object' &&
  args !== null &&
  typeof args.filePath === 'string' &&
  (args.sheetName === undefined || typeof args.sheetName === 'string') &&
  (args.startRow === undefined || typeof args.startRow === 'number') &&
  (args.maxRows === undefined || typeof args.maxRows === 'number');

class ExcelReaderServer {
  private server: Server;

  constructor() {
    this.server = new Server(
      {
        name: 'excel-reader',
        version: '2.0.0',
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
      const workbook = loadWorkbook(filePath);
      const fileName = filePath.split(/[\\/]/).pop() || '';
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
    this.server.setRequestHandler(ListToolsRequestSchema, async () => {
      // Build tools list including the legacy read_excel tool
      const tools = [
        {
          name: 'read_excel',
          description: 'Read an Excel file and return its contents as structured data with automatic chunking',
          inputSchema: {
            type: 'object',
            properties: {
              filePath: {
                type: 'string',
                description: 'Path to the Excel file to read',
              },
              sheetName: {
                type: 'string',
                description: 'Name of the sheet to read (optional, defaults to first sheet)',
              },
              startRow: {
                type: 'number',
                description: 'Starting row index for pagination (optional, default 0)',
              },
              maxRows: {
                type: 'number',
                description: 'Maximum number of rows to read (optional, auto-calculated based on size)',
              },
            },
            required: ['filePath'],
          },
        },
        // Add all new tools from the registry
        ...toolHandlers.map((tool) => ({
          name: tool.name,
          description: tool.description,
          inputSchema: tool.inputSchema,
        })),
      ];

      return { tools };
    });

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const toolName = request.params.name;

      // Handle legacy read_excel tool
      if (toolName === 'read_excel') {
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
      }

      // Handle new tools from registry
      const tool = getToolByName(toolName);
      if (!tool) {
        throw new McpError(
          ErrorCode.MethodNotFound,
          `Unknown tool: ${toolName}`
        );
      }

      if (!tool.validator(request.params.arguments)) {
        throw new McpError(
          ErrorCode.InvalidParams,
          `Invalid arguments for ${toolName}`
        );
      }

      try {
        const result = tool.handler(request.params.arguments);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(result, null, 2),
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
