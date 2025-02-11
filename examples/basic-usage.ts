import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';

// Define the response type for Excel data
interface ExcelResponse {
  fileName: string;
  totalSheets: number;
  currentSheet: {
    name: string;
    totalRows: number;
    totalColumns: number;
    chunk: {
      rowStart: number;
      rowEnd: number;
      columns: string[];
      data: Record<string, any>[];
    };
    hasMore: boolean;
    nextChunk?: {
      rowStart: number;
      columns: string[];
    };
  };
}

async function main() {
  // Create MCP client
  const transport = new StdioClientTransport();
  const client = new Client({
    name: "excel-reader-example",
    version: "1.0.0"
  });
  await client.connect(transport);

  // Example 1: Basic Usage
  const basicExample = async () => {
    const result = await client.request("call_tool", {
      method: "call_tool",
      params: {
        name: "read_excel",
        arguments: {
          filePath: "./data/sample.xlsx"
        }
      }
    });
    console.log("Basic read result:", result);
  };

  // Example 2: Pagination with chunks
  const paginationExample = async () => {
    // Read first chunk
    const chunk1 = await client.request("call_tool", {
      method: "call_tool",
      params: {
        name: "read_excel",
        arguments: {
          filePath: "./data/large-file.xlsx",
          maxRows: 100
        }
      }
    });
    console.log("First chunk:", chunk1);

    // Read next chunk if available
    const chunk1Response = chunk1 as { content: Array<{ type: string; text: string }> };
    if (chunk1Response.content?.[0]?.text) {
      const data = JSON.parse(chunk1Response.content[0].text) as ExcelResponse;
      if (data.currentSheet.hasMore && data.currentSheet.nextChunk) {
        const chunk2 = await client.request("call_tool", {
          method: "call_tool",
          params: {
            name: "read_excel",
            arguments: {
              filePath: "./data/large-file.xlsx",
              startRow: data.currentSheet.nextChunk.rowStart,
              maxRows: 100
            }
          }
        });
        console.log("Second chunk:", chunk2);
      }
    }
  };

  // Example 3: Specific sheet selection
  const sheetSelectionExample = async () => {
    const result = await client.request("call_tool", {
      method: "call_tool",
      params: {
        name: "read_excel",
        arguments: {
          filePath: "./data/multi-sheet.xlsx",
          sheetName: "Sheet2"
        }
      }
    });
    console.log("Sheet selection result:", result);
  };

  // Example 4: Error handling
  const errorHandlingExample = async () => {
    try {
      await client.request("call_tool", {
        method: "call_tool",
        params: {
          name: "read_excel",
          arguments: {
            filePath: "./data/non-existent.xlsx"
          }
        }
      });
    } catch (error) {
      console.error("Error reading file:", error);
    }
  };

  // Run examples
  console.log("Running Excel Reader examples...\n");
  
  console.log("1. Basic Usage Example:");
  await basicExample();
  
  console.log("\n2. Pagination Example:");
  await paginationExample();
  
  console.log("\n3. Sheet Selection Example:");
  await sheetSelectionExample();
  
  console.log("\n4. Error Handling Example:");
  await errorHandlingExample();
}

main().catch(console.error);
