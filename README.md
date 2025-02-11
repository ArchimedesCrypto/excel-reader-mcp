# MCP Excel Reader

A Model Context Protocol (MCP) server for reading Excel files with automatic chunking and pagination support. Built with SheetJS and TypeScript, this tool helps you handle large Excel files efficiently by automatically breaking them into manageable chunks.

## Features

- 📊 Read Excel files (.xlsx, .xls) with automatic size limits
- 🔄 Automatic chunking for large datasets
- 📑 Sheet selection and row pagination
- 📅 Proper date handling
- ⚡ Optimized for large files
- 🛡️ Error handling and validation

## Installation

### As an MCP Server

1. Install globally:
```bash
npm install -g @archimdescrypto/excel-reader
```

2. Add to your MCP settings file (usually at `~/.config/claude/settings.json` or equivalent):
```json
{
  "mcpServers": {
    "excel-reader": {
      "command": "excel-reader",
      "env": {}
    }
  }
}
```

### For Development

1. Clone the repository:
```bash
git clone https://github.com/ArchimdesCrypto/mcp-excel-reader.git
cd mcp-excel-reader
```

2. Install dependencies:
```bash
npm install
```

3. Build the project:
```bash
npm run build
```

## Usage

## Usage

The Excel Reader provides a single tool `read_excel` with the following parameters:

```typescript
interface ReadExcelArgs {
  filePath: string;      // Path to Excel file
  sheetName?: string;    // Optional sheet name (defaults to first sheet)
  startRow?: number;     // Optional starting row for pagination
  maxRows?: number;      // Optional maximum rows to read
}

// Response format
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
```

### Basic Usage

When used with Claude or another MCP-compatible AI:

```
Read the Excel file at path/to/file.xlsx
```

The AI will use the tool to read the file, automatically handling chunking for large files.

### Features

1. **Automatic Chunking**
   - Automatically splits large files into manageable chunks
   - Default chunk size of 100KB
   - Provides metadata for pagination

2. **Sheet Selection**
   - Read specific sheets by name
   - Defaults to first sheet if not specified

3. **Row Pagination**
   - Control which rows to read with startRow and maxRows
   - Get next chunk information for continuous reading

4. **Error Handling**
   - Validates file existence and format
   - Provides clear error messages
   - Handles malformed Excel files gracefully

## Extending with SheetJS Features

The Excel Reader is built on SheetJS and can be extended with its powerful features:

### Available Extensions

1. **Formula Handling**
   ```typescript
   // Enable formula parsing
   const wb = XLSX.read(data, {
     cellFormula: true,
     cellNF: true
   });
   ```

2. **Cell Formatting**
   ```typescript
   // Access cell styles and formatting
   const styles = Object.keys(worksheet)
     .filter(key => key[0] !== '!')
     .map(key => ({
       cell: key,
       style: worksheet[key].s
     }));
   ```

3. **Data Validation**
   ```typescript
   // Access data validation rules
   const validation = worksheet['!dataValidation'];
   ```

4. **Sheet Features**
   - Merged Cells: `worksheet['!merges']`
   - Hidden Rows/Columns: `worksheet['!rows']`, `worksheet['!cols']`
   - Sheet Protection: `worksheet['!protect']`

For more features and detailed documentation, visit the [SheetJS Documentation](https://docs.sheetjs.com/).

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Built with [SheetJS](https://sheetjs.com/)
- Part of the [Model Context Protocol](https://github.com/modelcontextprotocol/mcp) ecosystem
