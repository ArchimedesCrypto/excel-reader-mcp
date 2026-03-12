// Cell address parsing and manipulation utilities

export interface CellAddress {
  column: string;  // Letter: A, B, AA, etc.
  row: number;     // 1-based row number
  colIndex: number; // 0-based column index
}

export interface RangeAddress {
  start: CellAddress;
  end: CellAddress;
}

/**
 * Convert column letter to 0-based index
 * A -> 0, B -> 1, Z -> 25, AA -> 26, etc.
 */
export function columnLetterToIndex(letter: string): number {
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return index - 1;
}

/**
 * Convert 0-based index to column letter
 * 0 -> A, 1 -> B, 25 -> Z, 26 -> AA, etc.
 */
export function indexToColumnLetter(index: number): string {
  let letter = '';
  let num = index + 1;
  
  while (num > 0) {
    const remainder = (num - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    num = Math.floor((num - 1) / 26);
  }
  
  return letter;
}

/**
 * Parse a cell address like "A1" or "AB25"
 */
export function parseCellAddress(cell: string): CellAddress {
  const match = cell.match(/^([A-Z]+)(\d+)$/);
  
  if (!match) {
    throw new Error(`Invalid cell address: ${cell}`);
  }
  
  const column = match[1];
  const row = parseInt(match[2], 10);
  const colIndex = columnLetterToIndex(column);
  
  return { column, row, colIndex };
}

/**
 * Parse a range like "A1:C10"
 */
export function parseRange(range: string): RangeAddress {
  const match = range.match(/^([A-Z]+\d+):([A-Z]+\d+)$/);
  
  if (!match) {
    throw new Error(`Invalid range: ${range}`);
  }
  
  const start = parseCellAddress(match[1]);
  const end = parseCellAddress(match[2]);
  
  // Ensure start is before end
  if (start.row > end.row || start.colIndex > end.colIndex) {
    throw new Error(`Invalid range: start must be before end in ${range}`);
  }
  
  return { start, end };
}

/**
 * Validate cell address format
 */
export function isValidCellAddress(cell: string): boolean {
  return /^[A-Z]+\d+$/.test(cell);
}

/**
 * Validate range format
 */
export function isValidRange(range: string): boolean {
  return /^[A-Z]+\d+:[A-Z]+\d+$/.test(range);
}

/**
 * Create a cell address from column and row
 */
export function createCellAddress(column: string | number, row: number): string {
  const colLetter = typeof column === 'string' ? column : indexToColumnLetter(column);
  return `${colLetter}${row}`;
}

/**
 * Create a range address from coordinates
 */
export function createRange(
  startCol: string | number,
  startRow: number,
  endCol: string | number,
  endRow: number
): string {
  const startCell = createCellAddress(startCol, startRow);
  const endCell = createCellAddress(endCol, endRow);
  return `${startCell}:${endCell}`;
}

/**
 * Get the dimensions of a range
 */
export function getRangeDimensions(range: string): {
  startRow: number;
  endRow: number;
  startCol: string;
  endCol: string;
  rowCount: number;
  colCount: number;
} {
  const { start, end } = parseRange(range);
  
  return {
    startRow: start.row,
    endRow: end.row,
    startCol: start.column,
    endCol: end.column,
    rowCount: end.row - start.row + 1,
    colCount: end.colIndex - start.colIndex + 1,
  };
}
