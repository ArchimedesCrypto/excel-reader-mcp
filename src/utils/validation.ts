// Input validation utilities

import { existsSync } from 'fs';
import { isValidCellAddress, isValidRange } from './cell-utils.js';
import { FilterCondition } from '../types/index.js';

/**
 * Validate file path exists
 */
export function isValidFilePath(path: string): boolean {
  return typeof path === 'string' && path.length > 0 && existsSync(path);
}

/**
 * Validate column reference (either letter or header name)
 */
export function isValidColumnRef(col: string, headers?: string[]): boolean {
  if (typeof col !== 'string' || col.length === 0) {
    return false;
  }
  
  // Check if it's a column letter (A, B, AA, etc.)
  if (/^[A-Z]+$/.test(col)) {
    return true;
  }
  
  // Check if it's a valid header name
  if (headers && headers.includes(col)) {
    return true;
  }
  
  return false;
}

/**
 * Validate filter condition
 */
export function isValidFilterCondition(condition: FilterCondition): boolean {
  if (!condition || typeof condition !== 'object') {
    return false;
  }
  
  const { column, operator, value } = condition;
  
  // Check column
  if (typeof column !== 'string' || column.length === 0) {
    return false;
  }
  
  // Check operator
  const validOperators = [
    'equals', 'not_equals', 'contains', 'not_contains',
    'starts_with', 'ends_with', 'greater_than', 'less_than',
    'greater_equal', 'less_equal', 'is_empty', 'is_not_empty'
  ];
  
  if (!validOperators.includes(operator)) {
    return false;
  }
  
  // For operators that need a value, check it's provided
  const operatorsNeedingValue = [
    'equals', 'not_equals', 'contains', 'not_contains',
    'starts_with', 'ends_with', 'greater_than', 'less_than',
    'greater_equal', 'less_equal'
  ];
  
  if (operatorsNeedingValue.includes(operator) && value === undefined) {
    return false;
  }
  
  return true;
}

/**
 * Validate sheet name exists in workbook
 */
export function isValidSheetName(sheetName: string, availableSheets: string[]): boolean {
  return availableSheets.includes(sheetName);
}

export { isValidCellAddress, isValidRange };
