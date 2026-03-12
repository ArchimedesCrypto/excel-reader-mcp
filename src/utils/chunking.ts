// Chunking utilities for handling large datasets

import { ChunkingOptions, ChunkedResult } from '../types/index.js';

export const MAX_RESPONSE_SIZE = 100 * 1024; // 100KB default max response size

/**
 * Estimate size of stringified JSON
 */
export function estimateJsonSize(obj: any): number {
  const str = JSON.stringify(obj);
  return str.length * 2; // Rough estimate, multiply by 2 for unicode
}

/**
 * Calculate optimal chunk size based on sample data
 */
export function calculateChunkSize(data: any[], maxSize: number = MAX_RESPONSE_SIZE): number {
  if (data.length === 0) {
    return 100; // Default if no data
  }
  
  const singleRowSize = estimateJsonSize(data[0]);
  if (singleRowSize === 0) {
    return 100;
  }
  
  return Math.max(1, Math.floor(maxSize / singleRowSize));
}

/**
 * Apply chunking to an array of data
 */
export function applyChunking<T>(
  data: T[],
  options: ChunkingOptions = {}
): ChunkedResult<T> {
  const {
    maxResponseSize = MAX_RESPONSE_SIZE,
    maxRows,
    startRow = 0,
  } = options;
  
  const totalItems = data.length;
  
  // Calculate effective max rows
  let effectiveMaxRows = maxRows;
  if (!effectiveMaxRows && data.length > 0) {
    effectiveMaxRows = calculateChunkSize(data.slice(0, 100), maxResponseSize);
  } else if (!effectiveMaxRows) {
    effectiveMaxRows = 100;
  }
  
  const endRow = Math.min(startRow + effectiveMaxRows, totalItems);
  const chunkedData = data.slice(startRow, endRow);
  const hasMore = endRow < totalItems;
  
  return {
    data: chunkedData,
    totalItems,
    hasMore,
    nextChunk: hasMore ? { startRow: endRow } : undefined,
  };
}

/**
 * Check if response would exceed size limit
 */
export function wouldExceedSizeLimit(obj: any, maxSize: number = MAX_RESPONSE_SIZE): boolean {
  return estimateJsonSize(obj) > maxSize;
}

/**
 * Split data into chunks that fit within size limits
 */
export function splitIntoChunks<T>(
  data: T[],
  maxSize: number = MAX_RESPONSE_SIZE
): T[][] {
  if (data.length === 0) {
    return [[]];
  }
  
  const chunks: T[][] = [];
  let currentChunk: T[] = [];
  let currentSize = 0;
  
  for (const item of data) {
    const itemSize = estimateJsonSize(item);
    
    if (currentSize + itemSize > maxSize && currentChunk.length > 0) {
      chunks.push(currentChunk);
      currentChunk = [item];
      currentSize = itemSize;
    } else {
      currentChunk.push(item);
      currentSize += itemSize;
    }
  }
  
  if (currentChunk.length > 0) {
    chunks.push(currentChunk);
  }
  
  return chunks;
}
