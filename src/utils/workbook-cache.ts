// Workbook caching for performance optimization

import * as XLSX from 'xlsx';
import { existsSync, statSync, readFileSync } from 'fs';

interface CacheEntry {
  filePath: string;
  workbook: XLSX.WorkBook;
  lastAccessed: number;
  fileModTime: number;
}

class WorkbookCache {
  private cache: Map<string, CacheEntry>;
  private maxEntries: number = 5;
  private maxAge: number = 60000; // 1 minute

  constructor(maxEntries?: number, maxAge?: number) {
    this.cache = new Map();
    if (maxEntries !== undefined) {
      this.maxEntries = maxEntries;
    }
    if (maxAge !== undefined) {
      this.maxAge = maxAge;
    }
  }

  /**
   * Get workbook from cache or load it
   */
  get(filePath: string): XLSX.WorkBook | null {
    const entry = this.cache.get(filePath);
    
    if (!entry) {
      return null;
    }
    
    // Check if file has been modified
    if (existsSync(filePath)) {
      const stats = statSync(filePath);
      const currentModTime = stats.mtimeMs;
      
      if (currentModTime !== entry.fileModTime) {
        // File has been modified, invalidate cache
        this.cache.delete(filePath);
        return null;
      }
    } else {
      // File no longer exists, invalidate cache
      this.cache.delete(filePath);
      return null;
    }
    
    // Check if entry has expired
    const age = Date.now() - entry.lastAccessed;
    if (age > this.maxAge) {
      this.cache.delete(filePath);
      return null;
    }
    
    // Update last accessed time
    entry.lastAccessed = Date.now();
    
    return entry.workbook;
  }

  /**
   * Set workbook in cache
   */
  set(filePath: string, workbook: XLSX.WorkBook): void {
    // Check if we need to evict entries
    if (this.cache.size >= this.maxEntries) {
      this.evictOldest();
    }
    
    const stats = statSync(filePath);
    
    this.cache.set(filePath, {
      filePath,
      workbook,
      lastAccessed: Date.now(),
      fileModTime: stats.mtimeMs,
    });
  }

  /**
   * Invalidate a specific file from cache
   */
  invalidate(filePath: string): void {
    this.cache.delete(filePath);
  }

  /**
   * Clear all cache entries
   */
  clear(): void {
    this.cache.clear();
  }

  /**
   * Evict the oldest entry from cache
   */
  private evictOldest(): void {
    let oldestKey: string | null = null;
    let oldestTime = Infinity;
    
    for (const [key, entry] of this.cache.entries()) {
      if (entry.lastAccessed < oldestTime) {
        oldestTime = entry.lastAccessed;
        oldestKey = key;
      }
    }
    
    if (oldestKey) {
      this.cache.delete(oldestKey);
    }
  }

  /**
   * Get cache statistics
   */
  getStats(): {
    size: number;
    maxEntries: number;
    entries: Array<{ filePath: string; age: number }>;
  } {
    const now = Date.now();
    const entries = Array.from(this.cache.entries()).map(([filePath, entry]) => ({
      filePath,
      age: now - entry.lastAccessed,
    }));
    
    return {
      size: this.cache.size,
      maxEntries: this.maxEntries,
      entries,
    };
  }
}

// Global cache instance
const globalCache = new WorkbookCache();

/**
 * Load workbook with caching
 */
export function loadWorkbook(filePath: string, useCache: boolean = true): XLSX.WorkBook {
  if (useCache) {
    const cached = globalCache.get(filePath);
    if (cached) {
      return cached;
    }
  }
  
  const data = readFileSync(filePath);
  const workbook = XLSX.read(data, {
    type: 'buffer',
    cellDates: true,
    cellNF: false,
    cellText: false,
    dateNF: 'yyyy-mm-dd'
  });
  
  if (useCache) {
    globalCache.set(filePath, workbook);
  }
  
  return workbook;
}

/**
 * Invalidate cache for a specific file
 */
export function invalidateCache(filePath: string): void {
  globalCache.invalidate(filePath);
}

/**
 * Clear all cached workbooks
 */
export function clearCache(): void {
  globalCache.clear();
}

/**
 * Get cache statistics
 */
export function getCacheStats() {
  return globalCache.getStats();
}

export { WorkbookCache };
