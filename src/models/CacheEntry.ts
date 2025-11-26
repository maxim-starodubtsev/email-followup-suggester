export interface CacheEntry<T> {
  data: T;
  timestamp: number;
  ttl: number;
  accessCount: number;
  lastAccessed: number;
  size: number; // Estimated memory size in bytes
  hash: string; // Content hash for integrity checking
}

export interface CacheStats {
  totalEntries: number;
  totalMemoryUsage: number; // bytes
  hitRate: number; // percentage
  totalHits: number;
  totalMisses: number;
  totalEvictions: number;
  oldestEntry?: number; // timestamp
  newestEntry?: number; // timestamp
  averageAccessCount: number;
}

export interface CacheOptions {
  defaultTtl: number; // milliseconds
  maxMemoryUsage: number; // bytes
  maxEntries: number;
  evictionPolicy: "lru" | "lfu" | "fifo" | "ttl";
  cleanupInterval: number; // milliseconds
  enableStatistics: boolean;
  enableContentHashing: boolean;
}

export interface CacheEvictionResult {
  evictedCount: number;
  freedMemory: number;
  reason: "ttl" | "memory_pressure" | "max_entries" | "manual";
}
