import { CacheEntry, CacheStats, CacheOptions, CacheEvictionResult } from '../models/CacheEntry';

export interface ICacheService {
    get<T>(key: string): T | null;
    set<T>(key: string, value: T, ttl?: number): void;
    has(key: string): boolean;
    invalidate(key: string): boolean;
    clear(): void;
    getStats(): CacheStats;
    cleanup(): CacheEvictionResult;
    updateOptions(options: Partial<CacheOptions>): void;
    bulkInvalidate(pattern: RegExp): number;
}

export class CacheService implements ICacheService {
    private cache = new Map<string, CacheEntry<any>>();
    private stats = {
        totalHits: 0,
        totalMisses: 0,
        totalEvictions: 0
    };
    
    private options: CacheOptions = {
        defaultTtl: 24 * 60 * 60 * 1000, // 24 hours
        maxMemoryUsage: 50 * 1024 * 1024, // 50MB
        maxEntries: 10000,
        evictionPolicy: 'lru',
        cleanupInterval: 5 * 60 * 1000, // 5 minutes
        enableStatistics: true,
        enableContentHashing: true
    };

    private cleanupTimer?: NodeJS.Timeout;

    constructor(options?: Partial<CacheOptions>) {
        if (options) {
            this.updateOptions(options);
        }
        this.startCleanupTimer();
    }

    public get<T>(key: string): T | null {
        const entry = this.cache.get(key);
        
        if (!entry) {
            this.stats.totalMisses++;
            return null;
        }

        const now = Date.now();
        
        // Check TTL expiration
        if (now - entry.timestamp > entry.ttl) {
            this.cache.delete(key);
            this.stats.totalMisses++;
            this.stats.totalEvictions++;
            return null;
        }

        // Verify content integrity if hashing is enabled
        if (this.options.enableContentHashing && entry.hash) {
            const currentHash = this.generateHash(entry.data);
            if (currentHash !== entry.hash) {
                console.warn(`Cache integrity check failed for key: ${key}`);
                this.cache.delete(key);
                this.stats.totalMisses++;
                return null;
            }
        }

        // Update access statistics
        entry.accessCount++;
        entry.lastAccessed = now;
        this.stats.totalHits++;

        return entry.data;
    }

    public set<T>(key: string, value: T, ttl?: number): void {
        const now = Date.now();
        const entryTtl = ttl ?? this.options.defaultTtl;
        const size = this.estimateSize(value);
        const hash = this.options.enableContentHashing ? this.generateHash(value) : '';

        const entry: CacheEntry<T> = {
            data: value,
            timestamp: now,
            ttl: entryTtl,
            accessCount: 0,
            lastAccessed: now,
            size,
            hash
        };

        // Check if we need to evict entries before adding new one
        this.ensureCapacity(size);

        this.cache.set(key, entry);
    }

    public has(key: string): boolean {
        const entry = this.cache.get(key);
        if (!entry) return false;

        // Check TTL expiration
        const now = Date.now();
        if (now - entry.timestamp > entry.ttl) {
            this.cache.delete(key);
            this.stats.totalEvictions++;
            return false;
        }

        return true;
    }

    public invalidate(key: string): boolean {
        return this.cache.delete(key);
    }

    public clear(): void {
        this.cache.clear();
        this.stats.totalEvictions += this.cache.size;
    }

    public getStats(): CacheStats {
        const entries = Array.from(this.cache.values());
        const totalMemory = entries.reduce((sum, entry) => sum + entry.size, 0);
        const totalRequests = this.stats.totalHits + this.stats.totalMisses;
        const hitRate = totalRequests > 0 ? (this.stats.totalHits / totalRequests) * 100 : 0;

        const timestamps = entries.map(e => e.timestamp);
        const accessCounts = entries.map(e => e.accessCount);

        return {
            totalEntries: this.cache.size,
            totalMemoryUsage: totalMemory,
            hitRate: Math.round(hitRate * 100) / 100,
            totalHits: this.stats.totalHits,
            totalMisses: this.stats.totalMisses,
            totalEvictions: this.stats.totalEvictions,
            oldestEntry: timestamps.length > 0 ? Math.min(...timestamps) : undefined,
            newestEntry: timestamps.length > 0 ? Math.max(...timestamps) : undefined,
            averageAccessCount: accessCounts.length > 0 ? 
                Math.round((accessCounts.reduce((sum, count) => sum + count, 0) / accessCounts.length) * 100) / 100 : 0
        };
    }

    public cleanup(): CacheEvictionResult {
        const now = Date.now();
        let evictedCount = 0;
        let freedMemory = 0;

        // Remove expired entries
        for (const [key, entry] of this.cache.entries()) {
            if (now - entry.timestamp > entry.ttl) {
                freedMemory += entry.size;
                evictedCount++;
                this.cache.delete(key);
            }
        }

        if (evictedCount > 0) {
            this.stats.totalEvictions += evictedCount;
        }

        return {
            evictedCount,
            freedMemory,
            reason: 'ttl'
        };
    }

    public updateOptions(options: Partial<CacheOptions>): void {
        this.options = { ...this.options, ...options };
        
        // Restart cleanup timer if interval changed
        if (options.cleanupInterval !== undefined) {
            this.stopCleanupTimer();
            this.startCleanupTimer();
        }
    }

    public destroy(): void {
        this.stopCleanupTimer();
        this.clear();
    }

    // Advanced cache management methods
    public getMemoryPressure(): number {
        const stats = this.getStats();
        return (stats.totalMemoryUsage / this.options.maxMemoryUsage) * 100;
    }

    public getCacheKeys(pattern?: RegExp): string[] {
        const keys = Array.from(this.cache.keys());
        if (pattern) {
            return keys.filter(key => pattern.test(key));
        }
        return keys;
    }

    public bulkInvalidate(pattern: RegExp): number {
        const keysToDelete = this.getCacheKeys(pattern);
        keysToDelete.forEach(key => this.cache.delete(key));
        return keysToDelete.length;
    }

    public exportCache(): Array<{ key: string; entry: CacheEntry<any> }> {
        return Array.from(this.cache.entries()).map(([key, entry]) => ({ key, entry }));
    }

    public importCache(data: Array<{ key: string; entry: CacheEntry<any> }>): void {
        data.forEach(({ key, entry }) => {
            // Validate entry before importing
            if (this.isValidCacheEntry(entry)) {
                this.cache.set(key, entry);
            }
        });
    }

    // Private helper methods
    private ensureCapacity(newEntrySize: number): void {
        const stats = this.getStats();
        const projectedMemory = stats.totalMemoryUsage + newEntrySize;
        const projectedEntries = stats.totalEntries + 1;

        // Check entry count limit first
        if (projectedEntries > this.options.maxEntries) {
            const entriesToEvict = projectedEntries - this.options.maxEntries;
            this.evictByPolicy(entriesToEvict);
        }

        // Check memory limit
        if (projectedMemory > this.options.maxMemoryUsage) {
            this.evictByMemoryPressure(newEntrySize);
        }
    }

    private evictByMemoryPressure(requiredSpace: number): CacheEvictionResult {
        let freedMemory = 0;
        let evictedCount = 0;
        const targetMemory = this.options.maxMemoryUsage * 0.8; // Target 80% usage

        const entries = Array.from(this.cache.entries());
        const sortedEntries = this.sortEntriesForEviction(entries);

        for (const [key, entry] of sortedEntries) {
            if (this.getStats().totalMemoryUsage - freedMemory <= targetMemory && 
                freedMemory >= requiredSpace) {
                break;
            }

            freedMemory += entry.size;
            evictedCount++;
            this.cache.delete(key);
        }

        this.stats.totalEvictions += evictedCount;
        return {
            evictedCount,
            freedMemory,
            reason: 'memory_pressure'
        };
    }

    private evictByPolicy(count: number): CacheEvictionResult {
        let evictedCount = 0;
        let freedMemory = 0;

        const entries = Array.from(this.cache.entries());
        const sortedEntries = this.sortEntriesForEviction(entries);

        for (let i = 0; i < Math.min(count, sortedEntries.length); i++) {
            const [key, entry] = sortedEntries[i];
            freedMemory += entry.size;
            evictedCount++;
            this.cache.delete(key);
        }

        this.stats.totalEvictions += evictedCount;
        return {
            evictedCount,
            freedMemory,
            reason: 'max_entries'
        };
    }

    private sortEntriesForEviction(entries: Array<[string, CacheEntry<any>]>): Array<[string, CacheEntry<any>]> {
        switch (this.options.evictionPolicy) {
            case 'lru':
                return entries.sort(([,a], [,b]) => a.lastAccessed - b.lastAccessed);
            case 'lfu':
                return entries.sort(([,a], [,b]) => a.accessCount - b.accessCount);
            case 'fifo':
                return entries.sort(([,a], [,b]) => a.timestamp - b.timestamp);
            case 'ttl':
                return entries.sort(([,a], [,b]) => {
                    const aExpiry = a.timestamp + a.ttl;
                    const bExpiry = b.timestamp + b.ttl;
                    return aExpiry - bExpiry;
                });
            default:
                return entries;
        }
    }

    private estimateSize(obj: any): number {
        if (obj === null || obj === undefined) return 0;
        
        const type = typeof obj;
        switch (type) {
            case 'boolean':
                return 4;
            case 'number':
                return 8;
            case 'string':
                return obj.length * 2; // Assuming UTF-16
            case 'object':
                // Prevent circular reference issues
                try {
                    if (Array.isArray(obj)) {
                        return obj.reduce((size, item) => size + this.estimateSize(item), 0);
                    }
                    
                    // Use a visited set to prevent infinite recursion on circular references
                    const visited = new WeakSet();
                    return this.estimateSizeRecursive(obj, visited);
                } catch (error) {
                    // Fallback for circular references or other issues
                    return JSON.stringify(obj).length * 2;
                }
            default:
                return JSON.stringify(obj).length * 2;
        }
    }

    private estimateSizeRecursive(obj: any, visited: WeakSet<object>): number {
        if (obj === null || obj === undefined) return 0;
        if (typeof obj !== 'object') return this.estimateSize(obj);
        if (visited.has(obj)) return 0; // Circular reference, skip
        
        visited.add(obj);
        
        try {
            if (Array.isArray(obj)) {
                return obj.reduce((size, item) => size + this.estimateSizeRecursive(item, visited), 0);
            }
            
            return Object.keys(obj).reduce((size, key) => {
                return size + this.estimateSize(key) + this.estimateSizeRecursive(obj[key], visited);
            }, 0);
        } catch (error) {
            return 100; // Fallback size for problematic objects
        }
    }

    private generateHash(data: any): string {
        if (!this.options.enableContentHashing) return '';
        
        try {
            const content = typeof data === 'string' ? data : JSON.stringify(data);
            return this.simpleHash(content);
        } catch (error) {
            console.warn('Failed to generate content hash:', error);
            return '';
        }
    }

    private simpleHash(str: string): string {
        let hash = 0;
        if (str.length === 0) return hash.toString();
        for (let i = 0; i < str.length; i++) {
            const char = str.charCodeAt(i);
            hash = ((hash << 5) - hash) + char;
            hash = hash & hash; // Convert to 32-bit integer
        }
        return Math.abs(hash).toString(36);
    }

    private isValidCacheEntry(entry: any): entry is CacheEntry<any> {
        return entry &&
            typeof entry.timestamp === 'number' &&
            typeof entry.ttl === 'number' &&
            typeof entry.accessCount === 'number' &&
            typeof entry.lastAccessed === 'number' &&
            typeof entry.size === 'number' &&
            entry.data !== undefined;
    }

    private startCleanupTimer(): void {
        if (this.cleanupTimer) return;
        
        this.cleanupTimer = setInterval(() => {
            this.cleanup();
        }, this.options.cleanupInterval);
    }

    private stopCleanupTimer(): void {
        if (this.cleanupTimer) {
            clearInterval(this.cleanupTimer);
            this.cleanupTimer = undefined;
        }
    }
}

// Singleton instance for global use
export const globalCacheService = new CacheService();