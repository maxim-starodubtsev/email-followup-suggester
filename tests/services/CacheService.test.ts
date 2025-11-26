import { CacheService } from "../../src/services/CacheService";

describe("CacheService", () => {
  let cacheService: CacheService;

  beforeEach(() => {
    cacheService = new CacheService({
      defaultTtl: 1000, // 1 second for faster tests
      maxMemoryUsage: 1024 * 1024, // 1MB
      maxEntries: 100,
      cleanupInterval: 100, // 100ms for faster tests
      enableContentHashing: true,
    });
  });

  afterEach(() => {
    cacheService.destroy();
  });

  describe("Basic Cache Operations", () => {
    it("should set and get cache entries", () => {
      const key = "test-key";
      const value = "test-value";

      cacheService.set(key, value);
      const retrieved = cacheService.get(key);

      expect(retrieved).toBe(value);
    });

    it("should return null for non-existent keys", () => {
      const result = cacheService.get("non-existent");
      expect(result).toBeNull();
    });

    it("should check if key exists", () => {
      const key = "test-key";
      expect(cacheService.has(key)).toBe(false);

      cacheService.set(key, "value");
      expect(cacheService.has(key)).toBe(true);
    });

    it("should invalidate cache entries", () => {
      const key = "test-key";
      cacheService.set(key, "value");

      const result = cacheService.invalidate(key);
      expect(result).toBe(true);
      expect(cacheService.has(key)).toBe(false);
    });

    it("should clear all cache entries", () => {
      cacheService.set("key1", "value1");
      cacheService.set("key2", "value2");

      cacheService.clear();

      expect(cacheService.has("key1")).toBe(false);
      expect(cacheService.has("key2")).toBe(false);
    });
  });

  describe("TTL and Expiration", () => {
    it("should expire entries after TTL", async () => {
      const key = "expire-test";
      const value = "expire-value";

      cacheService.set(key, value, 50); // 50ms TTL
      expect(cacheService.get(key)).toBe(value);

      await new Promise((resolve) => setTimeout(resolve, 60));
      expect(cacheService.get(key)).toBeNull();
    });

    it("should use default TTL when not specified", () => {
      const key = "default-ttl-test";
      cacheService.set(key, "value");

      expect(cacheService.has(key)).toBe(true);
    });

    it("should handle custom TTL per entry", async () => {
      cacheService.set("short", "value", 50);
      cacheService.set("long", "value", 500);

      await new Promise((resolve) => setTimeout(resolve, 60));

      expect(cacheService.has("short")).toBe(false);
      expect(cacheService.has("long")).toBe(true);
    });
  });

  describe("Statistics Tracking", () => {
    it("should track cache hits and misses", () => {
      cacheService.set("key1", "value1");

      // Generate hits
      cacheService.get("key1");
      cacheService.get("key1");

      // Generate misses
      cacheService.get("non-existent1");
      cacheService.get("non-existent2");

      const stats = cacheService.getStats();
      expect(stats.totalHits).toBe(2);
      expect(stats.totalMisses).toBe(2);
      expect(stats.hitRate).toBe(50);
    });

    it("should track memory usage", () => {
      const largeValue = "x".repeat(1000);
      cacheService.set("large-key", largeValue);

      const stats = cacheService.getStats();
      expect(stats.totalMemoryUsage).toBeGreaterThan(1000);
      expect(stats.totalEntries).toBe(1);
    });

    it("should track access counts", () => {
      cacheService.set("popular-key", "value");

      // Access multiple times
      for (let i = 0; i < 5; i++) {
        cacheService.get("popular-key");
      }

      const stats = cacheService.getStats();
      expect(stats.averageAccessCount).toBe(5);
    });
  });

  describe("Memory Management", () => {
    it("should evict entries when memory limit is reached", () => {
      const smallCache = new CacheService({
        maxMemoryUsage: 500, // Very small limit
        evictionPolicy: "lru",
      });

      // Fill cache beyond limit
      for (let i = 0; i < 10; i++) {
        smallCache.set(`key${i}`, "x".repeat(100));
      }

      const stats = smallCache.getStats();
      expect(stats.totalMemoryUsage).toBeLessThanOrEqual(500);

      smallCache.destroy();
    });

    it("should respect max entries limit", () => {
      const limitedCache = new CacheService({
        maxEntries: 3,
        evictionPolicy: "fifo",
      });

      // Add more entries than limit
      for (let i = 0; i < 5; i++) {
        limitedCache.set(`key${i}`, `value${i}`);
      }

      const stats = limitedCache.getStats();
      expect(stats.totalEntries).toBeLessThanOrEqual(3);

      limitedCache.destroy();
    });

    it("should report memory pressure correctly", () => {
      const smallCache = new CacheService({
        maxMemoryUsage: 1000,
      });

      smallCache.set("key1", "x".repeat(300));
      const pressure1 = smallCache.getMemoryPressure();

      smallCache.set("key2", "x".repeat(500));
      const pressure2 = smallCache.getMemoryPressure();

      expect(pressure2).toBeGreaterThan(pressure1);
      expect(pressure2).toBeGreaterThanOrEqual(80); // Should be at least 80%

      smallCache.destroy();
    });
  });

  describe("Eviction Policies", () => {
    it("should evict using LRU policy", async () => {
      const lruCache = new CacheService({
        maxEntries: 3,
        evictionPolicy: "lru",
      });

      // Add entries
      lruCache.set("key1", "value1");
      await new Promise((resolve) => setTimeout(resolve, 10)); // Ensure different timestamps
      lruCache.set("key2", "value2");
      await new Promise((resolve) => setTimeout(resolve, 10)); // Ensure different timestamps
      lruCache.set("key3", "value3");

      // Access key1 and key3 to make them recently used, leaving key2 as LRU
      await new Promise((resolve) => setTimeout(resolve, 10));
      lruCache.get("key1");
      await new Promise((resolve) => setTimeout(resolve, 10));
      lruCache.get("key3");

      // Add new entry to trigger eviction - key2 should be evicted (least recently used)
      await new Promise((resolve) => setTimeout(resolve, 10));
      lruCache.set("key4", "value4");

      // Verify that key2 was evicted and others remain
      expect(lruCache.has("key1")).toBe(true);
      expect(lruCache.has("key2")).toBe(false);
      expect(lruCache.has("key3")).toBe(true);
      expect(lruCache.has("key4")).toBe(true);

      lruCache.destroy();
    });

    it("should evict using LFU policy", () => {
      const lfuCache = new CacheService({
        maxEntries: 3,
        evictionPolicy: "lfu",
      });

      lfuCache.set("key1", "value1");
      lfuCache.set("key2", "value2");
      lfuCache.set("key3", "value3");

      // Access key1 multiple times
      for (let i = 0; i < 5; i++) {
        lfuCache.get("key1");
      }

      // Add new entry to trigger eviction
      lfuCache.set("key4", "value4");

      // key2 or key3 should be evicted (least frequently used)
      expect(lfuCache.has("key1")).toBe(true);
      const remainingKeys = ["key2", "key3", "key4"].filter((key) =>
        lfuCache.has(key),
      );
      expect(remainingKeys.length).toBe(2);

      lfuCache.destroy();
    });
  });

  describe("Content Hashing", () => {
    it("should detect content changes", () => {
      const key = "hash-test";
      const originalValue = { data: "original" };

      cacheService.set(key, originalValue);
      const retrieved = cacheService.get(key);

      expect(retrieved).toEqual(originalValue);
    });

    it("should handle hash generation errors gracefully", () => {
      const key = "circular-ref";
      const circularObj: any = { data: "test" };
      circularObj.self = circularObj; // Create circular reference

      // Should not throw error
      expect(() => {
        cacheService.set(key, circularObj);
      }).not.toThrow();
    });
  });

  describe("Bulk Operations", () => {
    it("should invalidate entries by pattern", () => {
      cacheService.set("user:1:profile", "profile1");
      cacheService.set("user:1:settings", "settings1");
      cacheService.set("user:2:profile", "profile2");
      cacheService.set("other:data", "other");

      const invalidatedCount = cacheService.bulkInvalidate(/^user:1:/);

      expect(invalidatedCount).toBe(2);
      expect(cacheService.has("user:1:profile")).toBe(false);
      expect(cacheService.has("user:1:settings")).toBe(false);
      expect(cacheService.has("user:2:profile")).toBe(true);
      expect(cacheService.has("other:data")).toBe(true);
    });

    it("should get cache keys by pattern", () => {
      cacheService.set("email:123", "email1");
      cacheService.set("email:456", "email2");
      cacheService.set("thread:789", "thread1");

      const emailKeys = cacheService.getCacheKeys(/^email:/);
      expect(emailKeys).toHaveLength(2);
      expect(emailKeys).toContain("email:123");
      expect(emailKeys).toContain("email:456");
    });
  });

  describe("Cache Export/Import", () => {
    it("should export and import cache data", () => {
      cacheService.set("key1", "value1");
      cacheService.set("key2", { data: "complex" });

      const exported = cacheService.exportCache();
      expect(exported).toHaveLength(2);

      const newCache = new CacheService();
      newCache.importCache(exported);

      expect(newCache.get("key1")).toBe("value1");
      expect(newCache.get("key2")).toEqual({ data: "complex" });

      newCache.destroy();
    });

    it("should validate entries during import", () => {
      const validEntry = {
        data: "test",
        timestamp: Date.now(),
        ttl: 1000,
        accessCount: 0,
        lastAccessed: Date.now(),
        size: 10,
        hash: "",
      };

      const invalidData: any[] = [
        {
          key: "invalid",
          entry: {
            data: "test",
            timestamp: 0,
            ttl: 0,
            accessCount: 0,
            lastAccessed: 0,
            size: 0,
            hash: "",
          },
        }, // Invalid structure
        { key: "valid", entry: validEntry },
      ];

      cacheService.importCache(invalidData);

      expect(cacheService.has("invalid")).toBe(false);
      expect(cacheService.has("valid")).toBe(true);
    });
  });

  describe("Cleanup and Maintenance", () => {
    it("should cleanup expired entries", async () => {
      cacheService.set("expire1", "value1", 50);
      cacheService.set("expire2", "value2", 50);
      cacheService.set("keep", "value3", 1000);

      await new Promise((resolve) => setTimeout(resolve, 60));

      const result = cacheService.cleanup();

      expect(result.evictedCount).toBe(2);
      expect(result.reason).toBe("ttl");
      expect(cacheService.has("keep")).toBe(true);
    });

    it("should handle automatic cleanup timer", async () => {
      const quickCleanupCache = new CacheService({
        cleanupInterval: 50,
        defaultTtl: 25,
      });

      quickCleanupCache.set("auto-expire", "value");

      await new Promise((resolve) => setTimeout(resolve, 75));

      expect(quickCleanupCache.has("auto-expire")).toBe(false);

      quickCleanupCache.destroy();
    });
  });

  describe("Configuration Updates", () => {
    it("should update cache options", () => {
      cacheService.updateOptions({
        defaultTtl: 5000,
        maxEntries: 50,
      });

      cacheService.set("new-key", "new-value");
      expect(cacheService.has("new-key")).toBe(true);
    });

    it("should restart cleanup timer when interval changes", () => {
      const newCleanupInterval = 200;

      cacheService.updateOptions({
        cleanupInterval: newCleanupInterval,
      });

      // Timer should be restarted with new interval
      expect(() => {
        cacheService.updateOptions({ cleanupInterval: 50 });
      }).not.toThrow();
    });
  });

  describe("Integration with Real-World Scenarios", () => {
    it("should handle email analysis caching scenario", () => {
      const emailData = {
        id: "email123",
        subject: "Test Email",
        body: "This is a test email content",
        from: "test@example.com",
      };

      const analysisResult = {
        sentiment: "neutral",
        priority: "medium",
        summary: "Test email summary",
      };

      // Cache email data
      cacheService.set(`email:${emailData.id}`, emailData);

      // Cache analysis result
      cacheService.set(`analysis:${emailData.id}`, analysisResult);

      // Retrieve from cache
      const cachedEmail = cacheService.get(`email:${emailData.id}`);
      const cachedAnalysis = cacheService.get(`analysis:${emailData.id}`);

      expect(cachedEmail).toEqual(emailData);
      expect(cachedAnalysis).toEqual(analysisResult);
    });

    it("should handle LLM response caching scenario", () => {
      const threadMessages = [
        { id: "1", content: "Hello" },
        { id: "2", content: "Hi there" },
      ];

      const llmResponse = {
        summary: "Greeting exchange",
        suggestion: "Follow up if needed",
        confidence: 0.8,
      };

      // Create content-based cache key
      const contentHash = JSON.stringify(threadMessages);
      const cacheKey = `llm:thread:${Buffer.from(contentHash).toString("base64").slice(0, 16)}`;

      cacheService.set(cacheKey, llmResponse, 24 * 60 * 60 * 1000); // 24 hour TTL

      const cachedResponse = cacheService.get(cacheKey);
      expect(cachedResponse).toEqual(llmResponse);
    });
  });
});
