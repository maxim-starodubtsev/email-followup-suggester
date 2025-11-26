import { BatchProcessor } from "../../src/services/BatchProcessor";

describe("BatchProcessor", () => {
  let batchProcessor: BatchProcessor;

  beforeEach(() => {
    batchProcessor = new BatchProcessor();
  });

  describe("Basic Batch Processing", () => {
    it("should process items in batches", async () => {
      const items = Array.from({ length: 25 }, (_, i) => i);
      const processor = vi
        .fn()
        .mockImplementation(async (item: number) => item * 2);

      const result = await batchProcessor.processBatch(items, processor, {
        batchSize: 10,
      });

      expect(result.success).toBe(true);
      expect(result.results).toEqual(items.map((i) => i * 2));
      expect(result.totalProcessed).toBe(25);
      expect(result.totalBatches).toBe(3);
      expect(result.cancelled).toBe(false);
      expect(processor).toHaveBeenCalledTimes(25);
    });

    it("should use default batch size when not specified", async () => {
      const items = Array.from({ length: 15 }, (_, i) => i);
      const processor = vi
        .fn()
        .mockImplementation(async (item: number) => item);

      const result = await batchProcessor.processBatch(items, processor);

      expect(result.totalBatches).toBe(2); // 15 items / 10 default batch size = 2 batches
    });

    it("should handle empty array", async () => {
      const items: number[] = [];
      const processor = vi.fn();

      const result = await batchProcessor.processBatch(items, processor);

      expect(result.success).toBe(true);
      expect(result.results).toEqual([]);
      expect(result.totalProcessed).toBe(0);
      expect(result.totalBatches).toBe(0);
      expect(processor).not.toHaveBeenCalled();
    });
  });

  describe("Progress Tracking", () => {
    it("should report progress during processing", async () => {
      const items = Array.from({ length: 10 }, (_, i) => i);
      const processor = vi.fn().mockImplementation(async (item: number) => {
        await new Promise((resolve) => setTimeout(resolve, 10));
        return item;
      });
      const onProgress = vi.fn();

      await batchProcessor.processBatch(items, processor, {
        batchSize: 5,
        onProgress,
      });

      // Should report initial progress (0 items) plus progress after each item
      expect(onProgress).toHaveBeenCalledWith(0, 10, 0, 2);
      expect(onProgress).toHaveBeenCalledWith(10, 10, 2, 2);
      expect(onProgress.mock.calls.length).toBeGreaterThan(2);
    });

    it("should report batch completion", async () => {
      const items = Array.from({ length: 6 }, (_, i) => i);
      const processor = vi
        .fn()
        .mockImplementation(async (item: number) => item * 2);
      const onBatchComplete = vi.fn();

      await batchProcessor.processBatch(items, processor, {
        batchSize: 3,
        onBatchComplete,
      });

      expect(onBatchComplete).toHaveBeenCalledTimes(2);
      expect(onBatchComplete).toHaveBeenCalledWith(0, [0, 2, 4], []);
      expect(onBatchComplete).toHaveBeenCalledWith(1, [6, 8, 10], []);
    });
  });

  describe("Error Handling and Isolation", () => {
    it("should isolate errors to individual batches", async () => {
      const items = [1, 2, 3, 4, 5];
      const processor = vi.fn().mockImplementation(async (item: number) => {
        if (item === 3) {
          throw new Error(`Error processing item ${item}`);
        }
        return item * 2;
      });
      const onBatchError = vi.fn();

      const result = await batchProcessor.processBatch(items, processor, {
        batchSize: 2,
        onBatchError,
      });

      expect(result.success).toBe(false);
      expect(result.results).toEqual([2, 4, 8, 10]); // Items 1, 2, 4, 5 processed successfully
      expect(result.errors).toHaveLength(1);
      expect(result.errors[0].message).toContain("Error processing item 3");
      expect(onBatchError).toHaveBeenCalledTimes(1);
    }, 10000); // Increase timeout to 10 seconds

    it("should continue processing other batches when one batch has errors", async () => {
      const items = [1, 2, 3, 4, 5, 6];
      const processor = vi.fn().mockImplementation(async (item: number) => {
        if (item === 2) {
          throw new Error(`Error processing item ${item}`);
        }
        return item * 2;
      });

      const result = await batchProcessor.processBatch(items, processor, {
        batchSize: 2,
      });

      expect(result.success).toBe(false);
      expect(result.results).toEqual([2, 6, 8, 10, 12]); // All items except 2
      expect(result.errors).toHaveLength(1);
      expect(result.totalProcessed).toBe(5);
    }, 10000); // Increase timeout to 10 seconds

    it("should retry failed operations according to retry options", async () => {
      const items = [1];
      let attemptCount = 0;
      const processor = vi.fn().mockImplementation(async (item: number) => {
        attemptCount++;
        if (attemptCount < 3) {
          throw new Error("Temporary failure");
        }
        return item * 2;
      });

      const result = await batchProcessor.processBatch(items, processor, {
        retryOptions: {
          maxRetries: 3,
          baseDelay: 10,
          maxDelay: 100,
        },
      });

      expect(result.success).toBe(true);
      expect(result.results).toEqual([2]);
      expect(processor).toHaveBeenCalledTimes(3);
    });

    it("should fail after max retries exceeded", async () => {
      const items = [1];
      const processor = vi
        .fn()
        .mockRejectedValue(new Error("Persistent failure"));

      const result = await batchProcessor.processBatch(items, processor, {
        retryOptions: {
          maxRetries: 2,
          baseDelay: 10,
          maxDelay: 100,
        },
      });

      expect(result.success).toBe(false);
      expect(result.results).toEqual([]);
      expect(result.errors).toHaveLength(1);
      expect(processor).toHaveBeenCalledTimes(3); // Initial + 2 retries
    });
  });

  describe("Cancellation Support", () => {
    it("should support cancellation during processing", async () => {
      const items = Array.from({ length: 20 }, (_, i) => i);
      const processor = vi.fn().mockImplementation(async (item: number) => {
        await new Promise((resolve) => setTimeout(resolve, 50));
        return item;
      });

      const processingPromise = batchProcessor.processBatch(items, processor, {
        batchSize: 5,
      });

      // Cancel after a short delay
      setTimeout(() => {
        batchProcessor.cancelProcessing();
      }, 100);

      const result = await processingPromise;

      expect(result.cancelled).toBe(true);
      expect(result.totalProcessed).toBeLessThan(20);
    });

    it("should track active processing tasks", async () => {
      const items = [1, 2, 3];
      const processor = vi.fn().mockImplementation(async (item: number) => {
        await new Promise((resolve) => setTimeout(resolve, 100));
        return item;
      });

      const processingPromise = batchProcessor.processBatch(items, processor);

      // Check active tasks during processing
      setTimeout(() => {
        const activeTasks = batchProcessor.getActiveProcessingTasks();
        expect(activeTasks.length).toBeGreaterThan(0);
      }, 10);

      await processingPromise;

      // No active tasks after completion
      const activeTasks = batchProcessor.getActiveProcessingTasks();
      expect(activeTasks.length).toBe(0);
    });
  });

  describe("Concurrency Control", () => {
    it("should respect max concurrent batches limit", async () => {
      const items = Array.from({ length: 20 }, (_, i) => i);
      let concurrentCount = 0;
      let maxConcurrent = 0;

      const processor = vi.fn().mockImplementation(async (item: number) => {
        concurrentCount++;
        maxConcurrent = Math.max(maxConcurrent, concurrentCount);
        await new Promise((resolve) => setTimeout(resolve, 50));
        concurrentCount--;
        return item;
      });

      await batchProcessor.processBatch(items, processor, {
        batchSize: 5,
        maxConcurrentBatches: 2,
      });

      // Should not exceed the concurrent batch limit
      expect(maxConcurrent).toBeLessThanOrEqual(10); // 2 batches * 5 items per batch
    });
  });

  describe("Performance Metrics", () => {
    it("should track processing time", async () => {
      const items = [1, 2, 3];
      const processor = vi.fn().mockImplementation(async (item: number) => {
        await new Promise((resolve) => setTimeout(resolve, 10));
        return item;
      });

      const result = await batchProcessor.processBatch(items, processor);

      expect(result.processingTime).toBeGreaterThan(0);
      expect(result.processingTime).toBeGreaterThan(25); // At least 30ms for 3 items with 10ms delay each
    });

    it("should provide comprehensive batch result information", async () => {
      const items = [1, 2, 3, 4, 5];
      const processor = vi.fn().mockImplementation(async (item: number) => {
        if (item === 3) throw new Error("Test error");
        return item * 2;
      });

      const result = await batchProcessor.processBatch(items, processor, {
        batchSize: 2,
      });

      expect(result).toMatchObject({
        success: false,
        results: [2, 4, 8, 10],
        totalProcessed: 4,
        totalBatches: 3,
        cancelled: false,
      });
      expect(result.errors).toHaveLength(1);
      expect(result.processingTime).toBeGreaterThan(0);
    }, 10000); // Increase timeout to 10 seconds
  });

  describe("Integration Scenarios", () => {
    it("should handle complex processing scenarios", async () => {
      interface EmailData {
        id: string;
        subject: string;
        priority: "high" | "medium" | "low";
      }

      const emails: EmailData[] = [
        { id: "1", subject: "Urgent: Meeting", priority: "high" },
        { id: "2", subject: "Regular update", priority: "medium" },
        { id: "3", subject: "FYI", priority: "low" },
        { id: "4", subject: "CRITICAL: Server down", priority: "high" },
        { id: "5", subject: "Weekly report", priority: "medium" },
      ];

      const processor = vi.fn().mockImplementation(async (email: EmailData) => {
        // Simulate processing time based on priority
        const delay =
          email.priority === "high"
            ? 100
            : email.priority === "medium"
              ? 50
              : 25;
        await new Promise((resolve) => setTimeout(resolve, delay));

        return {
          ...email,
          processed: true,
          processedAt: new Date(),
        };
      });

      const progressUpdates: Array<{ current: number; total: number }> = [];
      const onProgress = (current: number, total: number) => {
        progressUpdates.push({ current, total });
      };

      const result = await batchProcessor.processBatch(emails, processor, {
        batchSize: 2,
        onProgress,
      });

      expect(result.success).toBe(true);
      expect(result.results).toHaveLength(5);
      expect(result.results.every((r) => (r as any).processed)).toBe(true);
      expect(progressUpdates.length).toBeGreaterThan(0);
      expect(progressUpdates[progressUpdates.length - 1]).toEqual({
        current: 5,
        total: 5,
      });
    });
  });
});
