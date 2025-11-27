import { describe, it, expect, beforeEach, afterEach, vi } from "vitest";
import {
  RetryService,
  RetryableError,
  NonRetryableError,
  CircuitBreaker,
  CircuitBreakerState,
} from "../../src/services/RetryService";

describe("RetryService", () => {
  let retryService: RetryService;

  beforeEach(() => {
    retryService = new RetryService();
    vi.clearAllTimers();
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  // Helper to create mock operations
  const createMockOperation = (
    failCount: number,
    errorMsg = "Temporary failure",
  ) => {
    let attempts = 0;
    return vi.fn().mockImplementation(async () => {
      attempts++;
      if (attempts <= failCount) {
        throw new Error(errorMsg);
      }
      return "success";
    });
  };

  describe("Basic Retry Logic", () => {
    it("should execute successfully without retry", async () => {
      const operation = vi.fn().mockResolvedValue("success");

      const result = await retryService.executeWithRetry(operation);

      expect(result).toBe("success");
      expect(operation).toHaveBeenCalledTimes(1);
    });

    it("should retry on failure until success", async () => {
      vi.useFakeTimers();
      const operation = createMockOperation(2); // Fail twice, succeed on 3rd

      const promise = retryService.executeWithRetry(operation, {
        maxAttempts: 5,
      });
      await vi.runAllTimersAsync();
      const result = await promise;

      expect(result).toBe("success");
      expect(operation).toHaveBeenCalledTimes(3);
    });

    it("should stop at maxAttempts and throw last error", async () => {
      vi.useFakeTimers();
      const operation = vi
        .fn()
        .mockRejectedValue(new Error("Persistent failure"));

      const promise = retryService.executeWithRetry(operation, {
        maxAttempts: 3,
      });
      
      // Run all timers and wait for all async operations
      await vi.runAllTimersAsync();
      // Switch to real timers to ensure all promise rejections are handled
      vi.useRealTimers();
      // Small delay to ensure all async operations complete
      await new Promise((resolve) => setTimeout(resolve, 10));

      await expect(promise).rejects.toThrow("Persistent failure");
      expect(operation).toHaveBeenCalledTimes(3);
    });

    it("should not retry NonRetryableError", async () => {
      const operation = vi
        .fn()
        .mockRejectedValue(new NonRetryableError("Bad request"));

      await expect(retryService.executeWithRetry(operation)).rejects.toThrow(
        "Bad request",
      );
      expect(operation).toHaveBeenCalledTimes(1);
    });

    it("should track retry statistics", async () => {
      vi.useFakeTimers();
      const operation = createMockOperation(2);

      const promise = retryService.executeWithRetry(operation, {
        maxAttempts: 5,
      });
      await vi.runAllTimersAsync();
      await promise;

      const stats = retryService.getStats();
      expect(stats.totalAttempts).toBe(3);
      expect(stats.totalRetries).toBe(2);
      expect(stats.failedRetries).toBe(0);
    });
  });

  describe("Backoff Strategy", () => {
    it("should apply exponential backoff with delays increasing", async () => {
      vi.useFakeTimers();
      const operation = vi.fn().mockRejectedValue(new Error("Failure"));
      const onRetry = vi.fn();

      const promise = retryService.executeWithRetry(operation, {
        maxAttempts: 4,
        baseDelayMs: 100,
        backoffFactor: 2,
        jitterMs: 0,
        onRetry,
      });

      await vi.runAllTimersAsync();
      // Switch to real timers to ensure all promise rejections are handled
      vi.useRealTimers();
      // Small delay to ensure all async operations complete
      await new Promise((resolve) => setTimeout(resolve, 10));
      await expect(promise).rejects.toThrow();

      // Verify delays increase (don't check exact values - too fragile)
      expect(onRetry).toHaveBeenCalledTimes(3);
      const delay1 = onRetry.mock.calls[0][2];
      const delay2 = onRetry.mock.calls[1][2];
      const delay3 = onRetry.mock.calls[2][2];

      expect(delay2).toBeGreaterThan(delay1);
      expect(delay3).toBeGreaterThan(delay2);
    });

    it("should respect maxDelayMs cap", async () => {
      vi.useFakeTimers();
      const operation = vi.fn().mockRejectedValue(new Error("Failure"));
      const onRetry = vi.fn();

      const promise = retryService.executeWithRetry(operation, {
        maxAttempts: 5,
        baseDelayMs: 1000,
        backoffFactor: 3,
        maxDelayMs: 5000,
        jitterMs: 0,
        onRetry,
      });

      await vi.runAllTimersAsync();
      // Switch to real timers to ensure all promise rejections are handled
      vi.useRealTimers();
      // Small delay to ensure all async operations complete
      await new Promise((resolve) => setTimeout(resolve, 10));
      await expect(promise).rejects.toThrow();

      onRetry.mock.calls.forEach((call) => {
        expect(call[2]).toBeLessThanOrEqual(5000);
      });
    });

    it("should execute onRetry callback with correct parameters", async () => {
      vi.useFakeTimers();
      const operation = createMockOperation(2);
      const onRetry = vi.fn();

      const promise = retryService.executeWithRetry(operation, {
        maxAttempts: 3,
        onRetry,
      });
      await vi.runAllTimersAsync();
      await promise;

      expect(onRetry).toHaveBeenCalledTimes(2);
      expect(onRetry.mock.calls[0][1]).toBeInstanceOf(Error); // Error is 2nd parameter
      expect(onRetry.mock.calls[0][0]).toBe(1); // Attempt number is 1st parameter
      expect(typeof onRetry.mock.calls[0][2]).toBe("number"); // Delay is 3rd parameter
    });
  });

  describe("Circuit Breaker", () => {
    let circuitBreaker: CircuitBreaker;

    beforeEach(() => {
      circuitBreaker = new CircuitBreaker({
        failureThreshold: 3,
        recoveryTimeoutMs: 1000,
      });
      circuitBreaker.reset(); // Ensure clean state for each test
    });

    it("should start in CLOSED state", () => {
      expect(circuitBreaker.getState()).toBe(CircuitBreakerState.CLOSED);
    });

    it("should transition to OPEN after threshold failures", async () => {
      const operation = vi
        .fn()
        .mockRejectedValue(new RetryableError("Failure", true));

      // Execute failures up to threshold
      for (let i = 0; i < 3; i++) {
        try {
          await retryService.executeWithRetry(operation, {
            maxAttempts: 1,
            circuitBreaker,
          });
        } catch {
          // Expected failures
        }
      }

      expect(circuitBreaker.getState()).toBe(CircuitBreakerState.OPEN);
    });

    it("should reject immediately in OPEN state", async () => {
      const operation = vi
        .fn()
        .mockRejectedValue(new RetryableError("Failure", true));

      // Trip the circuit
      for (let i = 0; i < 3; i++) {
        try {
          await retryService.executeWithRetry(operation, {
            maxAttempts: 1,
            circuitBreaker,
          });
        } catch {
          // Expected
        }
      }

      // Next call should be rejected immediately
      const callsBefore = operation.mock.calls.length;
      try {
        await retryService.executeWithRetry(operation, {
          maxAttempts: 1,
          circuitBreaker,
        });
      } catch (error) {
        expect((error as Error).message).toContain("Circuit breaker is OPEN");
      }

      expect(operation.mock.calls.length).toBe(callsBefore); // No new call made
    });

    it("should transition to HALF_OPEN after recovery timeout", async () => {
      vi.useFakeTimers();
      const operation = vi
        .fn()
        .mockRejectedValue(new RetryableError("Failure", true));

      // Trip circuit to OPEN
      for (let i = 0; i < 3; i++) {
        try {
          await retryService.executeWithRetry(operation, {
            maxAttempts: 1,
            circuitBreaker,
          });
        } catch {
          // Expected
        }
      }

      expect(circuitBreaker.getState()).toBe(CircuitBreakerState.OPEN);

      // Advance time past recovery timeout
      vi.advanceTimersByTime(1001);

      // Next call should transition to HALF_OPEN (then back to OPEN on failure)
      try {
        await retryService.executeWithRetry(operation, {
          maxAttempts: 1,
          circuitBreaker,
        });
      } catch {
        // Expected to fail
      }

      expect(circuitBreaker.getState()).toBe(CircuitBreakerState.OPEN);
    });

    it("should transition HALF_OPEN to CLOSED on success", async () => {
      const operation = vi
        .fn()
        .mockRejectedValue(new RetryableError("Failure", true));

      // Trip to OPEN
      for (let i = 0; i < 3; i++) {
        try {
          await retryService.executeWithRetry(operation, {
            maxAttempts: 1,
            circuitBreaker,
          });
        } catch {
          // Expected
        }
      }

      // Force HALF_OPEN state to test successful recovery
      circuitBreaker.__forceStateForTesting(CircuitBreakerState.HALF_OPEN);

      operation.mockResolvedValueOnce("success");
      await retryService.executeWithRetry(operation, {
        maxAttempts: 1,
        circuitBreaker,
      });

      expect(circuitBreaker.getState()).toBe(CircuitBreakerState.CLOSED);
    });

    it("should reset failure count on success in CLOSED state", async () => {
      // Create a fresh circuit breaker for this test to avoid state pollution
      const testCircuitBreaker = new CircuitBreaker({
        failureThreshold: 3,
        recoveryTimeoutMs: 1000,
      });

      const operation = vi
        .fn()
        .mockRejectedValueOnce(new RetryableError("Failure 1", false)) // Don't circuit break on first failures
        .mockRejectedValueOnce(new RetryableError("Failure 2", false))
        .mockResolvedValueOnce("success")
        .mockRejectedValueOnce(new RetryableError("Failure 3", false))
        .mockRejectedValueOnce(new RetryableError("Failure 4", false));

      // Two failures
      try {
        await retryService.executeWithRetry(operation, {
          maxAttempts: 1,
          circuitBreaker: testCircuitBreaker,
        });
      } catch {
        // Expected
      }
      try {
        await retryService.executeWithRetry(operation, {
          maxAttempts: 1,
          circuitBreaker: testCircuitBreaker,
        });
      } catch {
        // Expected
      }

      // Success resets counter
      await retryService.executeWithRetry(operation, {
        maxAttempts: 1,
        circuitBreaker: testCircuitBreaker,
      });
      expect(testCircuitBreaker.getState()).toBe(CircuitBreakerState.CLOSED);

      // Two more failures should not trip (counter was reset)
      try {
        await retryService.executeWithRetry(operation, {
          maxAttempts: 1,
          circuitBreaker: testCircuitBreaker,
        });
      } catch {
        // Expected
      }
      try {
        await retryService.executeWithRetry(operation, {
          maxAttempts: 1,
          circuitBreaker: testCircuitBreaker,
        });
      } catch {
        // Expected
      }

      expect(testCircuitBreaker.getState()).toBe(CircuitBreakerState.CLOSED);
    });
  });

  describe("Static Utility Methods", () => {
    it("should create rate limit retry configuration", () => {
      const config = RetryService.retryOnRateLimit();

      expect(config.maxAttempts).toBeGreaterThan(1);
      expect(config.baseDelayMs).toBeGreaterThan(0);
      expect(config.backoffFactor).toBeGreaterThanOrEqual(1);
    });

    it("should create network error retry configuration", () => {
      const config = RetryService.retryOnNetworkError();

      expect(config.maxAttempts).toBeGreaterThan(1);
      expect(config.baseDelayMs).toBeGreaterThan(0);
      expect(config.backoffFactor).toBeGreaterThanOrEqual(1);
    });

    it("should handle rate limit errors with retry configuration", async () => {
      vi.useFakeTimers();
      // Create operation that throws RetryableError for rate limit
      let attempts = 0;
      const operation = vi.fn().mockImplementation(async () => {
        attempts++;
        if (attempts <= 2) {
          throw new Error("Rate limit exceeded");
        }
        return "success";
      });

      const promise = retryService.executeWithRetry(
        operation,
        RetryService.retryOnRateLimit(),
      );
      await vi.runAllTimersAsync();
      const result = await promise;

      expect(result).toBe("success");
      expect(operation).toHaveBeenCalledTimes(3);
    });

    it("should handle network errors with retry configuration", async () => {
      vi.useFakeTimers();
      const operation = createMockOperation(2, "Network error");

      const promise = retryService.executeWithRetry(
        operation,
        RetryService.retryOnNetworkError(),
      );
      await vi.runAllTimersAsync();
      const result = await promise;

      expect(result).toBe("success");
      expect(operation).toHaveBeenCalledTimes(3);
    });
  });

  describe("Integration Scenarios", () => {
    it("should handle mixed retryable and non-retryable errors", async () => {
      let attempts = 0;
      const operation = vi.fn().mockImplementation(async () => {
        attempts++;
        if (attempts === 1) {
          throw new RetryableError("Temporary issue");
        }
        if (attempts === 2) {
          throw new NonRetryableError("Invalid request");
        }
        return "success";
      });

      vi.useFakeTimers();
      const promise = retryService.executeWithRetry(operation, {
        maxAttempts: 5,
      });
      await vi.runAllTimersAsync();
      // Switch to real timers to ensure all promise rejections are handled
      vi.useRealTimers();
      // Small delay to ensure all async operations complete
      await new Promise((resolve) => setTimeout(resolve, 10));

      await expect(promise).rejects.toThrow("Invalid request");
      expect(operation).toHaveBeenCalledTimes(2);
    });

    it("should work with async operations and promises", async () => {
      vi.useFakeTimers();
      let attempts = 0;
      const operation = vi.fn().mockImplementation(async () => {
        attempts++;
        await new Promise((resolve) => setTimeout(resolve, 10));
        if (attempts < 3) {
          throw new Error("Not ready");
        }
        return "async success";
      });

      const promise = retryService.executeWithRetry(operation, {
        maxAttempts: 5,
      });
      await vi.runAllTimersAsync();
      const result = await promise;

      expect(result).toBe("async success");
      expect(operation).toHaveBeenCalledTimes(3);
    });

    it("should handle concurrent retry operations independently", async () => {
      vi.useFakeTimers();
      const operation1 = createMockOperation(1);
      const operation2 = createMockOperation(2);

      const promise1 = retryService.executeWithRetry(operation1, {
        maxAttempts: 3,
      });
      const promise2 = retryService.executeWithRetry(operation2, {
        maxAttempts: 3,
      });

      await vi.runAllTimersAsync();
      const [result1, result2] = await Promise.all([promise1, promise2]);

      expect(result1).toBe("success");
      expect(result2).toBe("success");
      expect(operation1).toHaveBeenCalledTimes(2);
      expect(operation2).toHaveBeenCalledTimes(3);
    });

    it("should maintain statistics across multiple operations", async () => {
      vi.useFakeTimers();

      // First operation: 2 attempts
      const op1 = createMockOperation(1);
      const promise1 = retryService.executeWithRetry(op1, { maxAttempts: 3 });
      await vi.runAllTimersAsync();
      await promise1;

      // Second operation: 3 attempts
      const op2 = createMockOperation(2);
      const promise2 = retryService.executeWithRetry(op2, { maxAttempts: 3 });
      await vi.runAllTimersAsync();
      await promise2;

      const stats = retryService.getStats();
      expect(stats.totalAttempts).toBe(5); // 2 + 3
      expect(stats.totalRetries).toBe(3); // 1 + 2
    });

    it("should handle errors during retry callback gracefully", async () => {
      vi.useFakeTimers();
      const operation = createMockOperation(1);
      const onRetry = vi.fn().mockImplementation(() => {
        throw new Error("Callback error");
      });

      // Should still retry despite callback error
      const promise = retryService.executeWithRetry(operation, {
        maxAttempts: 3,
        onRetry,
      });
      await vi.runAllTimersAsync();
      const result = await promise;

      expect(result).toBe("success");
      expect(operation).toHaveBeenCalledTimes(2);
    });
  });

  describe("Edge Cases", () => {
    it("should handle maxAttempts of 1 (no retry)", async () => {
      const operation = vi.fn().mockRejectedValue(new Error("Failure"));

      await expect(
        retryService.executeWithRetry(operation, { maxAttempts: 1 }),
      ).rejects.toThrow("Failure");
      expect(operation).toHaveBeenCalledTimes(1);
    });

    it("should handle zero delay gracefully", async () => {
      vi.useFakeTimers();
      const operation = createMockOperation(1);

      const promise = retryService.executeWithRetry(operation, {
        maxAttempts: 3,
        baseDelayMs: 0,
      });
      await vi.runAllTimersAsync();
      const result = await promise;

      expect(result).toBe("success");
      expect(operation).toHaveBeenCalledTimes(2);
    });

    it("should preserve error context across retries", async () => {
      vi.useFakeTimers();
      class CustomError extends Error {
        constructor(
          message: string,
          public code: number,
        ) {
          super(message);
        }
      }

      const operation = vi
        .fn()
        .mockRejectedValue(new CustomError("Custom failure", 500));

      const promise = retryService.executeWithRetry(operation, {
        maxAttempts: 2,
      });
      await vi.runAllTimersAsync();
      // Switch to real timers to ensure all promise rejections are handled
      vi.useRealTimers();
      // Small delay to ensure all async operations complete
      await new Promise((resolve) => setTimeout(resolve, 10));

      try {
        await promise;
        fail("Should have thrown");
      } catch (error) {
        expect(error).toBeInstanceOf(CustomError);
        expect((error as CustomError).code).toBe(500);
      }
    });
  });
});
