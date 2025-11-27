// Mock Office.js for testing using shared factory
import { vi } from "vitest";
import { OfficeMockFactory } from "./mocks/OfficeMockFactory";

(global as any).Office = OfficeMockFactory.createJestMock({
  userEmail: "test@example.com",
  userName: "Test User",
});

// Setup fetch mock for API calls
global.fetch = vi.fn() as any;

// Handle unhandled promise rejections during tests
// This prevents false positives when testing error scenarios
const unhandledRejections = new Set<Promise<any>>();
process.on("unhandledRejection", (reason, promise) => {
  // Only track rejections that are part of our test scenarios
  // These are expected when testing retry logic with failures
  unhandledRejections.add(promise);
  // Silently handle expected test rejections
  promise.catch(() => {
    unhandledRejections.delete(promise);
  });
});

beforeEach(() => {
  vi.clearAllMocks();
  // Clear tracked rejections before each test
  unhandledRejections.clear();
});
