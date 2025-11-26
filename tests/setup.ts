// Mock Office.js for testing using shared factory
import { vi } from "vitest";
import { OfficeMockFactory } from "./mocks/OfficeMockFactory";

(global as any).Office = OfficeMockFactory.createJestMock({
  userEmail: "test@example.com",
  userName: "Test User",
});

// Setup fetch mock for API calls
global.fetch = vi.fn() as any;

beforeEach(() => {
  vi.clearAllMocks();
});
