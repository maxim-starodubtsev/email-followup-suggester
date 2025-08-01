// Mock Office.js for testing
(global as any).Office = {
  context: {
    mailbox: {
      userProfile: {
        emailAddress: 'test@example.com',
        displayName: 'Test User',
        timeZone: 'UTC',
        accountType: 'enterprise'
      },
      makeEwsRequestAsync: jest.fn()
    },
    roamingSettings: {
      get: jest.fn(),
      set: jest.fn(),
      saveAsync: jest.fn()
    }
  },
  AsyncResultStatus: {
    Succeeded: 'succeeded',
    Failed: 'failed'
  }
};

// Setup fetch mock for API calls
global.fetch = jest.fn();

beforeEach(() => {
  jest.clearAllMocks();
});