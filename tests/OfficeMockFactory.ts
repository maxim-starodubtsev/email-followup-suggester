/**
 * Shared Office.js Mock Factory
 *
 * Provides reusable mock implementations for both:
 * - Jest unit tests (tests/setup.ts)
 * - Standalone debug harness (debug/mock-office.js)
 */

import { vi } from "vitest";

export interface OfficeMockConfig {
  userEmail?: string;
  userName?: string;
  timeZone?: string;
  mode?: "jest" | "browser";
}

/**
 * Type definition for Office.js mock object
 */
export interface OfficeMock {
  context: {
    mailbox: {
      userProfile: {
        emailAddress: string;
        displayName: string;
        timeZone: string;
        accountType: string;
      };
      makeEwsRequestAsync: Function;
      displayNewMessageForm: Function;
      getUserIdentityTokenAsync?: Function;
    };
    roamingSettings: {
      get: Function;
      set: Function;
      saveAsync: Function;
      remove?: Function;
    };
  };
  AsyncResultStatus: {
    Succeeded: string;
    Failed: string;
  };
  HostType: {
    Outlook: string;
  };
  onReady: Function;
}

/**
 * Timing constants for mock operations (in milliseconds)
 */
const MOCK_TIMING = {
  EWS_REQUEST_DELAY: 100,
  USER_TOKEN_DELAY: 10,
  SETTINGS_SAVE_DELAY: 50,
  OFFICE_READY_DELAY: 10,
} as const;

export class OfficeMockFactory {
  private static readonly DEFAULT_CONFIG: Required<OfficeMockConfig> = {
    userEmail: "test@example.com",
    userName: "Test User",
    timeZone: "UTC",
    mode: "jest",
  };

  /**
   * Create Office.js mock for Jest tests
   */
  static createJestMock(config: OfficeMockConfig = {}): OfficeMock {
    const cfg = { ...this.DEFAULT_CONFIG, ...config, mode: "jest" as const };

    return {
      context: {
        mailbox: {
          userProfile: {
            emailAddress: cfg.userEmail,
            displayName: cfg.userName,
            timeZone: cfg.timeZone,
            accountType: "enterprise",
          },
          makeEwsRequestAsync: vi.fn(),
          displayNewMessageForm: vi.fn(),
        },
        roamingSettings: {
          get: vi.fn(),
          set: vi.fn(),
          saveAsync: vi.fn(),
        },
      },
      AsyncResultStatus: {
        Succeeded: "succeeded",
        Failed: "failed",
      },
      HostType: {
        Outlook: "Outlook",
      },
      onReady: (callback: (info: any) => void) => {
        callback({ host: "Outlook" });
      },
    };
  }

  /**
   * Create Office.js mock for browser-based debug harness
   */
  static createBrowserMock(config: OfficeMockConfig = {}): OfficeMock {
    const cfg = { ...this.DEFAULT_CONFIG, ...config, mode: "browser" as const };
    const mockData = {
      userProfile: {
        emailAddress: cfg.userEmail,
        displayName: cfg.userName,
        timeZone: cfg.timeZone,
        accountType: "enterprise",
      },
      roamingSettings: {} as Record<string, any>,
    };

    return {
      context: {
        mailbox: {
          userProfile: mockData.userProfile,

          makeEwsRequestAsync: function (
            ewsRequest: string,
            callback: Function,
          ) {
            console.log(
              "[Mock Office] EWS Request:",
              ewsRequest.substring(0, 100) + "...",
            );

            // Simple mock response - let the debug harness provide test data
            setTimeout(() => {
              callback({
                status: "succeeded",
                value:
                  '<?xml version="1.0"?><soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"><soap:Body><m:FindItemResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"><m:ResponseMessages><m:FindItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:RootFolder><t:Items xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"></t:Items></m:RootFolder></m:FindItemResponseMessage></m:ResponseMessages></m:FindItemResponse></soap:Body></soap:Envelope>',
              });
            }, MOCK_TIMING.EWS_REQUEST_DELAY);
          },

          displayNewMessageForm: function (messageDetails: any) {
            console.log("[Mock Office] Display New Message:", messageDetails);
          },

          getUserIdentityTokenAsync: function (callback: Function) {
            setTimeout(() => {
              callback({
                status: "succeeded",
                value: "mock-user-identity-token",
              });
            }, MOCK_TIMING.USER_TOKEN_DELAY);
          },
        },

        roamingSettings: {
          get: function (name: string) {
            return mockData.roamingSettings[name] || null;
          },

          set: function (name: string, value: any) {
            mockData.roamingSettings[name] = value;
            console.log("[Mock Office] Setting saved:", name);
          },

          saveAsync: function (callback?: Function) {
            setTimeout(() => {
              console.log("[Mock Office] Settings persisted");
              if (callback) {
                callback({ status: "succeeded" });
              }
            }, MOCK_TIMING.SETTINGS_SAVE_DELAY);
          },

          remove: function (name: string) {
            delete mockData.roamingSettings[name];
          },
        },
      },
      AsyncResultStatus: {
        Succeeded: "succeeded",
        Failed: "failed",
      },
      HostType: {
        Outlook: "Outlook",
      },
      onReady: function (callback: Function) {
        setTimeout(() => {
          console.log("[Mock Office] Office.onReady called");
          callback({ host: "Outlook" });
        }, MOCK_TIMING.OFFICE_READY_DELAY);
      },
    };
  }

  /**
   * Get mock data storage (for browser mode)
   */
  static getMockDataStorage() {
    return {
      roamingSettings: {} as Record<string, any>,
      userProfile: {
        emailAddress: "test@example.com",
        displayName: "Test User",
        timeZone: "UTC",
        accountType: "enterprise",
      },
    };
  }
}
