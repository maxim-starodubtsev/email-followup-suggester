import { ConfigurationService } from '../../src/services/ConfigurationService';
import { Configuration } from '../../src/models/Configuration';

describe('ConfigurationService', () => {
  let service: ConfigurationService;
  let mockRoamingSettings: any;

  beforeEach(() => {
    service = new ConfigurationService();
    
    // Mock Office.context.roamingSettings
    mockRoamingSettings = {
      get: jest.fn(),
      set: jest.fn(),
      saveAsync: jest.fn()
    };
    
    (global as any).Office.context.roamingSettings = mockRoamingSettings;
    (global as any).Office.context.mailbox.getUserIdentityTokenAsync = jest.fn();
    
    // Clear localStorage
    localStorage.clear();
  });

  describe('Configuration Loading', () => {
    it('should return default configuration when no stored config exists', async () => {
      mockRoamingSettings.get.mockReturnValue(null);
      
      const config = await service.getConfiguration();
      
      expect(config.emailCount).toBe(25);
      expect(config.daysBack).toBe(7);
      expect(config.enableLlmSummary).toBe(false);
      expect(config.enableLlmSuggestions).toBe(false);
      expect(config.snoozeOptions).toHaveLength(7);
    });

    it('should merge stored configuration with defaults', async () => {
      const storedConfig = {
        emailCount: 50,
        enableLlmSummary: true,
        customProperty: 'test'
      };
      
      mockRoamingSettings.get.mockReturnValue(storedConfig);
      
      const config = await service.getConfiguration();
      
      expect(config.emailCount).toBe(50);
      expect(config.enableLlmSummary).toBe(true);
      expect(config.daysBack).toBe(7); // Should have default value
      expect(config.snoozeOptions).toHaveLength(7); // Should have default array
    });

    it('should properly convert date strings to Date objects', async () => {
      const dateString = '2025-01-01T00:00:00.000Z';
      const storedConfig = {
        lastAnalysisDate: dateString
      };
      
      mockRoamingSettings.get.mockReturnValue(storedConfig);
      
      const config = await service.getConfiguration();
      
      expect(config.lastAnalysisDate).toBeInstanceOf(Date);
      expect(config.lastAnalysisDate.toISOString()).toBe(dateString);
    });
  });

  describe('Configuration Saving', () => {
    it('should save configuration to Office roaming settings', async () => {
      const config: Configuration = {
        emailCount: 50,
        daysBack: 14,
        lastAnalysisDate: new Date('2025-01-01'),
        autoRefreshInterval: 60,
        priorityThresholds: { high: 8, medium: 4, low: 2 },
        snoozeOptions: [],
        enableLlmSummary: true,
        enableLlmSuggestions: true,
        selectedAccounts: ['test@example.com'],
        showSnoozedEmails: true,
        showDismissedEmails: false
      };

      mockRoamingSettings.saveAsync.mockImplementation((callback: any) => {
        callback({ status: 'succeeded' });
      });

      await service.saveConfiguration(config);

      expect(mockRoamingSettings.set).toHaveBeenCalledWith(
        'followup-suggester-config',
        expect.objectContaining({
          emailCount: 50,
          enableLlmSummary: true,
          lastAnalysisDate: '2025-01-01T00:00:00.000Z'
        })
      );
      expect(mockRoamingSettings.saveAsync).toHaveBeenCalled();
    });

    it('should handle save failures and fallback to localStorage', async () => {
      const config = service.getDefaultConfiguration();
      
      mockRoamingSettings.saveAsync.mockImplementation((callback: any) => {
        callback({ 
          status: 'failed', 
          error: { message: 'Network error' }
        });
      });

      // Should not throw error due to localStorage fallback
      await expect(service.saveConfiguration(config)).resolves.not.toThrow();
      
      // Verify localStorage was used
      const stored = localStorage.getItem('followup-suggester-config');
      expect(stored).toBeTruthy();
    });
  });

  describe('Configuration Management', () => {
    it('should reset configuration to defaults', async () => {
      mockRoamingSettings.saveAsync.mockImplementation((callback: any) => {
        callback({ status: 'succeeded' });
      });

      await service.resetConfiguration();

      expect(mockRoamingSettings.set).toHaveBeenCalledWith(
        'followup-suggester-config',
        expect.objectContaining({
          emailCount: 25,
          daysBack: 7,
          enableLlmSummary: false
        })
      );
    });

    it('should return default configuration', () => {
      const defaultConfig = service.getDefaultConfiguration();
      
      expect(defaultConfig.emailCount).toBe(25);
      expect(defaultConfig.daysBack).toBe(7);
      expect(defaultConfig.enableLlmSummary).toBe(false);
      expect(defaultConfig.snoozeOptions).toHaveLength(7);
    });

    it('should update last analysis date', async () => {
      const originalDate = new Date('2025-01-01');
      mockRoamingSettings.get.mockReturnValue({
        lastAnalysisDate: originalDate.toISOString()
      });
      mockRoamingSettings.saveAsync.mockImplementation((callback: any) => {
        callback({ status: 'succeeded' });
      });

      await service.updateLastAnalysisDate();

      expect(mockRoamingSettings.set).toHaveBeenCalledWith(
        'followup-suggester-config',
        expect.objectContaining({
          lastAnalysisDate: expect.any(String)
        })
      );
    });

    it('should check if analysis is stale', async () => {
      const oldDate = new Date(Date.now() - 45 * 60 * 1000); // 45 minutes ago
      mockRoamingSettings.get.mockReturnValue({
        lastAnalysisDate: oldDate.toISOString()
      });

      const isStale = await service.isAnalysisStale(30); // 30 minute threshold
      
      expect(isStale).toBe(true);
    });

    it('should check if analysis is fresh', async () => {
      const recentDate = new Date(Date.now() - 15 * 60 * 1000); // 15 minutes ago
      mockRoamingSettings.get.mockReturnValue({
        lastAnalysisDate: recentDate.toISOString()
      });

      const isStale = await service.isAnalysisStale(30); // 30 minute threshold
      
      expect(isStale).toBe(false);
    });
  });

  describe('Account Management', () => {
    it('should get available accounts', async () => {
      (global as any).Office.context.mailbox.getUserIdentityTokenAsync.mockImplementation((callback: any) => {
        callback({ status: 'succeeded' });
      });
      (global as any).Office.context.mailbox.userProfile.emailAddress = 'user@example.com';

      const accounts = await service.getAvailableAccounts();

      expect(accounts).toEqual(['user@example.com']);
    });

    it('should handle errors when getting accounts', async () => {
      (global as any).Office.context.mailbox.getUserIdentityTokenAsync.mockImplementation((callback: any) => {
        callback({ status: 'failed' });
      });

      const accounts = await service.getAvailableAccounts();

      expect(accounts).toEqual([]);
    });

    it('should update selected accounts', async () => {
      mockRoamingSettings.get.mockReturnValue({});
      mockRoamingSettings.saveAsync.mockImplementation((callback: any) => {
        callback({ status: 'succeeded' });
      });

      const accounts = ['user1@example.com', 'user2@example.com'];
      await service.updateSelectedAccounts(accounts);

      expect(mockRoamingSettings.set).toHaveBeenCalledWith(
        'followup-suggester-config',
        expect.objectContaining({
          selectedAccounts: accounts
        })
      );
    });
  });

  describe('LLM Settings', () => {
    it('should update LLM settings', async () => {
      mockRoamingSettings.get.mockReturnValue({});
      mockRoamingSettings.saveAsync.mockImplementation((callback: any) => {
        callback({ status: 'succeeded' });
      });

      await service.updateLlmSettings(
        'https://api.example.com',
        'api-key-123',
        true,
        false
      );

      expect(mockRoamingSettings.set).toHaveBeenCalledWith(
        'followup-suggester-config',
        expect.objectContaining({
          llmApiEndpoint: 'https://api.example.com',
          llmApiKey: 'api-key-123',
          enableLlmSummary: true,
          enableLlmSuggestions: false
        })
      );
    });
  });
});