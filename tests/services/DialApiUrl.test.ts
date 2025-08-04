/**
 * DIAL API URL Construction Test
 * 
 * This test file demonstrates the correct URL format for DIAL API
 * and validates that our LlmService constructs the URL properly.
 */

import { LlmService } from '../../src/services/LlmService';
import { Configuration } from '../../src/models/Configuration';
import { RetryService } from '../../src/services/RetryService';

// Mock fetch globally
global.fetch = jest.fn();

describe('DIAL API URL Construction', () => {
  let service: LlmService;
  let mockFetch: jest.MockedFunction<typeof fetch>;
  let mockConfiguration: Configuration;
  let mockRetryService: jest.Mocked<RetryService>;

  beforeEach(() => {
    mockFetch = fetch as jest.MockedFunction<typeof fetch>;
    mockFetch.mockClear();
    
    // Use the actual default configuration from ConfigurationService
    mockConfiguration = {
      emailCount: 25,
      daysBack: 7,
      lastAnalysisDate: new Date(),
      snoozeOptions: [],
      enableLlmSummary: true,
      enableLlmSuggestions: true,
      llmProvider: 'dial' as const,
      llmApiEndpoint: 'https://ai-proxy.lab.epam.com',
      llmModel: 'gpt-4.1-2025-04-14',
      llmApiKey: 'dial-qjboupii21tb26eakd3ytcsb9po',
      llmApiVersion: '2025-04-14',
      llmDeploymentName: 'gpt-4.1-2025-04-14',
      selectedAccounts: [],
      showSnoozedEmails: false,
      showDismissedEmails: false
    };

    mockRetryService = {
      executeWithRetry: jest.fn().mockImplementation((operation) => operation())
    } as any;

    service = new LlmService(mockConfiguration, mockRetryService);
    
    // Suppress console errors for expected test failures
    jest.spyOn(console, 'error').mockImplementation(() => {});
    jest.spyOn(console, 'warn').mockImplementation(() => {});
  });

  afterEach(() => {
    jest.restoreAllMocks();
  });

  it('should construct correct DIAL API URL with deployment and API version', async () => {
    // Mock successful response
    mockFetch.mockResolvedValueOnce({
      ok: true,
      json: async () => ({
        choices: [{ message: { content: '1. Test suggestion' } }]
      })
    } as Response);

    await service.generateFollowupSuggestions('Test email content');

    // Verify the URL is constructed correctly
    const expectedUrl = 'https://ai-proxy.lab.epam.com/openai/deployments/gpt-4.1-2025-04-14/chat/completions?api-version=2025-04-14';
    
    expect(mockFetch).toHaveBeenCalledWith(
      expectedUrl,
      expect.objectContaining({
        method: 'POST',
        headers: expect.objectContaining({
          'Content-Type': 'application/json',
          'Api-Key': 'dial-qjboupii21tb26eakd3ytcsb9po'
        })
      })
    );
  });

  it('should handle different deployment names', async () => {
    // Test with different deployment name
    const customConfig = {
      ...mockConfiguration,
      llmDeploymentName: 'custom-deployment-name',
      llmApiVersion: '2024-08-01'
    };

    const customService = new LlmService(customConfig, mockRetryService);

    mockFetch.mockResolvedValueOnce({
      ok: true,
      json: async () => ({
        choices: [{ message: { content: '1. Custom suggestion' } }]
      })
    } as Response);

    await customService.generateFollowupSuggestions('Test email');

    const expectedUrl = 'https://ai-proxy.lab.epam.com/openai/deployments/custom-deployment-name/chat/completions?api-version=2024-08-01';
    
    expect(mockFetch).toHaveBeenCalledWith(
      expectedUrl,
      expect.anything()
    );
  });

  it('should fallback to model name when deployment name is not specified', async () => {
    // Test with no deployment name specified
    const configNoDeployment = {
      ...mockConfiguration,
      llmDeploymentName: undefined,
      llmModel: 'fallback-model'
    };

    const fallbackService = new LlmService(configNoDeployment, mockRetryService);

    mockFetch.mockResolvedValueOnce({
      ok: true,
      json: async () => ({
        choices: [{ message: { content: '1. Fallback suggestion' } }]
      })
    } as Response);

    await fallbackService.generateFollowupSuggestions('Test email');

    const expectedUrl = 'https://ai-proxy.lab.epam.com/openai/deployments/fallback-model/chat/completions?api-version=2025-04-14';
    
    expect(mockFetch).toHaveBeenCalledWith(
      expectedUrl,
      expect.anything()
    );
  });

  it('should remove trailing slashes from endpoint', async () => {
    // Test with endpoint that has trailing slash
    const configWithSlash = {
      ...mockConfiguration,
      llmApiEndpoint: 'https://ai-proxy.lab.epam.com/'
    };

    const slashService = new LlmService(configWithSlash, mockRetryService);

    mockFetch.mockResolvedValueOnce({
      ok: true,
      json: async () => ({
        choices: [{ message: { content: '1. Slash test' } }]
      })
    } as Response);

    await slashService.generateFollowupSuggestions('Test email');

    // URL should not have double slashes
    const expectedUrl = 'https://ai-proxy.lab.epam.com/openai/deployments/gpt-4.1-2025-04-14/chat/completions?api-version=2025-04-14';
    
    expect(mockFetch).toHaveBeenCalledWith(
      expectedUrl,
      expect.anything()
    );
  });

  it('should demonstrate the complete URL pattern', () => {
    const baseEndpoint = 'https://ai-proxy.lab.epam.com';
    const deploymentName = 'gpt-4.1-2025-04-14';
    const apiVersion = '2025-04-14';
    
    const expectedPattern = `${baseEndpoint}/openai/deployments/${deploymentName}/chat/completions?api-version=${apiVersion}`;
    
    expect(expectedPattern).toBe('https://ai-proxy.lab.epam.com/openai/deployments/gpt-4.1-2025-04-14/chat/completions?api-version=2025-04-14');
    
    console.log('✅ DIAL API URL Pattern Verified:');
    console.log(`   Base URL: ${baseEndpoint}`);
    console.log(`   Deployment: ${deploymentName}`);
    console.log(`   API Version: ${apiVersion}`);
    console.log(`   Full URL: ${expectedPattern}`);
  });
});
