import { LlmService } from '../../src/services/LlmService';
import { Configuration } from '../../src/models/Configuration';
import { RetryService } from '../../src/services/RetryService';

// Mock fetch globally
global.fetch = jest.fn();

describe('LlmService', () => {
  let service: LlmService;
  let mockFetch: jest.MockedFunction<typeof fetch>;
  let mockConfiguration: Configuration;
  let mockRetryService: jest.Mocked<RetryService>;

  beforeEach(() => {
    mockFetch = fetch as jest.MockedFunction<typeof fetch>;
    mockFetch.mockClear();
    
    // Mock configuration
    mockConfiguration = {
      emailCount: 10,
      daysBack: 7,
      lastAnalysisDate: new Date(),
      snoozeOptions: [],
      enableLlmSummary: true,
      enableLlmSuggestions: true,
      llmApiEndpoint: 'https://ai-proxy.lab.epam.com/openai/chat/completions',
      llmApiKey: 'test-api-key',
      llmModel: 'gpt-35-turbo',
      selectedAccounts: [],
      showSnoozedEmails: false,
      showDismissedEmails: false
    };

    // Mock retry service
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

  describe('generateFollowupSuggestions', () => {
    it('should generate followup suggestions successfully', async () => {
      const mockResponse = `1. Thank you for your email. I'll review this and get back to you.
2. I appreciate your patience while I look into this matter.
3. Let me follow up with the team and provide an update soon.`;

      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({
          choices: [{
            message: { content: mockResponse }
          }]
        })
      } as Response);

      const suggestions = await service.generateFollowupSuggestions('Test email content');

      expect(suggestions).toHaveLength(3);
      expect(suggestions[0]).toBe("Thank you for your email. I'll review this and get back to you.");
      expect(suggestions[1]).toBe("I appreciate your patience while I look into this matter.");
      expect(suggestions[2]).toBe("Let me follow up with the team and provide an update soon.");
    });

    it('should handle API errors', async () => {
      mockFetch.mockResolvedValueOnce({
        ok: false,
        status: 500,
        statusText: 'Internal Server Error',
        text: async () => 'Server Error'
      } as Response);

      await expect(service.generateFollowupSuggestions('Test email'))
        .rejects.toThrow('Failed to generate followup suggestions: DIAL API request failed: 500 Internal Server Error - Server Error');
    });

    it('should include context in prompt when provided', async () => {
      const mockResponse = '1. Contextual response';

      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({
          choices: [{
            message: { content: mockResponse }
          }]
        })
      } as Response);

      await service.generateFollowupSuggestions('Test email', 'Important context');

      expect(mockFetch).toHaveBeenCalledWith(
        'https://ai-proxy.lab.epam.com/openai/chat/completions',
        expect.objectContaining({
          body: expect.stringContaining('Important context')
        })
      );
    });
  });

  describe('analyzeTone', () => {
    it('should analyze email tone successfully', async () => {
      const mockResponse = 'The tone is professional and courteous.';

      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({
          choices: [{
            message: { content: mockResponse }
          }]
        })
      } as Response);

      const tone = await service.analyzeTone('Thank you for your email.');

      expect(tone).toBe('The tone is professional and courteous.');
    });

    it('should handle tone analysis errors', async () => {
      mockFetch.mockResolvedValueOnce({
        ok: false,
        status: 429,
        statusText: 'Too Many Requests'
      } as Response);

      await expect(service.analyzeTone('Test email'))
        .rejects.toThrow('Failed to analyze tone');
    });
  });

  describe('summarizeEmail', () => {
    it('should summarize email successfully', async () => {
      const mockResponse = 'This email discusses project updates and timelines.';

      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({
          choices: [{
            message: { content: mockResponse }
          }]
        })
      } as Response);

      const summary = await service.summarizeEmail('Long email content about project...');

      expect(summary).toBe('This email discusses project updates and timelines.');
    });
  });

  describe('analyzeThread', () => {
    it('should analyze email thread successfully', async () => {
      const mockResponse = 'The thread discusses project status with increasing urgency.';

      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({
          choices: [{
            message: { content: mockResponse }
          }]
        })
      } as Response);

      const mockEmails = [
        { subject: 'Project Update', content: 'Initial email' },
        { subject: 'Re: Project Update', content: 'Follow up email' }
      ];

      const analysis = await service.analyzeThread({ 
        emails: mockEmails, 
        context: 'Follow-up analysis' 
      });

      expect(analysis).toBe('The thread discusses project status with increasing urgency.');
    });

    it('should handle thread analysis without context', async () => {
      const mockResponse = 'Basic thread analysis.';

      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({
          choices: [{
            message: { content: mockResponse }
          }]
        })
      } as Response);

      const analysis = await service.analyzeThread({ 
        emails: [{ subject: 'Test', content: 'Test content' }] 
      });

      expect(analysis).toBe('Basic thread analysis.');
    });
  });

  describe('analyzeSentiment', () => {
    it('should analyze sentiment as positive', async () => {
      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({
          choices: [{
            message: { content: 'positive' }
          }]
        })
      } as Response);

      const result = await service.analyzeSentiment('Thank you so much!');

      expect(result.sentiment).toBe('positive');
    });

    it('should analyze sentiment as urgent', async () => {
      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({
          choices: [{
            message: { content: 'urgent' }
          }]
        })
      } as Response);

      const result = await service.analyzeSentiment('This is urgent!');

      expect(result.sentiment).toBe('urgent');
    });

    it('should default to neutral on API error', async () => {
      mockFetch.mockRejectedValueOnce(new Error('API Error'));

      const result = await service.analyzeSentiment('Test content');

      expect(result.sentiment).toBe('neutral');
    });
  });

  describe('API Configuration', () => {
    it('should use DIAL API for EPAM endpoints', async () => {
      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({
          choices: [{ message: { content: '1. Test suggestion' } }]
        })
      } as Response);

      await service.generateFollowupSuggestions('Test');

      expect(mockFetch).toHaveBeenCalledWith(
        'https://ai-proxy.lab.epam.com/openai/chat/completions',
        expect.objectContaining({
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Api-Key': 'test-api-key'
          }
        })
      );
    });

    it('should use Azure OpenAI format for other endpoints', async () => {
      // Update configuration to use Azure endpoint
      const azureConfig = {
        ...mockConfiguration,
        llmApiEndpoint: 'https://myazure.openai.azure.com'
      };
      
      const azureService = new LlmService(azureConfig, mockRetryService);

      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({
          choices: [{ message: { content: '1. Test suggestion' } }]
        })
      } as Response);

      await azureService.generateFollowupSuggestions('Test');

      expect(mockFetch).toHaveBeenCalledWith(
        'https://myazure.openai.azure.com/openai/deployments/gpt-35-turbo/chat/completions?api-version=2023-12-01-preview',
        expect.objectContaining({
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer test-api-key'
          }
        })
      );
    });

    it('should throw error when API endpoint is not configured', async () => {
      const invalidConfig = { ...mockConfiguration, llmApiEndpoint: undefined };
      const invalidService = new LlmService(invalidConfig, mockRetryService);

      await expect(invalidService.generateFollowupSuggestions('Test'))
        .rejects.toThrow('LLM API endpoint is not configured');
    });
  });

  describe('Retry Integration', () => {
    it('should use retry service for API calls', async () => {
      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({
          choices: [{ message: { content: '1. Test' } }]
        })
      } as Response);

      await service.generateFollowupSuggestions('Test');

      expect(mockRetryService.executeWithRetry).toHaveBeenCalledWith(
        expect.any(Function),
        expect.objectContaining({
          maxAttempts: 3,
          baseDelayMs: 1000,
          maxDelayMs: 10000,
          backoffFactor: 2
        }),
        'llm-api'
      );
    });
  });
});