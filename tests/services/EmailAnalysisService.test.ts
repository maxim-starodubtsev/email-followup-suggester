import { EmailAnalysisService } from '../../src/services/EmailAnalysisService';
import { LlmService } from '../../src/services/LlmService';
import { ThreadMessage } from '../../src/models/FollowupEmail';

describe('EmailAnalysisService', () => {
  let service: EmailAnalysisService;
  let mockLlmService: jest.Mocked<LlmService>;

  beforeEach(() => {
    service = new EmailAnalysisService();
    mockLlmService = {
      analyzeThread: jest.fn(),
      analyzeSentiment: jest.fn().mockResolvedValue({ sentiment: 'neutral' }),
      getAvailableModels: jest.fn(),
      checkModelLimits: jest.fn()
    } as any;
    service.setLlmService(mockLlmService);
  });

  describe('Priority Calculation', () => {
    it('should calculate high priority for emails older than 7 days', () => {
      const priority = (service as any).calculateEnhancedPriority(8, false, 'neutral');
      expect(priority).toBe('high');
    });

    it('should calculate medium priority for emails 3-6 days old', () => {
      const priority = (service as any).calculateEnhancedPriority(5, false, 'neutral');
      expect(priority).toBe('medium');
    });

    it('should calculate low priority for emails less than 3 days old', () => {
      const priority = (service as any).calculateEnhancedPriority(2, false, 'neutral');
      expect(priority).toBe('low');
    });

    it('should boost priority for threaded conversations', () => {
      const lowPriority = (service as any).calculateEnhancedPriority(2, true, 'neutral');
      expect(lowPriority).toBe('medium');

      const mediumPriority = (service as any).calculateEnhancedPriority(5, true, 'neutral');
      expect(mediumPriority).toBe('high');
    });

    it('should boost priority for urgent sentiment', () => {
      const urgentPriority = (service as any).calculateEnhancedPriority(2, false, 'urgent');
      expect(urgentPriority).toBe('high');
    });

    it('should boost priority for negative sentiment', () => {
      const negativePriority = (service as any).calculateEnhancedPriority(1, false, 'negative');
      expect(negativePriority).toBe('medium');
    });
  });

  describe('Summary Generation', () => {
    it('should generate summary from email body', () => {
      const body = 'This is a test email about project updates and timeline discussions.';
      const subject = 'Project Update';
      
      const summary = (service as any).generateSummary(body, subject);
      expect(summary).toBe(body);
    });

    it('should truncate long email bodies', () => {
      const longBody = 'A'.repeat(200);
      const subject = 'Long Email';
      
      const summary = (service as any).generateSummary(longBody, subject);
      expect(summary.length).toBeLessThanOrEqual(153); // 150 + '...'
      expect(summary.endsWith('...')).toBe(true);
    });

    it('should use subject when body is empty', () => {
      const summary = (service as any).generateSummary('', 'Test Subject');
      expect(summary).toBe('Email about: Test Subject');
    });

    it('should remove HTML tags from body', () => {
      const htmlBody = '<p>This is a <strong>test</strong> email.</p>';
      const subject = 'HTML Email';
      
      const summary = (service as any).generateSummary(htmlBody, subject);
      expect(summary).toBe('This is a test email.');
    });
  });

  describe('Cache Operations', () => {
    it('should get cache statistics', () => {
      const stats = service.getCacheStats();
      
      expect(stats).toHaveProperty('totalEntries');
      expect(stats).toHaveProperty('hitRate');
      expect(stats).toHaveProperty('totalMemoryUsage');
    });

    it('should clear cache with pattern', () => {
      const clearedCount = service.clearCache('email:.*');
      expect(typeof clearedCount).toBe('number');
    });

    it('should get memory pressure', () => {
      const pressure = service.getCacheMemoryPressure();
      expect(typeof pressure).toBe('number');
      expect(pressure).toBeGreaterThanOrEqual(0);
    });
  });

  describe('Bulk Operations', () => {
    it('should bulk snooze emails', async () => {
      const emailIds = ['email1', 'email2', 'email3'];
      const minutes = 30;

      await service.bulkSnoozeEmails(emailIds, minutes);

      emailIds.forEach(emailId => {
        expect((service as any).isEmailSnoozed(emailId)).toBe(true);
      });
    });

    it('should bulk dismiss emails', async () => {
      const emailIds = ['email1', 'email2', 'email3'];

      await service.bulkDismissEmails(emailIds);

      emailIds.forEach(emailId => {
        expect((service as any).isEmailDismissed(emailId)).toBe(true);
      });
    });
  });

  describe('Sentiment Analysis', () => {
    it('should analyze email sentiment', async () => {
      const emailBody = 'This is urgent! Please respond ASAP.';
      
      const sentiment = await service.analyzeEmailSentiment(emailBody);
      
      expect(['positive', 'neutral', 'negative', 'urgent']).toContain(sentiment);
    });

    it('should detect urgent keywords', () => {
      const urgentBody = 'This is an urgent matter that needs immediate attention.';
      const sentiment = (service as any).analyzeSentimentBasic(urgentBody);
      
      expect(sentiment).toBe('urgent');
    });

    it('should detect negative sentiment', () => {
      const negativeBody = 'There is a problem with the system that failed to work.';
      const sentiment = (service as any).analyzeSentimentBasic(negativeBody);
      
      expect(sentiment).toBe('negative');
    });

    it('should detect positive sentiment', () => {
      const positiveBody = 'Thanks for the excellent work! This is perfect.';
      const sentiment = (service as any).analyzeSentimentBasic(positiveBody);
      
      expect(sentiment).toBe('positive');
    });

    it('should default to neutral sentiment', () => {
      const neutralBody = 'Here is the report you requested.';
      const sentiment = (service as any).analyzeSentimentBasic(neutralBody);
      
      expect(sentiment).toBe('neutral');
    });
  });

  describe('LLM Integration', () => {
    it('should use LLM summary when available', async () => {
      mockLlmService.analyzeThread.mockResolvedValue('AI-generated summary');
      mockLlmService.generateFollowupSuggestions = jest.fn().mockResolvedValue(['AI suggestion']);

      const mockLastMessage: ThreadMessage = {
        id: 'msg1',
        subject: 'Test Subject',
        from: 'test@example.com',
        to: ['recipient@example.com'],
        sentDate: new Date('2025-01-25T10:00:00Z'),
        body: 'Test body',
        isFromCurrentUser: true
      };

      const followupEmail = await (service as any).createFollowupEmailEnhanced(
        mockLastMessage,
        [mockLastMessage],
        'test@example.com'
      );

      expect(followupEmail.summary).toBe('AI-generated summary');
      expect(followupEmail.llmSuggestion).toBe('AI suggestion');
      expect(followupEmail.llmSummary).toBe('AI-generated summary');
    });

    it('should fallback to basic summary when LLM fails', async () => {
      mockLlmService.analyzeThread.mockRejectedValue(new Error('API Error'));

      const mockLastMessage: ThreadMessage = {
        id: 'msg1',
        subject: 'Test Subject',
        from: 'test@example.com',
        to: ['recipient@example.com'],
        sentDate: new Date('2025-01-25T10:00:00Z'),
        body: 'This is a test email body for fallback testing.',
        isFromCurrentUser: true
      };

      const followupEmail = await (service as any).createFollowupEmailEnhanced(
        mockLastMessage,
        [mockLastMessage],
        'test@example.com'
      );

      expect(followupEmail.summary).toBe('This is a test email body for fallback testing.');
      expect(followupEmail.llmSuggestion).toBeUndefined();
      expect(followupEmail.llmSummary).toBeUndefined();
    });
  });
});