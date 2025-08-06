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

  describe('Thread Analysis and Response Detection (Bug Fixes)', () => {
    describe('getLastMessageInThread', () => {
      it('should return the chronologically latest message', () => {
        const messages: ThreadMessage[] = [
          {
            id: 'msg1',
            subject: 'Original Email',
            from: 'user@example.com',
            to: ['recipient@example.com'],
            sentDate: new Date('2025-01-20T10:00:00Z'),
            body: 'Original message',
            isFromCurrentUser: true
          },
          {
            id: 'msg2',
            subject: 'Re: Original Email',
            from: 'recipient@example.com',
            to: ['user@example.com'],
            sentDate: new Date('2025-01-22T15:00:00Z'),
            body: 'Response message',
            isFromCurrentUser: false
          },
          {
            id: 'msg3',
            subject: 'Re: Original Email',
            from: 'user@example.com',
            to: ['recipient@example.com'],
            sentDate: new Date('2025-01-21T12:00:00Z'),
            body: 'Follow-up message',
            isFromCurrentUser: true
          }
        ];

        const lastMessage = (service as any).getLastMessageInThread(messages);
        
        expect(lastMessage).not.toBeNull();
        expect(lastMessage.id).toBe('msg2'); // Latest by date
        expect(lastMessage.sentDate).toEqual(new Date('2025-01-22T15:00:00Z'));
        expect(lastMessage.isFromCurrentUser).toBe(false);
      });

      it('should return null for empty thread', () => {
        const lastMessage = (service as any).getLastMessageInThread([]);
        expect(lastMessage).toBeNull();
      });

      it('should handle single message thread', () => {
        const messages: ThreadMessage[] = [
          {
            id: 'msg1',
            subject: 'Single Email',
            from: 'user@example.com',
            to: ['recipient@example.com'],
            sentDate: new Date('2025-01-20T10:00:00Z'),
            body: 'Single message',
            isFromCurrentUser: true
          }
        ];

        const lastMessage = (service as any).getLastMessageInThread(messages);
        
        expect(lastMessage).not.toBeNull();
        expect(lastMessage.id).toBe('msg1');
      });
    });

    describe('checkForResponseInThread', () => {
      it('should detect response after last sent message (Bug Fix Case 1)', () => {
        const messages: ThreadMessage[] = [
          {
            id: 'msg1',
            subject: 'Original Email',
            from: 'user@example.com',
            to: ['recipient@example.com'],
            sentDate: new Date('2025-01-20T10:00:00Z'),
            body: 'Original message',
            isFromCurrentUser: true
          },
          {
            id: 'msg2',
            subject: 'Re: Original Email',
            from: 'recipient@example.com',
            to: ['user@example.com'],
            sentDate: new Date('2025-01-22T15:00:00Z'),
            body: 'Response message',
            isFromCurrentUser: false
          }
        ];

        const lastSentDate = new Date('2025-01-20T10:00:00Z');
        const hasResponse = (service as any).checkForResponseInThread(messages, lastSentDate);
        
        expect(hasResponse).toBe(true);
      });

      it('should NOT detect response when last message is from current user (Bug Fix Case 2)', () => {
        const messages: ThreadMessage[] = [
          {
            id: 'msg1',
            subject: 'Original Email',
            from: 'recipient@example.com',
            to: ['user@example.com'],
            sentDate: new Date('2025-01-20T10:00:00Z'),
            body: 'Original message',
            isFromCurrentUser: false
          },
          {
            id: 'msg2',
            subject: 'Re: Original Email',
            from: 'user@example.com',
            to: ['recipient@example.com'],
            sentDate: new Date('2025-01-22T15:00:00Z'),
            body: 'My response message',
            isFromCurrentUser: true
          }
        ];

        const lastSentDate = new Date('2025-01-22T15:00:00Z');
        const hasResponse = (service as any).checkForResponseInThread(messages, lastSentDate);
        
        expect(hasResponse).toBe(false);
      });

      it('should handle complex thread with multiple back-and-forth messages', () => {
        const messages: ThreadMessage[] = [
          {
            id: 'msg1',
            subject: 'Original Email',
            from: 'user@example.com',
            to: ['recipient@example.com'],
            sentDate: new Date('2025-01-20T10:00:00Z'),
            body: 'Original message',
            isFromCurrentUser: true
          },
          {
            id: 'msg2',
            subject: 'Re: Original Email',
            from: 'recipient@example.com',
            to: ['user@example.com'],
            sentDate: new Date('2025-01-21T09:00:00Z'),
            body: 'First response',
            isFromCurrentUser: false
          },
          {
            id: 'msg3',
            subject: 'Re: Original Email',
            from: 'user@example.com',
            to: ['recipient@example.com'],
            sentDate: new Date('2025-01-22T14:00:00Z'),
            body: 'My follow-up',
            isFromCurrentUser: true
          },
          {
            id: 'msg4',
            subject: 'Re: Original Email',
            from: 'recipient@example.com',
            to: ['user@example.com'],
            sentDate: new Date('2025-01-23T11:00:00Z'),
            body: 'Final response',
            isFromCurrentUser: false
          }
        ];

        // Check if there's a response after the user's last message (msg3)
        const lastUserSentDate = new Date('2025-01-22T14:00:00Z');
        const hasResponse = (service as any).checkForResponseInThread(messages, lastUserSentDate);
        
        expect(hasResponse).toBe(true); // msg4 is a response after msg3
      });

      it('should not detect responses from current user as external responses', () => {
        const messages: ThreadMessage[] = [
          {
            id: 'msg1',
            subject: 'Original Email',
            from: 'user@example.com',
            to: ['recipient@example.com'],
            sentDate: new Date('2025-01-20T10:00:00Z'),
            body: 'Original message',
            isFromCurrentUser: true
          },
          {
            id: 'msg2',
            subject: 'Re: Original Email',
            from: 'user@example.com',
            to: ['recipient@example.com'],
            sentDate: new Date('2025-01-22T15:00:00Z'),
            body: 'Another message from me',
            isFromCurrentUser: true
          }
        ];

        const lastSentDate = new Date('2025-01-20T10:00:00Z');
        const hasResponse = (service as any).checkForResponseInThread(messages, lastSentDate);
        
        expect(hasResponse).toBe(false); // Only messages from current user after the last sent date
      });
    });

    describe('parseMessageElement - Case-insensitive email comparison', () => {
      beforeEach(() => {
        // Mock Office context for these tests
        (global as any).Office = {
          context: {
            mailbox: {
              userProfile: {
                emailAddress: 'User@Example.com' // Mixed case to test normalization
              }
            }
          }
        };
      });

      it('should correctly identify current user messages with case variations', () => {
        const mockElement = document.createElement('div');
        mockElement.innerHTML = `
          <t:ItemId Id="test-id" />
          <t:Subject>Test Subject</t:Subject>
          <t:DateTimeSent>2025-01-20T10:00:00Z</t:DateTimeSent>
          <t:From>
            <t:Mailbox>
              <t:EmailAddress>user@example.com</t:EmailAddress>
            </t:Mailbox>
          </t:From>
          <t:ToRecipients>
            <t:Mailbox>
              <t:EmailAddress>recipient@example.com</t:EmailAddress>
            </t:Mailbox>
          </t:ToRecipients>
          <t:Body>Test body</t:Body>
        `;

        const currentUserEmail = 'User@Example.com';
        const message = (service as any).parseMessageElement(mockElement, currentUserEmail);
        
        expect(message).not.toBeNull();
        expect(message.isFromCurrentUser).toBe(true); // Should be true despite case difference
        expect(message.from).toBe('user@example.com');
      });

      it('should correctly identify external messages', () => {
        const mockElement = document.createElement('div');
        mockElement.innerHTML = `
          <t:ItemId Id="test-id" />
          <t:Subject>Test Subject</t:Subject>
          <t:DateTimeSent>2025-01-20T10:00:00Z</t:DateTimeSent>
          <t:From>
            <t:Mailbox>
              <t:EmailAddress>OTHER@EXAMPLE.COM</t:EmailAddress>
            </t:Mailbox>
          </t:From>
          <t:ToRecipients>
            <t:Mailbox>
              <t:EmailAddress>user@example.com</t:EmailAddress>
            </t:Mailbox>
          </t:ToRecipients>
          <t:Body>Test body</t:Body>
        `;

        const currentUserEmail = 'user@example.com';
        const message = (service as any).parseMessageElement(mockElement, currentUserEmail);
        
        expect(message).not.toBeNull();
        expect(message.isFromCurrentUser).toBe(false);
        expect(message.from).toBe('OTHER@EXAMPLE.COM');
      });
    });

    describe('Conversation Processing Integration Tests', () => {
      beforeEach(() => {
        // Mock Office context
        (global as any).Office = {
          context: {
            mailbox: {
              userProfile: {
                emailAddress: 'user@example.com'
              }
            }
          }
        };
      });

      it('should filter out conversations where last message has a response (Bug Case 1)', async () => {
        // Mock the thread retrieval to return a conversation with a response
        const mockThreadMessages: ThreadMessage[] = [
          {
            id: 'msg1',
            subject: 'Project Update',
            from: 'user@example.com',
            to: ['client@example.com'],
            sentDate: new Date('2025-01-20T10:00:00Z'),
            body: 'Please review the project update',
            isFromCurrentUser: true
          },
          {
            id: 'msg2',
            subject: 'Re: Project Update',
            from: 'client@example.com',
            to: ['user@example.com'],
            sentDate: new Date('2025-01-22T15:00:00Z'),
            body: 'Thanks for the update. Looks good!',
            isFromCurrentUser: false
          }
        ];

        // Spy on the cached method to return our mock data
        jest.spyOn(service as any, 'getConversationThreadCached')
          .mockResolvedValue(mockThreadMessages);

        const result = await (service as any).processConversationWithCaching(
          'conv1',
          [{ id: 'msg1' }],
          'user@example.com',
          []
        );

        expect(result).toBeNull(); // Should be filtered out because there's a response
      });

      it('should include conversations where current user sent last message without response (Bug Case 2)', async () => {
        // Mock the thread retrieval to return a conversation where user sent the last message
        const mockThreadMessages: ThreadMessage[] = [
          {
            id: 'msg1',
            subject: 'Project Question',
            from: 'client@example.com',
            to: ['user@example.com'],
            sentDate: new Date('2025-01-20T10:00:00Z'),
            body: 'I have a question about the project',
            isFromCurrentUser: false
          },
          {
            id: 'msg2',
            subject: 'Re: Project Question',
            from: 'user@example.com',
            to: ['client@example.com'],
            sentDate: new Date('2025-01-22T15:00:00Z'),
            body: 'Here is my response to your question',
            isFromCurrentUser: true
          }
        ];

        // Spy on the cached method to return our mock data
        jest.spyOn(service as any, 'getConversationThreadCached')
          .mockResolvedValue(mockThreadMessages);

        jest.spyOn(service as any, 'createFollowupEmailEnhanced')
          .mockResolvedValue({
            id: 'msg2',
            subject: 'Re: Project Question',
            recipients: ['client@example.com'],
            sentDate: new Date('2025-01-22T15:00:00Z'),
            body: 'Here is my response to your question',
            summary: 'Response to project question',
            priority: 'medium' as const,
            daysWithoutResponse: 3,
            conversationId: 'conv1',
            hasAttachments: false,
            accountEmail: 'user@example.com',
            threadMessages: mockThreadMessages,
            isSnoozed: false,
            isDismissed: false
          });

        const result = await (service as any).processConversationWithCaching(
          'conv1',
          [{ id: 'msg1' }],
          'user@example.com',
          []
        );

        expect(result).not.toBeNull(); // Should be included for followup
        expect(result.id).toBe('msg2');
      });

      it('should handle conversations with mixed case email addresses', async () => {
        const mockThreadMessages: ThreadMessage[] = [
          {
            id: 'msg1',
            subject: 'Test Email',
            from: 'USER@EXAMPLE.COM', // Different case
            to: ['client@example.com'],
            sentDate: new Date('2025-01-20T10:00:00Z'),
            body: 'Test message',
            isFromCurrentUser: true // Should be correctly identified
          }
        ];

        jest.spyOn(service as any, 'getConversationThreadCached')
          .mockResolvedValue(mockThreadMessages);

        jest.spyOn(service as any, 'createFollowupEmailEnhanced')
          .mockResolvedValue({
            id: 'msg1',
            subject: 'Test Email',
            recipients: ['client@example.com'],
            sentDate: new Date('2025-01-20T10:00:00Z'),
            body: 'Test message',
            summary: 'Test message',
            priority: 'low' as const,
            daysWithoutResponse: 16,
            conversationId: 'conv1',
            hasAttachments: false,
            accountEmail: 'USER@EXAMPLE.COM',
            threadMessages: mockThreadMessages,
            isSnoozed: false,
            isDismissed: false
          });

        const result = await (service as any).processConversationWithCaching(
          'conv1',
          [{ id: 'msg1' }],
          'user@example.com', // Different case
          []
        );

        expect(result).not.toBeNull();
        expect(result.accountEmail).toBe('USER@EXAMPLE.COM');
      });
    });
  });
});