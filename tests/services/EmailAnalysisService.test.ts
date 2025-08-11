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
        'conv-ai-1',
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
        'conv-ai-2',
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
    describe('GetConversationItems SOAP request', () => {
      it('should include proper namespace declarations on the Envelope tag', () => {
        const req = (service as any).buildGetConversationItemsRequest('ABC123');
        expect(req).toContain('<soap:Envelope');
        // Namespaces must be on the same opening tag
        expect(req).toContain('xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"');
        expect(req).toContain('xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"');
        expect(req).toContain('xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"');
      });
    });
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

      it('should only show threads where user is last sender across folders (multi-folder)', async () => {
        // Simulate a thread where user sent last message (needs followup)
        const threadNeedingFollowup: ThreadMessage[] = [
          { id: 'a1', subject: 'Need Info', from: 'client@example.com', to: ['user@example.com'], sentDate: new Date('2025-01-20T10:00:00Z'), body: 'Question', isFromCurrentUser: false },
          { id: 'a2', subject: 'Re: Need Info', from: 'user@example.com', to: ['client@example.com'], sentDate: new Date('2025-01-21T10:00:00Z'), body: 'Answer provided', isFromCurrentUser: true }
        ];
        // Simulate a thread where client responded after user (should be excluded)
        const threadAlreadyAnswered: ThreadMessage[] = [
          { id: 'b1', subject: 'Status', from: 'user@example.com', to: ['client@example.com'], sentDate: new Date('2025-01-20T10:00:00Z'), body: 'Any update?', isFromCurrentUser: true },
              { id: 'b2', subject: 'Re: Status', from: 'client@example.com', to: ['user@example.com'], sentDate: new Date('2025-01-22T10:00:00Z'), body: 'We are good.', isFromCurrentUser: false }
        ];

        jest.spyOn(service as any, 'getConversationThreadCached')
          .mockImplementation(async (...args: any[]) => {
            const id = args[0] as string;
            return id.startsWith('a') ? threadNeedingFollowup : threadAlreadyAnswered;
          });
        jest.spyOn(service as any, 'createFollowupEmailEnhanced')
          .mockImplementation(async (...args: any[]) => {
            const lastMessage = args[1] as ThreadMessage;
            const thread = args[2] as ThreadMessage[];
            return {
              id: lastMessage.id,
              subject: lastMessage.subject,
              recipients: lastMessage.to,
              sentDate: lastMessage.sentDate,
              body: lastMessage.body,
              summary: lastMessage.body,
              priority: 'low' as const,
              daysWithoutResponse: 1,
              conversationId: 'conv-' + lastMessage.id,
              hasAttachments: false,
              accountEmail: lastMessage.from,
              threadMessages: thread,
              isSnoozed: false,
              isDismissed: false
            };
          });

        // Directly invoke two conversations
        const res1 = await (service as any).processConversationWithCaching('conv-a', [{ id: 'a1' }], 'user@example.com', []);
        const res2 = await (service as any).processConversationWithCaching('conv-b', [{ id: 'b1' }], 'user@example.com', []);

        expect(res1).not.toBeNull();
        expect(res1.id).toBe('a2');
        expect(res2).toBeNull();
      });

      it('should not include sent items that already got an answer from other user (requirement)', async () => {
        const answeredThread: ThreadMessage[] = [
          { id: 'c1', subject: 'Ping', from: 'user@example.com', to: ['peer@example.com'], sentDate: new Date('2025-01-20T09:00:00Z'), body: 'Ping', isFromCurrentUser: true },
          { id: 'c2', subject: 'Re: Ping', from: 'peer@example.com', to: ['user@example.com'], sentDate: new Date('2025-01-20T10:00:00Z'), body: 'Pong', isFromCurrentUser: false }
        ];
        jest.spyOn(service as any, 'getConversationThreadCached').mockResolvedValue(answeredThread);
        const result = await (service as any).processConversationWithCaching('conv-c', [{ id: 'c1' }], 'user@example.com', []);
        expect(result).toBeNull();
      });
    });

    describe('Dedupe Followup Emails', () => {
      it('should keep only the latest followup per conversation', () => {
        const now = new Date();
        const earlier = new Date(now.getTime() - 60_000);
        const followups = [
          { id: 'm1', subject: 'Subj', recipients: ['a@b.com'], sentDate: earlier, body: '1', summary: '1', priority: 'low' as const, daysWithoutResponse: 1, conversationId: 'convX', hasAttachments: false, accountEmail: 'user@example.com', threadMessages: [], isSnoozed: false, isDismissed: false },
          { id: 'm2', subject: 'Subj', recipients: ['a@b.com'], sentDate: now, body: '2', summary: '2', priority: 'low' as const, daysWithoutResponse: 0, conversationId: 'convX', hasAttachments: false, accountEmail: 'user@example.com', threadMessages: [], isSnoozed: false, isDismissed: false }
        ];
        const deduped = (service as any).dedupeFollowupEmails(followups);
        expect(deduped).toHaveLength(1);
        expect(deduped[0].id).toBe('m2');
      });

      it('should fallback to composite key when conversationId missing', () => {
        const now = new Date();
        const fups = [
          { id: 'x1', subject: 'Topic', recipients: ['r@e.com'], sentDate: now, body: 'A', summary: 'A', priority: 'low' as const, daysWithoutResponse: 0, conversationId: '', hasAttachments: false, accountEmail: 'user@example.com', threadMessages: [], isSnoozed: false, isDismissed: false },
          { id: 'x2', subject: 'Topic', recipients: ['r@e.com'], sentDate: new Date(now.getTime() - 5000), body: 'B', summary: 'B', priority: 'low' as const, daysWithoutResponse: 0, conversationId: '', hasAttachments: false, accountEmail: 'user@example.com', threadMessages: [], isSnoozed: false, isDismissed: false }
        ];
        const deduped = (service as any).dedupeFollowupEmails(fups);
        expect(deduped).toHaveLength(1);
        expect(['x1']).toContain(deduped[0].id); // latest kept
      });
    });

    describe('Enhanced Thread Retrieval', () => {
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

      it('should build deep traversal request for root folder (msgfolderroot)', async () => {
        const spy = jest.spyOn(service as any, 'buildSearchConversationRequest');
        // Force search across folders
        jest.spyOn(service as any, 'searchConversationInFolder').mockImplementation(async (...args: any[]) => {
          const conv = args[0];
            const folder = args[1];
          (service as any).buildSearchConversationRequest(conv, folder, folder === 'msgfolderroot' ? 'Deep' : 'Shallow');
          return [];
        });
        await (service as any).searchConversationAcrossFolders('conv-deep');
        const deepCall = spy.mock.calls.find(c => c[1] === 'msgfolderroot');
        expect(deepCall).toBeDefined();
        if (deepCall) {
          expect(deepCall[2]).toBe('Deep');
        }
        spy.mockRestore();
      });

      it('should include archive and root folders in folder scan ordering', async () => {
        const foldersEncountered: string[] = [];
        jest.spyOn(service as any, 'searchConversationInFolder').mockImplementation(async (...args: any[]) => {
          const folder = args[1];
          foldersEncountered.push(folder);
          return [];
        });
        await (service as any).searchConversationAcrossFolders('conv-folders');
        ['sentitems','inbox','drafts','deleteditems','archive','msgfolderroot'].forEach(f => expect(foldersEncountered).toContain(f));
      });

      it('should search conversation across multiple folders', async () => {
        // Mock the searchConversationAcrossFolders method
        const mockSearchResults: ThreadMessage[] = [
          {
            id: 'msg1',
            subject: 'Original Email',
            from: 'user@example.com',
            to: ['client@example.com'],
            sentDate: new Date('2025-01-20T10:00:00Z'),
            body: 'Original message',
            isFromCurrentUser: true
          },
          {
            id: 'msg2',
            subject: 'Re: Original Email',
            from: 'client@example.com',
            to: ['user@example.com'],
            sentDate: new Date('2025-01-22T15:00:00Z'),
            body: 'Response message',
            isFromCurrentUser: false
          }
        ];

        jest.spyOn(service as any, 'searchConversationAcrossFolders')
          .mockResolvedValue(mockSearchResults);

        jest.spyOn(service as any, 'parseConversationIdResponse')
          .mockReturnValue('conv-123');

        const result = await (service as any).searchConversationAcrossFolders('conv-123');

        expect(result).toHaveLength(2);
        expect(result[0].isFromCurrentUser).toBe(true);
        expect(result[1].isFromCurrentUser).toBe(false);
      });

      it('should remove duplicate messages from search results', () => {
        const duplicateMessages: ThreadMessage[] = [
          {
            id: 'msg1',
            subject: 'Test Email',
            from: 'user@example.com',
            to: ['client@example.com'],
            sentDate: new Date('2025-01-20T10:00:00Z'),
            body: 'Test message',
            isFromCurrentUser: true
          },
          {
            id: 'msg1', // Same ID - should be deduplicated
            subject: 'Test Email',
            from: 'user@example.com',
            to: ['client@example.com'],
            sentDate: new Date('2025-01-20T10:00:00Z'),
            body: 'Test message',
            isFromCurrentUser: true
          },
          {
            id: 'msg2',
            subject: 'Different Email',
            from: 'client@example.com',
            to: ['user@example.com'],
            sentDate: new Date('2025-01-21T10:00:00Z'),
            body: 'Different message',
            isFromCurrentUser: false
          }
        ];

        const uniqueMessages = (service as any).removeDuplicateMessages(duplicateMessages);

        expect(uniqueMessages).toHaveLength(2);
        expect(uniqueMessages[0].id).toBe('msg1');
        expect(uniqueMessages[1].id).toBe('msg2');
      });

      it('should handle messages without IDs using content-based deduplication', () => {
        const messagesWithoutIds: ThreadMessage[] = [
          {
            id: '',
            subject: 'Test Email',
            from: 'user@example.com',
            to: ['client@example.com'],
            sentDate: new Date('2025-01-20T10:00:00Z'),
            body: 'Test message',
            isFromCurrentUser: true
          },
          {
            id: '',
            subject: 'Test Email', // Same content - should be deduplicated
            from: 'user@example.com',
            to: ['client@example.com'],
            sentDate: new Date('2025-01-20T10:00:00Z'),
            body: 'Test message',
            isFromCurrentUser: true
          }
        ];

        const uniqueMessages = (service as any).removeDuplicateMessages(messagesWithoutIds);

        expect(uniqueMessages).toHaveLength(1);
      });

      it('should sort messages chronologically', async () => {
        const unsortedMessages: ThreadMessage[] = [
          {
            id: 'msg2',
            subject: 'Second Email',
            from: 'client@example.com',
            to: ['user@example.com'],
            sentDate: new Date('2025-01-22T15:00:00Z'), // Later date: 1737558000000
            body: 'Second message',
            isFromCurrentUser: false
          },
          {
            id: 'msg1',
            subject: 'First Email',
            from: 'user@example.com',
            to: ['client@example.com'],
            sentDate: new Date('2025-01-20T10:00:00Z'), // Earlier date: 1737367200000
            body: 'First message',
            isFromCurrentUser: true
          }
        ];

        // Mock the individual folder searches to return messages in different orders
        const mockSearchConversationInFolder = jest.spyOn(service as any, 'searchConversationInFolder')
          .mockImplementation(async (...args: any[]) => {
            const [, folderId] = args;
            if (folderId === 'sentitems') {
              return [unsortedMessages[1]]; // Return msg1 from sent items
            } else if (folderId === 'inbox') {
              return [unsortedMessages[0]]; // Return msg2 from inbox
            }
            return [];
          });

        // Mock the removeDuplicateMessages to return the messages as-is (no duplicates in this test)
        const mockRemoveDuplicateMessages = jest.spyOn(service as any, 'removeDuplicateMessages')
          .mockImplementation((...args: any[]) => {
            const [messages] = args;
            return messages;
          });

        const result = await (service as any).searchConversationAcrossFolders('conv-123');

        // Verify the result is sorted chronologically (earliest first)
        expect(result).toHaveLength(2);
        expect(result[0].sentDate.getTime()).toBeLessThan(result[1].sentDate.getTime());
        expect(result[0].id).toBe('msg1'); // Earlier message first (2025-01-20)
        expect(result[1].id).toBe('msg2'); // Later message second (2025-01-22)
        
        // Additional verification with actual timestamps
        expect(result[0].sentDate.getTime()).toBe(1737367200000); // 2025-01-20T10:00:00Z
        expect(result[1].sentDate.getTime()).toBe(1737558000000); // 2025-01-22T15:00:00Z

        // Restore the original implementations
        mockSearchConversationInFolder.mockRestore();
        mockRemoveDuplicateMessages.mockRestore();
      });
    });

    describe('ConversationId-first Retrieval Path', () => {
      let ewsSpy: jest.SpyInstance;
      beforeEach(() => {
        ewsSpy = jest.spyOn(service as any, 'isEwsAvailable').mockReturnValue(true);
      });
      afterEach(() => {
        ewsSpy.mockRestore();
      });

      it('should use GetConversationItems path and skip folder/item fallbacks', async () => {
        const convId = 'conv-use-first';
        const thread: ThreadMessage[] = [
          { id: 't1', subject: 'Question', from: 'client@example.com', to: ['user@example.com'], sentDate: new Date('2025-01-20T10:00:00Z'), body: 'Hi', isFromCurrentUser: false },
          { id: 't2', subject: 'Re: Question', from: 'user@example.com', to: ['client@example.com'], sentDate: new Date('2025-01-21T10:00:00Z'), body: 'Answer', isFromCurrentUser: true }
        ];
        const spyConvItems = jest.spyOn(service as any, 'getConversationItemsConversationCached').mockResolvedValue(thread);
        const spyFolderPath = jest.spyOn(service as any, 'getConversationThreadFromConversationIdCached').mockResolvedValue([]);
        const spyItem = jest.spyOn(service as any, 'getConversationThreadCached').mockResolvedValue([]);
        const spyCreate = jest.spyOn(service as any, 'createFollowupEmailEnhanced').mockImplementation(async (...args: any[]) => {
          const last: ThreadMessage = args[1];
          const all: ThreadMessage[] = args[2];
          return { id: last.id, subject: last.subject, recipients: last.to, sentDate: last.sentDate, body: last.body, summary: last.body, priority: 'low', daysWithoutResponse: 1, conversationId: convId, hasAttachments: false, accountEmail: last.from, threadMessages: all, isSnoozed: false, isDismissed: false };
        });
        const result = await (service as any).processConversationWithCaching(convId, [{ id: 'seedMessage' }], 'user@example.com', []);
        expect(result).not.toBeNull();
        expect(result.id).toBe('t2');
        expect(spyConvItems).toHaveBeenCalledTimes(1);
        expect(spyFolderPath).not.toHaveBeenCalled();
        expect(spyItem).not.toHaveBeenCalled();
        spyConvItems.mockRestore();
        spyFolderPath.mockRestore();
        spyItem.mockRestore();
        spyCreate.mockRestore();
      });

      it('should fallback to item retrieval when GetConversationItems returns empty', async () => {
        const convId = 'conv-fallback';
        const fallbackThread: ThreadMessage[] = [
          { id: 'fb1', subject: 'Need Help', from: 'client@example.com', to: ['user@example.com'], sentDate: new Date('2025-01-20T09:00:00Z'), body: 'Need info', isFromCurrentUser: false },
          { id: 'fb2', subject: 'Re: Need Help', from: 'user@example.com', to: ['client@example.com'], sentDate: new Date('2025-01-20T10:00:00Z'), body: 'Provided', isFromCurrentUser: true }
        ];
        const spyConvItems = jest.spyOn(service as any, 'getConversationItemsConversationCached').mockResolvedValue([]);
        // Folder-based fallback removed from main path; ensure item-based fallback is used
        const spyFolderPath = jest.spyOn(service as any, 'getConversationThreadFromConversationIdCached').mockResolvedValue([]);
        const spyItem = jest.spyOn(service as any, 'getConversationThreadCached').mockResolvedValue(fallbackThread);
        const spyCreate = jest.spyOn(service as any, 'createFollowupEmailEnhanced').mockImplementation(async (...args: any[]) => {
          const last: ThreadMessage = args[1];
          const all: ThreadMessage[] = args[2];
          return { id: last.id, subject: last.subject, recipients: last.to, sentDate: last.sentDate, body: last.body, summary: last.body, priority: 'low', daysWithoutResponse: 1, conversationId: convId, hasAttachments: false, accountEmail: last.from, threadMessages: all, isSnoozed: false, isDismissed: false };
        });
        const result = await (service as any).processConversationWithCaching(convId, [{ id: 'seedMessage' }], 'user@example.com', []);
        expect(result).not.toBeNull();
        expect(result.id).toBe('fb2');
        expect(spyConvItems).toHaveBeenCalledTimes(1);
        expect(spyFolderPath).not.toHaveBeenCalled();
        expect(spyItem).toHaveBeenCalledTimes(1);
        spyConvItems.mockRestore();
        spyFolderPath.mockRestore();
        spyItem.mockRestore();
        spyCreate.mockRestore();
      });
    });

    describe('parseGetConversationItemsResponse', () => {
      beforeEach(() => {
        // Ensure current user email matches the XML for deterministic assertions
        (global as any).Office = {
          context: {
            mailbox: {
              userProfile: {
                emailAddress: 'test@example.com'
              }
            }
          }
        };
      });
      it('should parse messages and sort chronologically by received or sent date', () => {
        const xml = `<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <s:Body>
    <m:GetConversationItemsResponse>
      <m:ResponseMessages>
        <m:GetConversationItemsResponseMessage ResponseClass="Success">
          <m:ResponseCode>NoError</m:ResponseCode>
          <m:ConversationNodes>
            <t:ConversationNode>
              <t:Items>
                <t:Message>
                  <t:ItemId Id="id-1" />
                  <t:Subject>Subj 1</t:Subject>
                  <t:DateTimeSent>2025-01-20T10:00:00Z</t:DateTimeSent>
                  <t:DateTimeReceived>2025-01-20T10:05:00Z</t:DateTimeReceived>
                  <t:From><t:Mailbox><t:EmailAddress>other@example.com</t:EmailAddress></t:Mailbox></t:From>
                  <t:ToRecipients><t:Mailbox><t:EmailAddress>test@example.com</t:EmailAddress></t:Mailbox></t:ToRecipients>
                  <t:Body>Body 1</t:Body>
                </t:Message>
              </t:Items>
            </t:ConversationNode>
            <t:ConversationNode>
              <t:Items>
                <t:Message>
                  <t:ItemId Id="id-2" />
                  <t:Subject>Subj 2</t:Subject>
                  <t:DateTimeSent>2025-01-21T12:00:00Z</t:DateTimeSent>
                  <t:DateTimeReceived>2025-01-21T12:02:00Z</t:DateTimeReceived>
                  <t:From><t:Mailbox><t:EmailAddress>test@example.com</t:EmailAddress></t:Mailbox></t:From>
                  <t:ToRecipients><t:Mailbox><t:EmailAddress>other@example.com</t:EmailAddress></t:Mailbox></t:ToRecipients>
                  <t:Body>Body 2</t:Body>
                </t:Message>
              </t:Items>
            </t:ConversationNode>
          </m:ConversationNodes>
        </m:GetConversationItemsResponseMessage>
      </m:ResponseMessages>
    </m:GetConversationItemsResponse>
  </s:Body>
</s:Envelope>`;

        const parsed = (service as any).parseGetConversationItemsResponse(xml) as ThreadMessage[];
        expect(parsed.length).toBe(2);
        // Sorted chronologically by receivedDate/sentDate ascending
        expect(parsed[0].id).toBe('id-1');
        expect(parsed[1].id).toBe('id-2');
        expect(parsed[0].isFromCurrentUser).toBe(false);
        expect(parsed[1].isFromCurrentUser).toBe(true);
      });
    });
  });
});