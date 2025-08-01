import { FollowupEmail, ThreadMessage } from '../models/FollowupEmail';
import { LlmService } from './LlmService';
import { BatchProcessor, BatchProcessingOptions, BatchResult } from './BatchProcessor';
import { CacheService, ICacheService } from './CacheService';

interface RetryOptions {
    maxRetries: number;
    baseDelay: number;
    maxDelay: number;
}

interface AnalyticsEvent {
    type: 'email_analyzed' | 'batch_processed' | 'cache_hit' | 'cache_miss' | 'error_occurred' | 'retry_attempted';
    timestamp: number;
    details: any;
}

export class EmailAnalysisService {
    private readonly SUMMARY_MAX_LENGTH = 150;
    private readonly BATCH_SIZE = 10;
    private readonly DEFAULT_RETRY_OPTIONS: RetryOptions = {
        maxRetries: 3,
        baseDelay: 1000,
        maxDelay: 10000
    };

    private llmService?: LlmService;
    private snoozedEmails: Map<string, Date> = new Map();
    private dismissedEmails: Set<string> = new Set();
    
    // Enhanced caching with CacheService
    private cacheService: ICacheService;
    
    // Analytics tracking
    private analyticsEvents: AnalyticsEvent[] = [];
    private performanceMetrics = {
        totalAnalyzed: 0,
        cacheHits: 0,
        cacheMisses: 0,
        errors: 0,
        averageProcessingTime: 0
    };

    private batchProcessor: BatchProcessor;

    constructor(cacheService?: ICacheService) {
        this.batchProcessor = new BatchProcessor();
        
        // Initialize enhanced caching system
        this.cacheService = cacheService || new CacheService({
            defaultTtl: 30 * 60 * 1000, // 30 minutes for email analysis
            maxMemoryUsage: 25 * 1024 * 1024, // 25MB for email cache
            maxEntries: 5000,
            evictionPolicy: 'lru',
            enableContentHashing: true,
            enableStatistics: true
        });
    }

    public setLlmService(llmService: LlmService): void {
        this.llmService = llmService;
    }

    public async analyzeEmails(emailCount: number, daysBack: number, selectedAccounts: string[]): Promise<FollowupEmail[]> {
        const startTime = Date.now();
        this.trackAnalyticsEvent('batch_processed', { emailCount, daysBack, selectedAccounts });

        try {
            const currentUserEmail = Office.context.mailbox.userProfile.emailAddress;
            const cutoffDate = new Date();
            cutoffDate.setDate(cutoffDate.getDate() - daysBack);

            // Get sent emails with enhanced caching and error retry
            const sentEmails = await this.getSentEmailsWithCaching(emailCount, cutoffDate);

            // Group emails by conversation ID
            const conversationGroups = this.groupEmailsByConversation(sentEmails);
            const conversationIds = Array.from(conversationGroups.keys());

            // Use BatchProcessor for improved performance and error handling
            const batchOptions: Partial<BatchProcessingOptions> = {
                batchSize: this.BATCH_SIZE,
                maxConcurrentBatches: 3,
                retryOptions: this.DEFAULT_RETRY_OPTIONS,
                onProgress: (current, total, currentBatch, totalBatches) => {
                    this.trackAnalyticsEvent('batch_processed', {
                        progress: { current, total, currentBatch, totalBatches }
                    });
                },
                onBatchComplete: (batchIndex, results, errors) => {
                    this.trackAnalyticsEvent('batch_processed', {
                        batch: { index: batchIndex, successCount: results.length, errorCount: errors.length }
                    });
                },
                onBatchError: (batchIndex, error) => {
                    this.trackAnalyticsEvent('error_occurred', {
                        batchIndex,
                        error: error.message
                    });
                }
            };

            // Process conversations using the BatchProcessor
            const batchResult: BatchResult<FollowupEmail | null> = await this.batchProcessor.processBatch(
                conversationIds,
                async (conversationId: string) => {
                    return await this.processConversationWithCaching(
                        conversationId,
                        conversationGroups.get(conversationId)!,
                        currentUserEmail,
                        selectedAccounts
                    );
                },
                batchOptions
            );

            // Filter out null results and collect successful followup emails
            const followupEmails: FollowupEmail[] = batchResult.results
                .filter((result): result is FollowupEmail => result !== null);

            // Log batch processing results with cache stats
            const cacheStats = this.cacheService.getStats();
            this.trackAnalyticsEvent('batch_processed', {
                totalConversations: conversationIds.length,
                successfullyProcessed: batchResult.totalProcessed,
                followupEmailsFound: followupEmails.length,
                errors: batchResult.errors.length,
                processingTime: batchResult.processingTime,
                cancelled: batchResult.cancelled,
                cacheHitRate: cacheStats.hitRate,
                cacheMemoryUsage: cacheStats.totalMemoryUsage
            });

            // Handle any errors from batch processing
            if (batchResult.errors.length > 0) {
                console.warn(`Batch processing completed with ${batchResult.errors.length} errors:`, batchResult.errors);
            }

            // Update performance metrics
            const processingTime = Date.now() - startTime;
            this.updatePerformanceMetrics(followupEmails.length, processingTime);

            // Sort by priority and date with enhanced algorithm
            return this.sortFollowupEmails(followupEmails);
        } catch (error) {
            this.trackAnalyticsEvent('error_occurred', { error: (error as Error).message });
            console.error('Error analyzing emails:', error);
            throw new Error(`Failed to analyze emails: ${(error as Error).message}`);
        }
    }

    // New method to cancel email analysis if needed
    public cancelAnalysis(): void {
        this.batchProcessor.cancelProcessing();
        this.trackAnalyticsEvent('batch_processed', { action: 'cancelled' });
    }

    // Enhanced analytics to include comprehensive cache and batch processing metrics
    public getAnalytics(): any {
        const activeTasks = this.batchProcessor.getActiveProcessingTasks();
        const cacheStats = this.cacheService.getStats();
        
        return {
            metrics: this.performanceMetrics,
            recentEvents: this.analyticsEvents.slice(-100), // Last 100 events
            cacheStats: {
                hitRate: cacheStats.hitRate,
                totalEntries: cacheStats.totalEntries,
                memoryUsage: cacheStats.totalMemoryUsage,
                totalHits: cacheStats.totalHits,
                totalMisses: cacheStats.totalMisses,
                totalEvictions: cacheStats.totalEvictions,
                averageAccessCount: cacheStats.averageAccessCount
            },
            batchProcessing: {
                activeTasks: activeTasks.length,
                activeTaskIds: activeTasks
            }
        };
    }

    // New bulk operations functionality
    public async bulkSnoozeEmails(emailIds: string[], minutes: number): Promise<void> {
        const snoozeUntil = new Date();
        snoozeUntil.setMinutes(snoozeUntil.getMinutes() + minutes);
        
        emailIds.forEach(emailId => {
            this.snoozedEmails.set(emailId, snoozeUntil);
        });
        
        // Invalidate cache for affected emails
        this.invalidateEmailCaches(emailIds);
        
        this.trackAnalyticsEvent('email_analyzed', { 
            action: 'bulk_snooze', 
            count: emailIds.length, 
            minutes 
        });
    }

    public async bulkDismissEmails(emailIds: string[]): Promise<void> {
        emailIds.forEach(emailId => {
            this.dismissedEmails.add(emailId);
        });
        
        // Invalidate cache for affected emails
        this.invalidateEmailCaches(emailIds);
        
        this.trackAnalyticsEvent('email_analyzed', { 
            action: 'bulk_dismiss', 
            count: emailIds.length 
        });
    }

    // Enhanced email analysis with sentiment and context
    public async analyzeEmailSentiment(emailBody: string): Promise<'positive' | 'neutral' | 'negative' | 'urgent'> {
        // Check cache first
        const cacheKey = this.generateCacheKey('sentiment', emailBody);
        const cachedSentiment = this.cacheService.get<'positive' | 'neutral' | 'negative' | 'urgent'>(cacheKey);
        
        if (cachedSentiment) {
            this.trackAnalyticsEvent('cache_hit', { type: 'sentiment' });
            return cachedSentiment;
        }

        this.trackAnalyticsEvent('cache_miss', { type: 'sentiment' });

        try {
            const response = await this.llmService!.analyzeSentiment(emailBody);
            const sentiment = response.sentiment || this.analyzeSentimentBasic(emailBody);
            
            // Cache the result
            this.cacheService.set(cacheKey, sentiment, 6 * 60 * 60 * 1000); // 6 hours for sentiment
            
            return sentiment;
        } catch (error) {
            console.warn('Failed to get LLM sentiment analysis:', error);
            return this.analyzeSentimentBasic(emailBody);
        }
    }

    // Enhanced cache management methods
    public clearCache(pattern?: string): number {
        if (pattern) {
            const regex = new RegExp(pattern);
            return this.cacheService.bulkInvalidate(regex);
        } else {
            const stats = this.cacheService.getStats();
            this.cacheService.clear();
            this.trackAnalyticsEvent('cache_hit', { action: 'cache_cleared', entriesCleared: stats.totalEntries });
            return stats.totalEntries;
        }
    }

    public getCacheStats() {
        return this.cacheService.getStats();
    }

    public getCacheMemoryPressure(): number {
        const cacheService = this.cacheService as CacheService;
        return cacheService.getMemoryPressure ? cacheService.getMemoryPressure() : 0;
    }

    // Private helper methods with enhanced caching
    private async getSentEmailsWithCaching(emailCount: number, cutoffDate: Date): Promise<any[]> {
        const cacheKey = this.generateCacheKey('sent_emails', { emailCount, cutoffDate: cutoffDate.toISOString() });
        
        // Check cache first
        const cached = this.cacheService.get<any[]>(cacheKey);
        if (cached) {
            this.trackAnalyticsEvent('cache_hit', { type: 'sent_emails' });
            return cached;
        }

        this.trackAnalyticsEvent('cache_miss', { type: 'sent_emails' });
        
        const emails = await this.withRetry(
            () => this.getSentEmailsRest(emailCount, cutoffDate),
            this.DEFAULT_RETRY_OPTIONS
        );
        
        // Cache with 15 minute TTL for sent emails
        this.cacheService.set(cacheKey, emails, 15 * 60 * 1000);
        return emails;
    }

    private async processConversationWithCaching(
        conversationId: string, 
        _conversationEmails: any[],
        currentUserEmail: string, 
        selectedAccounts: string[]
    ): Promise<FollowupEmail | null> {
        // Check analysis cache first
        const cacheKey = this.generateCacheKey('analysis', { conversationId, selectedAccounts });
        const cached = this.cacheService.get<FollowupEmail | null>(cacheKey);
        if (cached !== null) {
            this.trackAnalyticsEvent('cache_hit', { type: 'analysis' });
            return cached;
        }

        this.trackAnalyticsEvent('cache_miss', { type: 'analysis' });

        // Get the full thread for this conversation with caching
        const threadMessages = await this.getConversationThreadCached(conversationId);
        
        // Check if the last email in thread was sent by current user
        const lastMessage = this.getLastMessageInThread(threadMessages);
        if (!lastMessage || !lastMessage.isFromCurrentUser) {
            // Cache null result to avoid reprocessing
            this.cacheService.set(cacheKey, null, 5 * 60 * 1000); // 5 minutes for null results
            return null;
        }

        // Check if there's been a response after the last sent message
        const hasResponse = this.checkForResponseInThread(threadMessages, lastMessage.sentDate);
        if (hasResponse) {
            this.cacheService.set(cacheKey, null, 5 * 60 * 1000);
            return null;
        }

        // Filter by selected accounts if specified
        if (selectedAccounts.length > 0 && !selectedAccounts.includes(lastMessage.from)) {
            this.cacheService.set(cacheKey, null, 5 * 60 * 1000);
            return null;
        }

        // Check if email is snoozed or dismissed
        if (this.isEmailSnoozed(lastMessage.id) || this.isEmailDismissed(lastMessage.id)) {
            this.cacheService.set(cacheKey, null, 5 * 60 * 1000);
            return null;
        }

        const followupEmail = await this.createFollowupEmailEnhanced(lastMessage, threadMessages, currentUserEmail);
        
        // Cache the analysis result
        this.cacheService.set(cacheKey, followupEmail);
        
        return followupEmail;
    }

    private async getConversationThreadCached(conversationId: string): Promise<ThreadMessage[]> {
        const cacheKey = this.generateCacheKey('thread', conversationId);
        
        const cached = this.cacheService.get<ThreadMessage[]>(cacheKey);
        if (cached) {
            this.trackAnalyticsEvent('cache_hit', { type: 'thread' });
            return cached;
        }

        this.trackAnalyticsEvent('cache_miss', { type: 'thread' });
        
        const thread = await this.withRetry(
            () => this.getConversationThread(conversationId),
            this.DEFAULT_RETRY_OPTIONS
        );
        
        // Cache thread data for 20 minutes
        this.cacheService.set(cacheKey, thread, 20 * 60 * 1000);
        return thread;
    }

    private generateCacheKey(prefix: string, data: any): string {
        try {
            // Create deterministic hash for consistent caching
            const content = typeof data === 'string' ? data : JSON.stringify(data);
            const hash = require('crypto').createHash('md5').update(content).digest('hex');
            return `email:${prefix}:${hash}`;
        } catch (error) {
            // Fallback to timestamp-based key
            return `email:${prefix}:${Date.now()}:${Math.random().toString(36).substr(2, 9)}`;
        }
    }

    private invalidateEmailCaches(emailIds: string[]): void {
        // Invalidate related cache entries when emails are modified
        emailIds.forEach(emailId => {
            const patterns = [
                new RegExp(`email:analysis:.*${emailId}`),
                new RegExp(`email:thread:.*${emailId}`),
                new RegExp(`email:sentiment:.*${emailId}`)
            ];
            
            patterns.forEach(pattern => {
                this.cacheService.bulkInvalidate(pattern);
            });
        });
    }

    // Snooze and dismiss functionality
    public snoozeEmail(emailId: string, minutes: number): void {
        const snoozeUntil = new Date();
        snoozeUntil.setMinutes(snoozeUntil.getMinutes() + minutes);
        this.snoozedEmails.set(emailId, snoozeUntil);
    }

    public snoozeEmailUntil(emailId: string, until: Date): void {
        this.snoozedEmails.set(emailId, until);
    }

    public unsnoozeEmail(emailId: string): void {
        this.snoozedEmails.delete(emailId);
    }

    public dismissEmail(emailId: string): void {
        this.dismissedEmails.add(emailId);
    }

    private isEmailSnoozed(emailId: string): boolean {
        const snoozeUntil = this.snoozedEmails.get(emailId);
        if (!snoozeUntil) return false;
        
        if (new Date() >= snoozeUntil) {
            this.snoozedEmails.delete(emailId);
            return false;
        }
        
        return true;
    }

    private isEmailDismissed(emailId: string): boolean {
        return this.dismissedEmails.has(emailId);
    }

    // Email processing helper methods
    private groupEmailsByConversation(emails: any[]): Map<string, any[]> {
        const groups = new Map<string, any[]>();
        
        emails.forEach(email => {
            const conversationId = email.conversationId || email.id;
            if (!groups.has(conversationId)) {
                groups.set(conversationId, []);
            }
            groups.get(conversationId)!.push(email);
        });
        
        return groups;
    }

    private async getSentEmailsRest(emailCount: number, cutoffDate: Date): Promise<any[]> {
        return new Promise((resolve, reject) => {
            Office.context.mailbox.makeEwsRequestAsync(
                this.buildGetSentEmailsRequest(emailCount, cutoffDate),
                (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        try {
                            const response = result.value;
                            const emails = this.parseSentEmailsResponse(response);
                            resolve(emails);
                        } catch (error) {
                            reject(error);
                        }
                    } else {
                        reject(new Error(result.error?.message || 'Failed to get sent emails'));
                    }
                }
            );
        });
    }

    private buildGetSentEmailsRequest(emailCount: number, cutoffDate: Date): string {
        const cutoffDateISO = cutoffDate.toISOString();
        return `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" 
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" 
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013" />
  </soap:Header>
  <soap:Body>
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject" />
          <t:FieldURI FieldURI="item:DateTimeSent" />
          <t:FieldURI FieldURI="message:ToRecipients" />
          <t:FieldURI FieldURI="message:From" />
          <t:FieldURI FieldURI="item:Body" />
          <t:FieldURI FieldURI="conversation:ConversationId" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:IndexedPageItemView MaxEntriesReturned="${emailCount}" Offset="0" BasePoint="Beginning" />
      <m:Restriction>
        <t:And>
          <t:IsGreaterThan>
            <t:FieldURI FieldURI="item:DateTimeSent" />
            <t:FieldURIOrConstant>
              <t:Constant Value="${cutoffDateISO}" />
            </t:FieldURIOrConstant>
          </t:IsGreaterThan>
        </t:And>
      </m:Restriction>
      <m:ParentFolderIds>
        <t:DistinguishedFolderId Id="sentitems" />
      </m:ParentFolderIds>
    </m:FindItem>
  </soap:Body>
</soap:Envelope>`;
    }

    private parseSentEmailsResponse(xmlResponse: string): any[] {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlResponse, 'text/xml');
        const messages = xmlDoc.getElementsByTagName('t:Message');
        
        const emails: any[] = [];
        for (let i = 0; i < messages.length; i++) {
            const message = messages[i];
            const email = {
                id: this.getElementText(message, 't:ItemId', 'Id'),
                subject: this.getElementText(message, 't:Subject'),
                dateTimeSent: this.getElementText(message, 't:DateTimeSent'),
                conversationId: this.getElementText(message, 't:ConversationId', 'Id'),
                body: { content: this.getElementText(message, 't:Body') },
                from: { emailAddress: { address: this.getElementText(message, 't:From t:Mailbox t:EmailAddress') } },
                toRecipients: this.parseRecipients(message, 't:ToRecipients')
            };
            emails.push(email);
        }
        
        return emails;
    }

    private getElementText(parent: Element, selector: string, attribute?: string): string {
        const element = parent.querySelector(selector);
        if (!element) return '';
        
        if (attribute) {
            return element.getAttribute(attribute) || '';
        }
        
        return element.textContent || '';
    }

    private parseRecipients(message: Element, selector: string): any[] {
        const recipientsElement = message.querySelector(selector);
        if (!recipientsElement) return [];
        
        const mailboxes = recipientsElement.querySelectorAll('t:Mailbox');
        const recipients: any[] = [];
        
        for (let i = 0; i < mailboxes.length; i++) {
            const mailbox = mailboxes[i];
            const emailAddress = this.getElementText(mailbox, 't:EmailAddress');
            if (emailAddress) {
                recipients.push({ emailAddress: { address: emailAddress } });
            }
        }
        
        return recipients;
    }

    private async getConversationThread(conversationId: string): Promise<ThreadMessage[]> {
        return new Promise((resolve, reject) => {
            Office.context.mailbox.makeEwsRequestAsync(
                this.buildGetConversationRequest(conversationId),
                (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        try {
                            const response = result.value;
                            const threadMessages = this.parseConversationResponse(response);
                            resolve(threadMessages);
                        } catch (error) {
                            reject(error);
                        }
                    } else {
                        reject(new Error(result.error?.message || 'Failed to get conversation thread'));
                    }
                }
            );
        });
    }

    private buildGetConversationRequest(conversationId: string): string {
        return `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" 
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" 
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013" />
  </soap:Header>
  <soap:Body>
    <m:GetConversationItems>
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject" />
          <t:FieldURI FieldURI="item:DateTimeSent" />
          <t:FieldURI FieldURI="message:ToRecipients" />
          <t:FieldURI FieldURI="message:From" />
          <t:FieldURI FieldURI="item:Body" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:ConversationsRequestType>
        <t:ConversationId Id="${conversationId}" />
      </m:ConversationsRequestType>
    </m:GetConversationItems>
  </soap:Body>
</soap:Envelope>`;
    }

    private parseConversationResponse(xmlResponse: string): ThreadMessage[] {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlResponse, 'text/xml');
        const messages = xmlDoc.getElementsByTagName('t:Message');
        
        const currentUserEmail = Office.context.mailbox.userProfile.emailAddress;
        const threadMessages: ThreadMessage[] = [];
        
        for (let i = 0; i < messages.length; i++) {
            const message = messages[i];
            const fromEmail = this.getElementText(message, 't:From t:Mailbox t:EmailAddress');
            
            const threadMessage: ThreadMessage = {
                id: this.getElementText(message, 't:ItemId', 'Id'),
                subject: this.getElementText(message, 't:Subject'),
                from: fromEmail,
                to: this.parseRecipients(message, 't:ToRecipients').map(r => r.emailAddress.address),
                sentDate: new Date(this.getElementText(message, 't:DateTimeSent')),
                body: this.getElementText(message, 't:Body'),
                isFromCurrentUser: fromEmail === currentUserEmail
            };
            
            threadMessages.push(threadMessage);
        }
        
        return threadMessages.sort((a, b) => a.sentDate.getTime() - b.sentDate.getTime());
    }

    private getLastMessageInThread(threadMessages: ThreadMessage[]): ThreadMessage | null {
        if (threadMessages.length === 0) return null;
        
        return threadMessages.reduce((latest, current) => 
            current.sentDate > latest.sentDate ? current : latest
        );
    }

    private checkForResponseInThread(threadMessages: ThreadMessage[], lastSentDate: Date): boolean {
        return threadMessages.some(message => 
            !message.isFromCurrentUser && message.sentDate > lastSentDate
        );
    }

    private generateSummary(emailBody: string, subject: string): string {
        if (!emailBody || emailBody.trim().length === 0) {
            return `Email about: ${subject}`;
        }

        // Remove HTML tags
        const cleanBody = emailBody.replace(/<[^>]*>/g, '');
        
        // Truncate if too long
        if (cleanBody.length <= this.SUMMARY_MAX_LENGTH) {
            return cleanBody;
        }
        
        return cleanBody.substring(0, this.SUMMARY_MAX_LENGTH - 3) + '...';
    }

    // Enhanced sentiment analysis
    private analyzeSentimentBasic(emailBody: string): 'positive' | 'neutral' | 'negative' | 'urgent' {
        const urgentKeywords = ['urgent', 'asap', 'immediately', 'critical', 'emergency'];
        const negativeKeywords = ['problem', 'issue', 'error', 'failed', 'wrong'];
        const positiveKeywords = ['thanks', 'great', 'excellent', 'perfect', 'appreciate'];

        const lowerBody = emailBody.toLowerCase();
        
        if (urgentKeywords.some(keyword => lowerBody.includes(keyword))) {
            return 'urgent';
        }
        if (negativeKeywords.some(keyword => lowerBody.includes(keyword))) {
            return 'negative';
        }
        if (positiveKeywords.some(keyword => lowerBody.includes(keyword))) {
            return 'positive';
        }
        
        return 'neutral';
    }

    private async createFollowupEmailEnhanced(
        lastMessage: ThreadMessage, 
        threadMessages: ThreadMessage[], 
        currentUserEmail: string
    ): Promise<FollowupEmail> {
        const daysSinceSent = Math.floor((Date.now() - lastMessage.sentDate.getTime()) / (1000 * 60 * 60 * 24));
        
        let llmSummary: string | undefined;
        let llmSuggestion: string | undefined;
        let sentiment: 'positive' | 'neutral' | 'negative' | 'urgent' = 'neutral';

        // Get LLM analysis if available
        if (this.llmService) {
            try {
                const [llmResponse, sentimentResult] = await Promise.all([
                    this.llmService.analyzeThread({
                        emails: threadMessages,
                        context: `User email: ${currentUserEmail}`
                    }),
                    this.analyzeEmailSentiment(lastMessage.body)
                ]);
                
                llmSummary = llmResponse; // analyzeThread returns a string, not an object
                llmSuggestion = await this.llmService.generateFollowupSuggestions(lastMessage.body).then(suggestions => suggestions[0]);
                sentiment = sentimentResult;
            } catch (error) {
                console.warn('Failed to get LLM analysis:', error);
                sentiment = this.analyzeSentimentBasic(lastMessage.body);
            }
        } else {
            sentiment = this.analyzeSentimentBasic(lastMessage.body);
        }

        const priority = this.calculateEnhancedPriority(daysSinceSent, threadMessages.length > 1, sentiment);

        return {
            id: lastMessage.id,
            subject: lastMessage.subject,
            recipients: lastMessage.to,
            sentDate: lastMessage.sentDate,
            body: lastMessage.body,
            summary: llmSummary || this.generateSummary(lastMessage.body, lastMessage.subject),
            priority,
            daysWithoutResponse: daysSinceSent,
            conversationId: lastMessage.id,
            hasAttachments: false,
            accountEmail: lastMessage.from,
            threadMessages,
            isSnoozed: false,
            isDismissed: false,
            llmSuggestion,
            llmSummary,
            sentiment // New field
        };
    }

    private calculateEnhancedPriority(
        daysSinceSent: number, 
        hasThread: boolean, 
        sentiment: 'positive' | 'neutral' | 'negative' | 'urgent'
    ): 'high' | 'medium' | 'low' {
        let priority: 'high' | 'medium' | 'low' = 'low';
        
        // Base priority on days
        if (daysSinceSent >= 7) {
            priority = 'high';
        } else if (daysSinceSent >= 3) {
            priority = 'medium';
        }
        
        // Boost priority based on sentiment
        if (sentiment === 'urgent') {
            priority = 'high';
        } else if (sentiment === 'negative' && priority === 'low') {
            priority = 'medium';
        }
        
        // Boost priority if it's part of a longer thread
        if (hasThread && priority === 'low') {
            priority = 'medium';
        } else if (hasThread && priority === 'medium') {
            priority = 'high';
        }
        
        return priority;
    }

    private sortFollowupEmails(followupEmails: FollowupEmail[]): FollowupEmail[] {
        return followupEmails.sort((a, b) => {
            const priorityOrder = { high: 3, medium: 2, low: 1 };
            const priorityDiff = priorityOrder[b.priority] - priorityOrder[a.priority];
            
            if (priorityDiff !== 0) {
                return priorityDiff;
            }

            // Secondary sort by sentiment (urgent first)
            const sentimentOrder: { [key: string]: number } = { urgent: 4, negative: 3, neutral: 2, positive: 1 };
            const sentimentA = (a as any).sentiment as string || 'neutral';
            const sentimentB = (b as any).sentiment as string || 'neutral';
            const sentimentDiff = (sentimentOrder[sentimentB] || 2) - (sentimentOrder[sentimentA] || 2);
            
            if (sentimentDiff !== 0) {
                return sentimentDiff;
            }
            
            // Tertiary sort by date
            return b.sentDate.getTime() - a.sentDate.getTime();
        });
    }

    private async withRetry<T>(
        operation: () => Promise<T>,
        options: RetryOptions = this.DEFAULT_RETRY_OPTIONS
    ): Promise<T> {
        let lastError: Error | null = null;
        
        for (let attempt = 0; attempt <= options.maxRetries; attempt++) {
            try {
                return await operation();
            } catch (error) {
                lastError = error as Error;
                
                if (attempt === options.maxRetries) {
                    break;
                }
                
                const delay = Math.min(
                    options.baseDelay * Math.pow(2, attempt),
                    options.maxDelay
                );
                
                this.trackAnalyticsEvent('retry_attempted', { 
                    attempt, 
                    delay, 
                    error: lastError.message 
                });
                
                await new Promise(resolve => setTimeout(resolve, delay));
            }
        }
        
        throw lastError || new Error('Operation failed after all retries');
    }

    private trackAnalyticsEvent(type: AnalyticsEvent['type'], details: any): void {
        this.analyticsEvents.push({
            type,
            timestamp: Date.now(),
            details
        });
        
        // Keep only last 1000 events
        if (this.analyticsEvents.length > 1000) {
            this.analyticsEvents = this.analyticsEvents.slice(-1000);
        }
    }

    private updatePerformanceMetrics(analyzedCount: number, processingTime: number): void {
        this.performanceMetrics.totalAnalyzed += analyzedCount;
        const cacheStats = this.cacheService.getStats();
        this.performanceMetrics.cacheHits = cacheStats.totalHits;
        this.performanceMetrics.cacheMisses = cacheStats.totalMisses;
        this.performanceMetrics.errors = this.analyticsEvents.filter(e => e.type === 'error_occurred').length;
        
        // Calculate rolling average processing time
        const totalTime = this.performanceMetrics.averageProcessingTime * (this.performanceMetrics.totalAnalyzed - analyzedCount);
        this.performanceMetrics.averageProcessingTime = (totalTime + processingTime) / this.performanceMetrics.totalAnalyzed;
    }
}