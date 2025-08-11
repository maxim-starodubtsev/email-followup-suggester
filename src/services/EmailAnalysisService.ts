import { FollowupEmail, ThreadMessage } from '../models/FollowupEmail';
import { LlmService } from './LlmService';
import { BatchProcessor, BatchProcessingOptions, BatchResult } from './BatchProcessor';
import { CacheService, ICacheService } from './CacheService';
import { XmlParsingService, ParsedEmail } from './XmlParsingService';
import { Configuration } from '../models/Configuration';

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
    // Window for cross-conversation dedupe when ConversationId fragments but represents the same human thread
    private readonly CROSS_CONV_DEDUPE_WINDOW_MS = 3 * 24 * 60 * 60 * 1000; // 3 days
    private readonly DEFAULT_RETRY_OPTIONS: RetryOptions = {
        maxRetries: 3,
        baseDelay: 1000,
        maxDelay: 10000
    };

    private llmService?: LlmService;
    private snoozedEmails: Map<string, Date> = new Map();
    private dismissedEmails: Set<string> = new Set();
    private configuration?: Configuration;
    
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
    private xmlParsingService: XmlParsingService;
    // Cache of recent emails for artificial thread building in fallback mode
    private recentEmailsContext: ParsedEmail[] = [];

    constructor(cacheService?: ICacheService) {
        this.batchProcessor = new BatchProcessor();
        this.xmlParsingService = new XmlParsingService();
        
        // Initialize enhanced caching system
        this.cacheService = cacheService || new CacheService({
            defaultTtl: 30 * 60 * 1000, // 30 minutes for email analysis
            maxMemoryUsage: 25 * 1024 * 1024, // 25MB for email cache
            maxEntries: 5000,
            evictionPolicy: 'lru',
            enableContentHashing: true,
            enableStatistics: true
        });

        // Keep references to optional fallback methods so TS doesn't flag them as unused when wired only in tests
        if (false) {
            // eslint-disable-next-line @typescript-eslint/no-floating-promises
            this.getConversationThreadFromConversationIdCached('noop');
        }
    }

    public setLlmService(llmService: LlmService): void {
        this.llmService = llmService;
    }

    public setConfiguration(configuration: Configuration): void {
        this.configuration = configuration;
    }

    public async analyzeEmails(emailCount: number, daysBack: number, selectedAccounts: string[]): Promise<FollowupEmail[]> {
        console.log(`[DEBUG] Starting analyzeEmails - Count: ${emailCount}, Days: ${daysBack}, Accounts: ${JSON.stringify(selectedAccounts)}`);
        
        const startTime = Date.now();
        this.trackAnalyticsEvent('batch_processed', { emailCount, daysBack, selectedAccounts });

        try {
            const currentUserEmail = Office.context.mailbox.userProfile.emailAddress;
            console.log(`[DEBUG] Current user email: ${currentUserEmail}`);
            
            const cutoffDate = new Date();
            cutoffDate.setDate(cutoffDate.getDate() - daysBack);
            console.log(`[DEBUG] Cutoff date: ${cutoffDate.toISOString()}`);

            // Get recent emails across multiple folders (not just Sent) with caching and retry
            // This implements requirement to analyze emails across all folders
            const recentEmails = await this.getRecentEmailsWithCaching(emailCount, cutoffDate);
            console.log(`[DEBUG] Retrieved ${recentEmails.length} recent emails across folders`);
            // Persist recent emails for use in artificial thread reconstruction (single-email fallback)
            try { this.recentEmailsContext = recentEmails as unknown as ParsedEmail[]; } catch { this.recentEmailsContext = []; }

            // Group emails by conversation ID
            const conversationGroups = this.groupEmailsByConversation(recentEmails);
            const conversationIds = Array.from(conversationGroups.keys());
            console.log(`[DEBUG] Grouped into ${conversationIds.length} conversation groups`);
            
            // Log conversation details
            conversationGroups.forEach((emails, conversationId) => {
                console.log(`[DEBUG] Conversation ${conversationId}: ${emails.length} emails`);
                emails.forEach((email, index) => {
                    console.log(`[DEBUG]   Email ${index + 1}: "${email.subject}" sent ${email.dateTimeSent}`);
                });
            });

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
            let followupEmails: FollowupEmail[] = batchResult.results
                .filter((result): result is FollowupEmail => result !== null);

            // Dedupe by conversation to avoid multiple entries from same thread
            followupEmails = this.dedupeFollowupEmails(followupEmails);

            console.log(`[DEBUG] Final result: ${followupEmails.length} emails need followup out of ${recentEmails.length} retrieved emails`);
            
            if (followupEmails.length === 0) {
                console.log(`[DEBUG] No followup emails found. This could mean:`);
                console.log(`[DEBUG] - All emails have been replied to`);
                console.log(`[DEBUG] - You weren't the last person to send in the conversations`);
                console.log(`[DEBUG] - All emails are outside the time window`);
                console.log(`[DEBUG] - Account filtering excluded all emails`);
            }

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

        // NEW: Retrieve recent emails from multiple folders (sent + inbox) to evaluate threads globally
    private async getRecentEmailsWithCaching(emailCount: number, cutoffDate: Date): Promise<any[]> {
                const cacheKey = this.generateCacheKey('recent_emails', { emailCount, cutoffDate: cutoffDate.toISOString() });
                const cached = this.cacheService.get<any[]>(cacheKey);
                if (cached) {
                        this.trackAnalyticsEvent('cache_hit', { type: 'recent_emails' });
                        return cached;
                }
                this.trackAnalyticsEvent('cache_miss', { type: 'recent_emails' });

                // Fetch from both Sent Items and Inbox (could be extended further later)
                // Fetch full requested count from each folder, then merge & trim.
                // Rationale: splitting count 50/50 caused missed threads when items were moved between folders.
                const [sent, inbox] = await this.withRetry(
                    async () => Promise.all([
                        this.getSentEmailsRest(emailCount, cutoffDate),
                        this.getInboxEmailsRest(emailCount, cutoffDate)
                    ]),
                    this.DEFAULT_RETRY_OPTIONS
                );

            // Fallback: if inbox returns nothing, reuse legacy sent emails logic (ensures legacy method referenced)
            if (inbox.length === 0) {
                try {
                    const legacySent = await this.getSentEmailsWithCaching(emailCount, cutoffDate);
                    if (legacySent.length > sent.length) {
                        sent.push(...legacySent.filter(e => !sent.some(s => s.id === e.id)));
                    }
                } catch (_) {
                    // ignore fallback error
                }
            }

                // Merge and sort by DateTimeSent desc, then truncate to requested emailCount
                const merged = [...sent, ...inbox]
                        .sort((a, b) => new Date(b.dateTimeSent).getTime() - new Date(a.dateTimeSent).getTime())
                        .slice(0, emailCount);

                this.cacheService.set(cacheKey, merged, 15 * 60 * 1000);
                return merged;
        }

        private async getInboxEmailsRest(emailCount: number, cutoffDate: Date): Promise<any[]> {
                return new Promise((resolve, reject) => {
                        Office.context.mailbox.makeEwsRequestAsync(
                                this.buildGetInboxEmailsRequest(emailCount, cutoffDate),
                                (result) => {
                                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                                                try {
                                                        const response = result.value;
                                                        const emails = this.parseSentEmailsResponse(response); // same parsing logic applies
                                                        resolve(emails);
                                                } catch (error) {
                                                        reject(error);
                                                }
                                        } else {
                                                reject(new Error(result.error?.message || 'Failed to get inbox emails'));
                                        }
                                }
                        );
                });
        }

        private buildGetInboxEmailsRequest(emailCount: number, cutoffDate: Date): string {
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
                    <t:FieldURI FieldURI="message:CcRecipients" />
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
                <t:DistinguishedFolderId Id="inbox" />
            </m:ParentFolderIds>
        </m:FindItem>
    </soap:Body>
</soap:Envelope>`;
        }

    private async processConversationWithCaching(
        conversationId: string, 
        conversationEmails: any[],
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
        // Use the first email's item ID since conversationId might not be a valid conversation ID
        const firstEmail = conversationEmails[0];
        const emailItemId = firstEmail.id;
        
        console.log(`[DEBUG] Processing conversation ${conversationId} with email ID ${emailItemId}`);
        
    // Prefer using conversation APIs (GetConversationItems), else fall back to item-based retrieval seeded by the email
        let threadMessages: ThreadMessage[] = [];
    if (conversationId && this.isEwsAvailable()) {
            try {
                threadMessages = await this.getConversationItemsConversationCached(conversationId);
            } catch (e) {
        console.warn('[WARN] GetConversationItems retrieval failed, will fallback to item-based', e);
            }
        }
        if (threadMessages.length === 0) {
            threadMessages = await this.getConversationThreadCached(emailItemId);
        }
        // If we only have a single message (fallback path), try to build an artificial thread by scanning recent emails
        if (threadMessages.length <= 1) {
            const base = threadMessages[0];
            const artificial = this.buildArtificialThreadFromRecentEmails(base, currentUserEmail);
            if (artificial.length > 1) {
                console.log(`[DEBUG] 🧩 Artificial thread assembled with ${artificial.length} messages (fallback mode)`);
                threadMessages = artificial;
            }
        }
        console.log(`[DEBUG] Retrieved ${threadMessages.length} thread messages for conversation ${conversationId}`);
        
        // Log all messages in the thread for debugging
        threadMessages.forEach((msg, index) => {
            console.log(`[DEBUG] Thread message ${index + 1}: From "${msg.from}" (${msg.isFromCurrentUser ? 'CURRENT USER' : 'OTHER USER'}) at ${msg.sentDate.toISOString()}`);
            console.log(`[DEBUG]   Subject: "${msg.subject}"`);
        });
        
        // Check if the last email in thread was sent by current user
        const lastMessage = this.getLastMessageInThread(threadMessages);
        if (!lastMessage) {
            console.log(`[DEBUG] ❌ FILTERED: Conversation ${conversationId} - No messages found in thread`);
            this.cacheService.set(cacheKey, null, 5 * 60 * 1000);
            return null;
        }

        console.log(`[DEBUG] Last message in thread: From "${lastMessage.from}" (${lastMessage.isFromCurrentUser ? 'CURRENT USER' : 'OTHER USER'}) at ${lastMessage.sentDate.toISOString()}`);
        
        if (!lastMessage.isFromCurrentUser) {
            console.log(`[DEBUG] ❌ FILTERED: Conversation ${conversationId} - Last message NOT from current user`);
            console.log(`[DEBUG]   Last message from: "${lastMessage.from}"`);
            console.log(`[DEBUG]   Current user: "${currentUserEmail}"`);
            console.log(`[DEBUG]   This email thread will NOT be shown as it doesn't need followup`);
            this.cacheService.set(cacheKey, null, 5 * 60 * 1000);
            return null;
        }

        console.log(`[DEBUG] ✅ PASSED: Last message IS from current user - checking for responses`);

        // Check if there's been a response after the last sent message
        const hasResponse = this.checkForResponseInThread(threadMessages, lastMessage.sentDate);
        if (hasResponse) {
            console.log(`[DEBUG] ❌ FILTERED: Conversation ${conversationId} - Has response after last sent message`);
            console.log(`[DEBUG]   This email thread will NOT be shown as it has been responded to`);
            this.cacheService.set(cacheKey, null, 5 * 60 * 1000);
            return null;
        }

        console.log(`[DEBUG] ✅ PASSED: No responses after last sent message`);

        // Filter by selected accounts if specified
        if (selectedAccounts.length > 0 && !selectedAccounts.includes(lastMessage.from)) {
            console.log(`[DEBUG] ❌ FILTERED: Conversation ${conversationId} - Account filter (${lastMessage.from} not in selected accounts)`);
            this.cacheService.set(cacheKey, null, 5 * 60 * 1000);
            return null;
        }

        console.log(`[DEBUG] ✅ PASSED: Account filter check`);

        // Check if email is snoozed or dismissed (respect user preferences)
        const isSnoozed = this.isEmailSnoozed(lastMessage.id);
        const isDismissed = this.isEmailDismissed(lastMessage.id);
        
        if ((isSnoozed && !(this.configuration?.showSnoozedEmails)) || 
            (isDismissed && !(this.configuration?.showDismissedEmails))) {
            console.log(`[DEBUG] ❌ FILTERED: Conversation ${conversationId} - Email is snoozed or dismissed (user setting)`);
            this.cacheService.set(cacheKey, null, 5 * 60 * 1000);
            return null;
        }

        console.log(`[DEBUG] ✅ PASSED: Snooze/dismiss filter`);
        console.log(`[DEBUG] 🎯 CREATING FOLLOWUP EMAIL for conversation ${conversationId}`);

    const followupEmail = await this.createFollowupEmailEnhanced(conversationId, lastMessage, threadMessages, currentUserEmail);
        
        // Cache the analysis result
        this.cacheService.set(cacheKey, followupEmail);
        
        return followupEmail;
    }

    private async getConversationThreadCached(emailItemId: string): Promise<ThreadMessage[]> {
        const cacheKey = this.generateCacheKey('thread', emailItemId);
        
        const cached = this.cacheService.get<ThreadMessage[]>(cacheKey);
        if (cached) {
            this.trackAnalyticsEvent('cache_hit', { type: 'thread' });
            return cached;
        }

        this.trackAnalyticsEvent('cache_miss', { type: 'thread' });
        
        const thread = await this.withRetry(
            () => this.getConversationThread(emailItemId),
            this.DEFAULT_RETRY_OPTIONS
        );
        
        // Cache thread data for 20 minutes
        this.cacheService.set(cacheKey, thread, 20 * 60 * 1000);
        return thread;
    }

    // New: conversationId-based cached retrieval (skips GetItem + fragile parse).
    private async getConversationThreadFromConversationIdCached(conversationId: string): Promise<ThreadMessage[]> {
    if (!this.isEwsAvailable()) return [];
        const cacheKey = this.generateCacheKey('threadConv', conversationId);
        const cached = this.cacheService.get<ThreadMessage[]>(cacheKey);
        if (cached) {
            this.trackAnalyticsEvent('cache_hit', { type: 'threadConv' });
            return cached;
        }
        this.trackAnalyticsEvent('cache_miss', { type: 'threadConv' });
        const messages = await this.withRetry(
            () => this.searchConversationAcrossFolders(conversationId),
            this.DEFAULT_RETRY_OPTIONS
        );
        this.cacheService.set(cacheKey, messages, 20 * 60 * 1000);
        return messages;
    }

    // New: GetConversationItems-based retrieval + cache
    private async getConversationItemsConversationCached(conversationId: string): Promise<ThreadMessage[]> {
    if (!this.isEwsAvailable()) return [];
        const cacheKey = this.generateCacheKey('convItems', conversationId);
        const cached = this.cacheService.get<ThreadMessage[]>(cacheKey);
        if (cached) {
            this.trackAnalyticsEvent('cache_hit', { type: 'convItems' });
            return cached;
        }
        this.trackAnalyticsEvent('cache_miss', { type: 'convItems' });
        const messages = await this.withRetry(
            () => this.getConversationItems(conversationId),
            this.DEFAULT_RETRY_OPTIONS
        );
        this.cacheService.set(cacheKey, messages, 20 * 60 * 1000);
        return messages;
    }

    private buildGetConversationItemsRequest(conversationId: string, maxItems = 250): string {
        // Using shallow to respect folder boundaries but conversation API returns nodes.
        // Optionally FoldersToIgnore could list drafts/deleteditems if filtering earlier.
    return `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013" />
  </soap:Header>
  <soap:Body>
    <m:GetConversationItems ReturnSynchronizationCookie="false">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
    <t:BodyType>Text</t:BodyType>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject" />
          <t:FieldURI FieldURI="item:DateTimeSent" />
          <t:FieldURI FieldURI="item:DateTimeReceived" />
          <t:FieldURI FieldURI="message:ToRecipients" />
          <t:FieldURI FieldURI="message:From" />
          <t:FieldURI FieldURI="item:Body" />
          <t:FieldURI FieldURI="conversation:ConversationId" />
          <t:FieldURI FieldURI="item:HasAttachments" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:FoldersToIgnore>
        <t:DistinguishedFolderId Id="drafts" />
        <t:DistinguishedFolderId Id="deleteditems" />
      </m:FoldersToIgnore>
            <m:MaxItemsToReturn>${maxItems}</m:MaxItemsToReturn>
      <m:ConversationIds>
        <t:ConversationId Id="${conversationId}" />
      </m:ConversationIds>
    </m:GetConversationItems>
  </soap:Body>
</soap:Envelope>`;
    }

    private async getConversationItems(conversationId: string): Promise<ThreadMessage[]> {
        return new Promise((resolve, reject) => {
            Office.context.mailbox.makeEwsRequestAsync(
                this.buildGetConversationItemsRequest(conversationId),
                (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        try {
                            const messages = this.parseGetConversationItemsResponse(result.value);
                            resolve(messages);
                        } catch (e) {
                            reject(e);
                        }
                    } else {
                        reject(new Error(result.error?.message || 'GetConversationItems failed'));
                    }
                }
            );
        });
    }

    private parseGetConversationItemsResponse(xml: string): ThreadMessage[] {
        const currentUserEmail = Office.context.mailbox.userProfile.emailAddress.toLowerCase();
        const parser = new DOMParser();
        const doc = parser.parseFromString(xml, 'text/xml');
        const messages: ThreadMessage[] = [];

        try {
            // ConversationNodes can contain one or more Items of different classes (Message, Meeting*, etc.)
            const nodeLists = doc.getElementsByTagNameNS('*', 'ConversationNode');
            for (let i = 0; i < nodeLists.length; i++) {
                const node = nodeLists[i];

                // Prefer items under the Items container if present
                const itemsContainer = node.getElementsByTagNameNS('*', 'Items')[0] || node;
                // Collect candidate item elements across common classes
                const candidateTags = ['Message', 'MeetingMessage', 'MeetingRequest', 'MeetingResponse', 'MeetingCancellation'];
                let itemElements: Element[] = [];
                for (const tag of candidateTags) {
                    const found = Array.from(itemsContainer.getElementsByTagNameNS('*', tag));
                    itemElements = itemElements.concat(found);
                }

                // If nothing matched, fall back to scanning any child elements that look like items (have ItemId)
                if (itemElements.length === 0) {
                    const all = Array.from(itemsContainer.getElementsByTagName('*')) as Element[];
                    itemElements = all.filter(el => el.getElementsByTagNameNS('*', 'ItemId').length > 0);
                }

                for (const item of itemElements) {
                    const idEl = item.getElementsByTagNameNS('*', 'ItemId')[0];
                    if (!idEl) continue;
                    const subjEl = item.getElementsByTagNameNS('*', 'Subject')[0];
                    const sentEl = item.getElementsByTagNameNS('*', 'DateTimeSent')[0];
                    const recvEl = item.getElementsByTagNameNS('*', 'DateTimeReceived')[0];
                    const fromEmail = this.extractEmailAddress(item.getElementsByTagNameNS('*', 'From')[0]);
                    const toEmails = this.extractMultipleAddresses(item.getElementsByTagNameNS('*', 'ToRecipients')[0]);
                    const bodyEl = item.getElementsByTagNameNS('*', 'Body')[0];

                    const id = idEl.getAttribute('Id') || idEl.textContent || '';
                    const sentDate = sentEl ? new Date(sentEl.textContent || '') : new Date();
                    const receivedDate = recvEl ? new Date(recvEl.textContent || '') : undefined;
                    const from = (fromEmail || '').trim();
                    const subj = (subjEl?.textContent || '').trim();
                    const body = (bodyEl?.textContent || '').trim();
                    const isFromCurrentUser = from.toLowerCase() === currentUserEmail;

                    messages.push({
                        id,
                        subject: subj,
                        from,
                        to: toEmails,
                        sentDate,
                        receivedDate,
                        body,
                        isFromCurrentUser
                    });
                }
            }

            // Sort chronologically by receivedDate if available else sentDate
            messages.sort((a, b) => {
                const aTime = (a.receivedDate || a.sentDate).getTime();
                const bTime = (b.receivedDate || b.sentDate).getTime();
                return aTime - bTime;
            });
        } catch (err) {
            console.error('[ERROR] Failed to parse GetConversationItems response:', err);
        }
        return messages;
    }

    private extractEmailAddress(container?: Element | null): string | undefined {
        if (!container) return undefined;
        const addr = container.getElementsByTagNameNS('*', 'EmailAddress')[0];
        if (addr?.textContent) return addr.textContent.trim();
        const addressNode = container.getElementsByTagNameNS('*', 'Address')[0];
        return addressNode?.textContent?.trim();
    }

    private extractMultipleAddresses(container?: Element | null): string[] {
        if (!container) return [];
        const emails: string[] = [];
        const mailboxes = container.getElementsByTagNameNS('*', 'Mailbox');
        for (let i = 0; i < mailboxes.length; i++) {
            const e = this.extractEmailAddress(mailboxes[i]);
            if (e) emails.push(e);
        }
        return emails;
    }

    private isEwsAvailable(): boolean {
        try {
            return typeof (Office as any)?.context?.mailbox?.makeEwsRequestAsync === 'function';
        } catch {
            return false;
        }
    }

    private generateCacheKey(prefix: string, data: any): string {
        try {
            // Create deterministic hash for consistent caching using a simple hash function
            const content = typeof data === 'string' ? data : JSON.stringify(data);
            const hash = this.simpleHash(content);
            return `email:${prefix}:${hash}`;
        } catch (error) {
            // Fallback to timestamp-based key
            return `email:${prefix}:${Date.now()}:${Math.random().toString(36).substr(2, 9)}`;
        }
    }

    private simpleHash(str: string): string {
        let hash = 0;
        if (str.length === 0) return hash.toString();
        for (let i = 0; i < str.length; i++) {
            const char = str.charCodeAt(i);
            hash = ((hash << 5) - hash) + char;
            hash = hash & hash; // Convert to 32-bit integer
        }
        return Math.abs(hash).toString(36);
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
        const groups = new Map<string, any>();
        
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
          <t:FieldURI FieldURI="message:CcRecipients" />
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

    private parseSentEmailsResponse(xmlResponse: string): ParsedEmail[] {
        console.log(`[DEBUG] Parsing EWS FindItem response using XmlParsingService`);
        
        // Validate the XML response first
        const validation = this.xmlParsingService.validateEwsResponse(xmlResponse);
        if (!validation.isValid) {
            console.error(`[ERROR] Invalid EWS response: ${validation.error}`);
            return [];
        }
        
        // Use the new parsing service
        const emails = this.xmlParsingService.parseFindItemResponse(xmlResponse);
        console.log(`[DEBUG] XmlParsingService parsed ${emails.length} emails`);
        
        // Log details for first few emails
        emails.slice(0, 3).forEach((email, index) => {
            console.log(`[DEBUG] Email ${index + 1}: Subject="${email.subject}", Date=${email.dateTimeSent}, From=${email.from.emailAddress.address}`);
        });
        
        return emails;
    }

    private async getConversationThread(emailItemId: string): Promise<ThreadMessage[]> {
        console.log(`[DEBUG] 🧵 THREAD RETRIEVAL: Getting conversation thread for email ID: ${emailItemId}`);
        
        return new Promise((resolve, reject) => {
            // First get the conversation ID from the email
            Office.context.mailbox.makeEwsRequestAsync(
                this.buildGetConversationIdRequest(emailItemId),
                async (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        try {
                            const conversationId = this.parseConversationIdResponse(result.value);
                            if (!conversationId) {
                                console.log(`[DEBUG] 📧 SINGLE EMAIL: No conversation ID found, treating as single email thread`);
                                const singleMessage = this.parseConversationResponse(result.value);
                                console.log(`[DEBUG] 📧 SINGLE EMAIL RESULT: ${singleMessage.length} message(s) parsed`);
                                resolve(singleMessage);
                                return;
                            }
                            console.log(`[DEBUG] 🔗 CONVERSATION FOUND: ID = ${conversationId}, fetching via GetConversationItems`);
                            try {
                                const convMessages = await this.getConversationItems(conversationId);
                                if (convMessages.length > 0) {
                                    console.log(`[DEBUG] 🎯 THREAD COMPLETE: Retrieved ${convMessages.length} messages via GetConversationItems`);
                                    // Log the complete thread structure
                                    console.log(`[DEBUG] 📋 COMPLETE THREAD STRUCTURE:`);
                                    convMessages.forEach((msg, index) => {
                                        console.log(`[DEBUG]   ${index + 1}. ${msg.isFromCurrentUser ? 'CURRENT USER' : 'OTHER USER'} - ${msg.sentDate.toISOString()} - "${msg.subject}"`);
                                    });
                                    resolve(convMessages);
                                } else {
                                    console.log(`[DEBUG] ⚠️ FALLBACK: GetConversationItems returned no messages, falling back to single message`);
                                    const singleMessage = this.parseConversationResponse(result.value);
                                    resolve(singleMessage);
                                }
                            } catch (convErr) {
                                console.error('[ERROR] 💥 THREAD FETCH FAILED: GetConversationItems after GetItem failed:', convErr);
                                const singleMessage = this.parseConversationResponse(result.value);
                                console.log(`[DEBUG] 🔄 FALLBACK COMPLETE: Using single message instead`);
                                resolve(singleMessage);
                            }
                        } catch (error) {
                            console.error('[ERROR] 💥 CONVERSATION PARSING FAILED: Failed to parse conversation ID response:', error);
                            reject(error);
                        }
                    } else {
                        console.error('[ERROR] 💥 EWS REQUEST FAILED: Failed to get conversation thread:', result.error?.message);
                        reject(new Error(result.error?.message || 'Failed to get conversation thread'));
                    }
                }
            );
        });
    }

    private parseConversationResponse(xmlResponse: string): ThreadMessage[] {
        console.log(`[DEBUG] Parsing EWS GetItem response to extract email details`);
        
        const currentUserEmail = Office.context.mailbox.userProfile.emailAddress.toLowerCase();
        
        // Validate the XML response first
        const validation = this.xmlParsingService.validateEwsResponse(xmlResponse);
        if (!validation.isValid) {
            console.error(`[ERROR] Invalid EWS response: ${validation.error}`);
            return [];
        }
        
        // Parse the GetItem response which contains a single email item
        try {
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(xmlResponse, 'text/xml');
            
            // Look for the message element in the GetItem response
            const messageElements = xmlDoc.getElementsByTagName('t:Message');
            if (messageElements.length === 0) {
                console.log(`[DEBUG] No message found in GetItem response`);
                return [];
            }
            
            const messageElement = messageElements[0];
            const threadMessage = this.parseMessageElement(messageElement, currentUserEmail);
            
            if (threadMessage) {
                console.log(`[DEBUG] Successfully parsed single email: "${threadMessage.subject}" from ${threadMessage.from} (${threadMessage.isFromCurrentUser ? 'current user' : 'other'})`);
                return [threadMessage];
            } else {
                console.log(`[DEBUG] Failed to parse message element`);
                return [];
            }
            
        } catch (error) {
            console.error(`[ERROR] Failed to parse GetItem response:`, error);
            return [];
        }
    }

    private getLastMessageInThread(threadMessages: ThreadMessage[]): ThreadMessage | null {
        if (threadMessages.length === 0) {
            console.log(`[DEBUG] ❌ EMPTY THREAD: No messages in thread`);
            return null;
        }
        // Determine most recent using receivedDate when present (fallback to sentDate)
        const sortedMessages = [...threadMessages].sort((a, b) => {
            const at = (a.receivedDate || a.sentDate).getTime();
            const bt = (b.receivedDate || b.sentDate).getTime();
            return bt - at;
        });
        const lastMessage = sortedMessages[0];
        
        console.log(`[DEBUG] 🏁 LAST MESSAGE IDENTIFIED:`);
        console.log(`[DEBUG]   From: "${lastMessage.from}" (${lastMessage.isFromCurrentUser ? 'CURRENT USER' : 'OTHER USER'})`);
        console.log(`[DEBUG]   Subject: "${lastMessage.subject}"`);
        console.log(`[DEBUG]   Date: ${(lastMessage.receivedDate || lastMessage.sentDate).toISOString()} (received vs sent basis)`);
        console.log(`[DEBUG]   Thread size: ${threadMessages.length} messages`);
        
        return lastMessage;
    }

    private checkForResponseInThread(threadMessages: ThreadMessage[], lastSentDate: Date): boolean {
        console.log(`[DEBUG] 🔍 THREAD ANALYSIS: Checking for responses after ${lastSentDate.toISOString()}`);
        console.log(`[DEBUG] Thread has ${threadMessages.length} messages total`);
        
        // Sort messages by date to ensure chronological order
        const sortedMessages = [...threadMessages].sort((a, b) => a.sentDate.getTime() - b.sentDate.getTime());
        
        console.log(`[DEBUG] Chronological message order:`);
        sortedMessages.forEach((message, index) => {
            const baseDate = (message.receivedDate || message.sentDate);
            const timeStatus = baseDate > lastSentDate ? 'AFTER' : 'BEFORE';
            const userStatus = message.isFromCurrentUser ? 'CURRENT USER' : 'OTHER USER';
            console.log(`[DEBUG]   ${index + 1}. ${timeStatus} last sent: ${userStatus} at ${baseDate.toISOString()}`);
        });
        
        // Find all messages after the last sent date that are not from current user
        const responsesAfterLastSent = sortedMessages.filter(message => {
            const isAfterLastSent = message.sentDate > lastSentDate;
            const isFromOther = !message.isFromCurrentUser;
            
            console.log(`[DEBUG] Evaluating message from ${message.isFromCurrentUser ? 'CURRENT USER' : 'OTHER USER'} at ${message.sentDate.toISOString()}: ` +
                       `after last sent = ${isAfterLastSent}, from other = ${isFromOther}`);
            
            return isAfterLastSent && isFromOther;
        });
        
        const hasResponse = responsesAfterLastSent.length > 0;
        console.log(`[DEBUG] 📧 RESPONSE ANALYSIS: Found ${responsesAfterLastSent.length} responses after last sent message`);
        
        if (hasResponse) {
            console.log(`[DEBUG] ✅ RESPONSES FOUND:`);
            responsesAfterLastSent.forEach((response, index) => {
                console.log(`[DEBUG]   Response ${index + 1}: from ${response.from} at ${response.sentDate.toISOString()}`);
                console.log(`[DEBUG]     Subject: "${response.subject}"`);
            });
            console.log(`[DEBUG] 🔒 This thread will be FILTERED OUT because it has responses`);
        } else {
            console.log(`[DEBUG] ❌ NO RESPONSES FOUND - This thread NEEDS FOLLOWUP`);
        }
        
        return hasResponse;
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
        conversationId: string,
        lastMessage: ThreadMessage, 
        threadMessages: ThreadMessage[], 
        currentUserEmail: string
    ): Promise<FollowupEmail> {
        const daysSinceSent = Math.floor((Date.now() - lastMessage.sentDate.getTime()) / (1000 * 60 * 60 * 24));
        
        let llmSummary: string | undefined;
        let llmSuggestion: string | undefined;
        let sentiment: 'positive' | 'neutral' | 'negative' | 'urgent' = 'neutral';

        // Check if AI features are disabled globally
        const aiDisabled = localStorage.getItem('aiDisabled') === 'true';

        // Get LLM analysis if available and not disabled
        if (this.llmService && !aiDisabled) {
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
            conversationId: conversationId,
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

    // Remove duplicate followup emails representing same conversation/thread. Keep newest sentDate.
    private dedupeFollowupEmails(followups: FollowupEmail[]): FollowupEmail[] {
        if (followups.length <= 1) return followups;

        // Sort newest first so we keep the most recent when deduping
        const sorted = [...followups].sort((a, b) => b.sentDate.getTime() - a.sentDate.getTime());

        const kept: FollowupEmail[] = [];
        const seenByConversation = new Set<string>();
        const humanThreadBuckets = new Map<string, number[]>(); // key -> kept sentDate times

        for (const f of sorted) {
            const convId = (f.conversationId || '').trim();
            const hasConvId = !!convId && convId !== f.id; // ignore bogus convIds equal to item id

            // If we've already kept an entry for this exact conversation, skip older ones
            if (hasConvId && seenByConversation.has(convId)) {
                continue;
            }

            // Build a normalized human-thread key: subject (without prefixes) + participants
            const normSubject = this.normalizeSubject(f.subject);
            const participants = Array.from(new Set([f.accountEmail, ...f.recipients].map(e => (e || '').toLowerCase().trim())))
                .filter(Boolean)
                .sort()
                .join(';');
            const humanKey = `${normSubject}|${participants}`;

            // If another kept entry exists with same humanKey within the dedupe window, skip this one
            const keptTimes = humanThreadBuckets.get(humanKey) || [];
            const withinWindow = keptTimes.some(t => Math.abs(t - f.sentDate.getTime()) <= this.CROSS_CONV_DEDUPE_WINDOW_MS);
            if (withinWindow) {
                // Duplicate across conversations for same human thread/timeframe
                continue;
            }

            kept.push(f);
            if (hasConvId) seenByConversation.add(convId);
            humanThreadBuckets.set(humanKey, [...keptTimes, f.sentDate.getTime()]);
        }

        // Preserve original sort intention (the caller sorts later by priority/date); return in the same newest-first order here
        return kept;
    }

    // Normalize subjects by removing common reply/forward prefixes and collapsing whitespace
    private normalizeSubject(subject: string): string {
        if (!subject) return '';
        let s = subject.trim();
        // Remove multiple stacked prefixes like Re:, Fwd:, FW:, SV:, VS:, Ответ:, etc.
        // Common international variants included conservatively
        const prefixRe = /^(re|fwd|fw|sv|vs|aw|答复|回复|rép|antwort|tr|rv|回复|ответ|отв|转发)\s*:\s*/i;
        // Loop to strip repeated prefixes
        while (prefixRe.test(s)) {
            s = s.replace(prefixRe, '');
            s = s.trim();
        }
        // Collapse internal whitespace and lowercase for stable matching
        s = s.replace(/\s+/g, ' ').toLowerCase();
        return s;
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

    // Helper methods for enhanced thread analysis
    private async searchConversationAcrossFolders(conversationId: string): Promise<ThreadMessage[]> {
        console.log(`[DEBUG] Searching for conversation ${conversationId} across multiple folders`);
        
        // Include root (deep traversal) to satisfy requirement: analyze across all folders & sub-folders
        // Order chosen to prioritise cheaper targeted folders before deep mailbox scan
        const folders = [
            'sentitems',      // Sent Items
            'inbox',          // Inbox
            'drafts',         // Drafts
            'deleteditems',   // Deleted Items
            'archive',        // Archive (if present; harmless if not)
            'msgfolderroot'   // Root (Deep traversal)
        ];
        
        const allMessages: ThreadMessage[] = [];
        
        for (const folder of folders) {
            try {
                console.log(`[DEBUG] Searching in folder: ${folder}`);
                const folderMessages = await this.searchConversationInFolder(conversationId, folder);
                allMessages.push(...folderMessages);
                console.log(`[DEBUG] Found ${folderMessages.length} messages in ${folder}`);
                
                // Debug: Log each message found in this folder
                folderMessages.forEach((msg, index) => {
                    console.log(`[DEBUG] Folder ${folder} message ${index + 1}: ${msg.id} - ${msg.sentDate.toISOString()} (${msg.sentDate.getTime()})`);
                });
            } catch (error) {
                console.warn(`[WARN] Failed to search in folder ${folder}:`, error);
                // Continue with other folders
            }
        }
        
        console.log(`[DEBUG] Total messages before deduplication: ${allMessages.length}`);
        allMessages.forEach((msg, index) => {
            console.log(`[DEBUG] Before dedup ${index + 1}: ${msg.id} - ${msg.sentDate.toISOString()} (${msg.sentDate.getTime()})`);
        });
        
        // Remove duplicates based on message ID
        const uniqueMessages = this.removeDuplicateMessages(allMessages);
        
        console.log(`[DEBUG] Total messages after deduplication: ${uniqueMessages.length}`);
        uniqueMessages.forEach((msg, index) => {
            console.log(`[DEBUG] After dedup ${index + 1}: ${msg.id} - ${msg.sentDate.toISOString()} (${msg.sentDate.getTime()})`);
        });
        
        // CRITICAL FIX: Ensure chronological sort (earliest first) with proper comparison
        console.log(`[DEBUG] Applying final chronological sort (earliest first)...`);
        const sortedMessages = [...uniqueMessages].sort((a, b) => {
            const timeA = a.sentDate.getTime();
            const timeB = b.sentDate.getTime();
            const diff = timeA - timeB; // Ascending order: earlier dates first
            console.log(`[DEBUG] Sort comparison: ${a.id} (${timeA}) vs ${b.id} (${timeB}) = ${diff} ${diff < 0 ? '(A first)' : diff > 0 ? '(B first)' : '(same)'}`);
            return diff;
        });
        
        console.log(`[DEBUG] Final sorted message order (chronological - earliest first):`);
        sortedMessages.forEach((msg, index) => {
            console.log(`[DEBUG]   ${index + 1}. ${msg.id} - ${msg.sentDate.toISOString()} (${msg.sentDate.getTime()}) - ${msg.isFromCurrentUser ? 'CURRENT USER' : 'OTHER USER'}`);
        });
        
        // Validate sort order - ensure each message is chronologically before the next
        for (let i = 1; i < sortedMessages.length; i++) {
            const prevTime = sortedMessages[i - 1].sentDate.getTime();
            const currTime = sortedMessages[i].sentDate.getTime();
            if (prevTime > currTime) {
                console.error(`[ERROR] ❌ Sort order violation detected!`);
                console.error(`[ERROR] Message ${i - 1} (${sortedMessages[i - 1].id}) time ${prevTime} > Message ${i} (${sortedMessages[i].id}) time ${currTime}`);
                throw new Error(`Sort order violation: Message ${i - 1} (${prevTime}) should not be after Message ${i} (${currTime})`);
            } else {
                console.log(`[DEBUG] ✅ Sort order correct: ${prevTime} <= ${currTime}`);
            }
        }
        
        console.log(`[DEBUG] ✅ Sort validation passed. Total unique messages found: ${sortedMessages.length}`);
        return sortedMessages;
    }

    private async searchConversationInFolder(conversationId: string, folderId: string): Promise<ThreadMessage[]> {
        return new Promise((resolve) => {
            Office.context.mailbox.makeEwsRequestAsync(
                this.buildSearchConversationRequest(
                    conversationId,
                    folderId,
                    folderId === 'msgfolderroot' ? 'Deep' : 'Shallow'
                ),
                (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        try {
                            const messages = this.parseSearchConversationResponse(result.value);
                            resolve(messages);
                        } catch (error) {
                            console.error(`[ERROR] Failed to parse search results for folder ${folderId}:`, error);
                            resolve([]); // Return empty array instead of rejecting
                        }
                    } else {
                        console.warn(`[WARN] Failed to search folder ${folderId}:`, result.error?.message);
                        resolve([]); // Return empty array instead of rejecting
                    }
                }
            );
        });
    }

        private buildSearchConversationRequest(conversationId: string, folderId: string, traversal: 'Shallow' | 'Deep' = 'Shallow'): string {
                return `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
                             xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" 
                             xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" 
                             xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
    <soap:Header>
        <t:RequestServerVersion Version="Exchange2013" />
    </soap:Header>
    <soap:Body>
        <m:FindItem Traversal="${traversal}">
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
            <m:IndexedPageItemView MaxEntriesReturned="200" Offset="0" BasePoint="Beginning" />
            <m:Restriction>
                <t:IsEqualTo>
                    <t:FieldURI FieldURI="conversation:ConversationId" />
                    <t:FieldURIOrConstant>
                        <t:Constant Value="${conversationId}" />
                    </t:FieldURIOrConstant>
                </t:IsEqualTo>
            </m:Restriction>
            <m:ParentFolderIds>
                <t:DistinguishedFolderId Id="${folderId}" />
            </m:ParentFolderIds>
        </m:FindItem>
    </soap:Body>
</soap:Envelope>`;
    }

    private parseSearchConversationResponse(xmlResponse: string): ThreadMessage[] {
        const currentUserEmail = Office.context.mailbox.userProfile.emailAddress.toLowerCase();
        const threadMessages: ThreadMessage[] = [];
        
        try {
            // Validate the XML response first
            const validation = this.xmlParsingService.validateEwsResponse(xmlResponse);
            if (!validation.isValid) {
                console.error(`[ERROR] Invalid EWS search response: ${validation.error}`);
                return [];
            }

            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(xmlResponse, 'text/xml');
            
            // Look for all message elements in the search results
            const messageElements = xmlDoc.getElementsByTagName('t:Message');
            
            for (let i = 0; i < messageElements.length; i++) {
                const messageElement = messageElements[i];
                
                try {
                    const threadMessage = this.parseMessageElement(messageElement, currentUserEmail);
                    if (threadMessage) {
                        threadMessages.push(threadMessage);
                    }
                } catch (error) {
                    console.warn(`[WARN] Failed to parse message ${i} in search results:`, error);
                    // Continue processing other messages
                }
            }
            
            console.log(`[DEBUG] Parsed ${threadMessages.length} messages from search results`);
            return threadMessages;
            
        } catch (error) {
            console.error(`[ERROR] Failed to parse search conversation response:`, error);
            return [];
        }
    }

    private removeDuplicateMessages(messages: ThreadMessage[]): ThreadMessage[] {
        const seen = new Set<string>();
        const uniqueMessages: ThreadMessage[] = [];
        
        console.log(`[DEBUG] Starting deduplication with ${messages.length} messages`);
        
        // Process messages in order received - don't pre-sort here
        for (const message of messages) {
            // Create a unique key based on message ID, or fall back to content-based key
            const key = message.id || `${message.from}-${message.sentDate.getTime()}-${message.subject}`;
            
            if (!seen.has(key)) {
                seen.add(key);
                uniqueMessages.push(message);
                console.log(`[DEBUG] ✅ Keeping unique message: ${message.id} at ${message.sentDate.toISOString()} (${message.sentDate.getTime()})`);
            } else {
                console.log(`[DEBUG] ❌ Removing duplicate message: ${message.id} at ${message.sentDate.toISOString()} (${message.sentDate.getTime()})`);
            }
        }
        
        console.log(`[DEBUG] Deduplication complete: Removed ${messages.length - uniqueMessages.length} duplicate messages`);
        console.log(`[DEBUG] Unique messages (in order processed):`);
        uniqueMessages.forEach((msg, index) => {
            console.log(`[DEBUG]   ${index + 1}. ${msg.id} - ${msg.sentDate.toISOString()} (${msg.sentDate.getTime()})`);
        });
        
        return uniqueMessages;
    }

    private buildGetConversationIdRequest(emailItemId: string): string {
        return `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" 
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" 
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013" />
  </soap:Header>
  <soap:Body>
    <m:GetItem>
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="conversation:ConversationId" />
          <t:FieldURI FieldURI="item:Subject" />
          <t:FieldURI FieldURI="item:DateTimeSent" />
          <t:FieldURI FieldURI="message:ToRecipients" />
          <t:FieldURI FieldURI="message:From" />
          <t:FieldURI FieldURI="item:Body" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:ItemIds>
        <t:ItemId Id="${emailItemId}" />
      </m:ItemIds>
    </m:GetItem>
  </soap:Body>
</soap:Envelope>`;
    }

    private parseConversationIdResponse(xmlResponse: string): string | null {
        try {
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(xmlResponse, 'text/xml');

            // Strategy: namespace-agnostic search for element whose localName == ConversationId
            let conversationId: string | null = null;
            const allElements = xmlDoc.getElementsByTagName('*');
            for (let i = 0; i < allElements.length; i++) {
                const el = allElements[i];
                if (el.localName === 'ConversationId') {
                    conversationId = el.getAttribute('Id');
                    if (conversationId) break;
                }
            }
            if (conversationId) {
                console.log(`[DEBUG] Found conversation ID: ${conversationId}`);
                return conversationId;
            }
            console.log('[DEBUG] No conversation ID found in response (namespace-agnostic search)');
            return null;
        } catch (error) {
            console.error(`[ERROR] Failed to parse conversation ID response:`, error);
            return null;
        }
    }

    private parseMessageElement(messageElement: Element, currentUserEmail: string): ThreadMessage | null {
        try {
            // Extract ItemId with proper handling
            const itemIdElements = messageElement.getElementsByTagName('t:ItemId');
            const itemId = itemIdElements.length > 0 ? itemIdElements[0].getAttribute('Id') || '' : '';
            
            // Extract basic message details with robust parsing
            const getElementText = (tagName: string): string => {
                const elements = messageElement.getElementsByTagName(tagName);
                return elements.length > 0 ? (elements[0].textContent || '').trim() : '';
            };
            
            const subject = getElementText('t:Subject');
            const dateTimeSent = getElementText('t:DateTimeSent');
            
            // Extract body with different possible formats
            let body = '';
            const bodyElements = messageElement.getElementsByTagName('t:Body');
            if (bodyElements.length > 0) {
                body = bodyElements[0].textContent || '';
            }
            
            // Extract From address with proper nested parsing
            let fromAddress = currentUserEmail; // Default fallback
            const fromElements = messageElement.getElementsByTagName('t:From');
            if (fromElements.length > 0) {
                // Try to get EmailAddress from nested Mailbox
                const mailboxElements = fromElements[0].getElementsByTagName('t:Mailbox');
                if (mailboxElements.length > 0) {
                    const emailElements = mailboxElements[0].getElementsByTagName('t:EmailAddress');
                    if (emailElements.length > 0) {
                        fromAddress = emailElements[0].textContent || currentUserEmail;
                    }
                } else {
                    // Fallback to direct EmailAddress
                    const emailElements = fromElements[0].getElementsByTagName('t:EmailAddress');
                    if (emailElements.length > 0) {
                        fromAddress = emailElements[0].textContent || currentUserEmail;
                    }
                }
            }
            
            // Extract To recipients with robust parsing
            const toRecipients: string[] = [];
            const toRecipientsElements = messageElement.getElementsByTagName('t:ToRecipients');
            if (toRecipientsElements.length > 0) {
                const mailboxElements = toRecipientsElements[0].getElementsByTagName('t:Mailbox');
                for (let i = 0; i < mailboxElements.length; i++) {
                    const emailElements = mailboxElements[i].getElementsByTagName('t:EmailAddress');
                    if (emailElements.length > 0) {
                        const email = emailElements[0].textContent;
                        if (email && email.trim()) {
                            toRecipients.push(email.trim());
                        }
                    }
                }
            }
            
            // Parse date with fallback
            let sentDate: Date;
            try {
                sentDate = dateTimeSent ? new Date(dateTimeSent) : new Date();
                // Validate the date
                if (isNaN(sentDate.getTime())) {
                    sentDate = new Date();
                }
            } catch (error) {
                console.warn(`[WARN] Failed to parse date: ${dateTimeSent}, using current date`);
                sentDate = new Date();
            }

            // CRITICAL: Normalize email addresses for comparison (case-insensitive and trim whitespace)
            const normalizedFromAddress = fromAddress.toLowerCase().trim();
            const normalizedCurrentUserEmail = currentUserEmail.toLowerCase().trim();
            
            // Determine if this message is from the current user
            const isFromCurrentUser = normalizedFromAddress === normalizedCurrentUserEmail;
            
            const threadMessage: ThreadMessage = {
                id: itemId,
                subject: subject || 'No Subject',
                from: fromAddress,
                to: toRecipients,
                sentDate: sentDate,
                body: body,
                isFromCurrentUser: isFromCurrentUser
            };
            
            console.log(`[DEBUG] ✉️ PARSED MESSAGE: "${threadMessage.subject}"`);
            console.log(`[DEBUG]   From: "${threadMessage.from}" (normalized: "${normalizedFromAddress}")`);
            console.log(`[DEBUG]   Current user: "${currentUserEmail}" (normalized: "${normalizedCurrentUserEmail}")`);
            console.log(`[DEBUG]   Is from current user: ${threadMessage.isFromCurrentUser ? 'YES' : 'NO'}`);
            console.log(`[DEBUG]   Sent: ${threadMessage.sentDate.toISOString()}`);
            
            return threadMessage;
            
        } catch (error) {
            console.error(`[ERROR] 💥 MESSAGE PARSING FAILED: Failed to parse message element:`, error);
            return null;
        }
    }

    // Build an artificial thread when only a single message is available by scanning recent emails
    // for newer messages that quote the base message body and match normalized subject.
    private buildArtificialThreadFromRecentEmails(base: ThreadMessage | undefined, currentUserEmail: string): ThreadMessage[] {
        if (!base) return [];
        if (!this.recentEmailsContext || this.recentEmailsContext.length === 0) return [base];

        const normSubject = this.normalizeSubject(base.subject || '');
        const baseSig = this.normalizeBodyForMatch(base.body || '');
        if (!baseSig || baseSig.length < 20) return [base];

        const baseTime = base.sentDate.getTime();
        const currentLower = (currentUserEmail || '').toLowerCase();
        const baseParticipants = new Set([base.from.toLowerCase(), ...base.to.map(t => (t || '').toLowerCase())]);

        const replies: ThreadMessage[] = [];
        for (const e of this.recentEmailsContext) {
            try {
                const subj = this.normalizeSubject(e.subject || '');
                if (subj !== normSubject) continue;
                const sent = new Date(e.dateTimeSent);
                if (!(sent instanceof Date) || isNaN(sent.getTime()) || sent.getTime() <= baseTime) continue;
                const toList = (e.toRecipients || []).map(r => (r.emailAddress.address || '').toLowerCase());
                const ccList = ((e as any).ccRecipients || []).map((r: any) => (r.emailAddress.address || '').toLowerCase());
                const candRecipients = [...toList, ...ccList];
                const overlap = candRecipients.some(addr => baseParticipants.has(addr));
                if (!overlap && !candRecipients.includes(currentLower)) continue;
                const body = (e.body?.content || '').toString();
                const normBody = this.normalizeBodyForMatch(body);
                if (!normBody || normBody.indexOf(baseSig) === -1) continue;
                const fromAddr = (e.from?.emailAddress?.address || '').toLowerCase();
                replies.push({
                    id: e.id,
                    subject: e.subject,
                    from: fromAddr,
                    to: (e.toRecipients || []).map(r => r.emailAddress.address),
                    sentDate: sent,
                    body: body,
                    isFromCurrentUser: fromAddr === currentLower
                });
            } catch {
                // ignore malformed items
            }
        }
        const thread = [base, ...replies].sort((a, b) => a.sentDate.getTime() - b.sentDate.getTime());
        return thread;
    }

    private normalizeBodyForMatch(text: string): string {
        if (!text) return '';
        let t = text.replace(/<[^>]*>/g, ' ');
        t = t.replace(/^>+\s?/gm, ' ')
             .replace(/from:\s.*\n?/gi, ' ')
             .replace(/sent:\s.*\n?/gi, ' ')
             .replace(/subject:\s.*\n?/gi, ' ')
             .replace(/to:\s.*\n?/gi, ' ')
             .replace(/cc:\s.*\n?/gi, ' ');
        t = t.replace(/\s+/g, ' ').trim().toLowerCase();
        const maxLen = 800;
        return t.length > maxLen ? t.slice(0, maxLen) : t;
    }
}