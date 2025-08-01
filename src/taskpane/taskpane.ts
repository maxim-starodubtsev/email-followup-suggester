import { EmailAnalysisService } from '../services/EmailAnalysisService';
import { ConfigurationService } from '../services/ConfigurationService';
import { LlmService } from '../services/LlmService';
import { RetryService } from '../services/RetryService';
import { FollowupEmail, SnoozeOption } from '../models/FollowupEmail';
import { Configuration } from '../models/Configuration';

class TaskpaneManager {
    private emailAnalysisService: EmailAnalysisService;
    private configurationService: ConfigurationService;
    private llmService?: LlmService;
    private retryService: RetryService;
    private analyzeButton!: HTMLButtonElement;
    private refreshButton!: HTMLButtonElement;
    private settingsButton!: HTMLButtonElement;
    private emailCountSelect!: HTMLSelectElement;
    private daysBackSelect!: HTMLSelectElement;
    private accountFilterSelect!: HTMLSelectElement;
    private enableLlmSummaryCheckbox!: HTMLInputElement;
    private enableLlmSuggestionsCheckbox!: HTMLInputElement;
    private statusDiv!: HTMLDivElement;
    private loadingDiv!: HTMLDivElement;
    private emptyStateDiv!: HTMLDivElement;
    private emailListDiv!: HTMLDivElement;
    
    // Modal elements
    private snoozeModal!: HTMLDivElement;
    private settingsModal!: HTMLDivElement;
    private snoozeOptionsSelect!: HTMLSelectElement;
    private customSnoozeGroup!: HTMLDivElement;
    private customSnoozeDate!: HTMLInputElement;
    private llmEndpointInput!: HTMLInputElement;
    private llmApiKeyInput!: HTMLInputElement;
    private showSnoozedEmailsCheckbox!: HTMLInputElement;
    private showDismissedEmailsCheckbox!: HTMLInputElement;
    
    // New UI element properties
    private statsDashboard!: HTMLDivElement;
    private showStatsButton!: HTMLButtonElement;
    private toggleStatsButton!: HTMLButtonElement;
    private advancedFilters!: HTMLDivElement;
    private toggleAdvancedFiltersButton!: HTMLButtonElement;
    private threadModal!: HTMLDivElement;
    private threadSubject!: HTMLHeadingElement;
    private threadBody!: HTMLDivElement;
    
    // Statistics elements
    private totalEmailsAnalyzedSpan!: HTMLSpanElement;
    private needingFollowupSpan!: HTMLSpanElement;
    private highPriorityCountSpan!: HTMLSpanElement;
    private avgResponseTimeSpan!: HTMLSpanElement;
    
    // Filter elements
    private priorityFilter!: HTMLSelectElement;
    private responseTimeFilter!: HTMLSelectElement;
    private subjectFilter!: HTMLInputElement;
    private senderFilter!: HTMLInputElement;
    private aiSuggestionFilter!: HTMLSelectElement;
    private clearFiltersButton!: HTMLButtonElement;
    
    // Progress elements
    private loadingStep!: HTMLSpanElement;
    private loadingDetail!: HTMLDivElement;
    private progressFill!: HTMLDivElement;
    
    private currentEmailForSnooze: string = '';
    private availableAccounts: string[] = [];
    private allEmails: FollowupEmail[] = [];
    private filteredEmails: FollowupEmail[] = [];

    constructor() {
        this.retryService = new RetryService();
        this.emailAnalysisService = new EmailAnalysisService();
        this.configurationService = new ConfigurationService();
        this.initializeElements();
        this.attachEventListeners();
        this.loadCachedResults();
    }

    private initializeElements(): void {
        // Main controls
        this.analyzeButton = document.getElementById('analyzeButton') as HTMLButtonElement;
        this.refreshButton = document.getElementById('refreshButton') as HTMLButtonElement;
        this.settingsButton = document.getElementById('settingsButton') as HTMLButtonElement;
        this.emailCountSelect = document.getElementById('emailCount') as HTMLSelectElement;
        this.daysBackSelect = document.getElementById('daysBack') as HTMLSelectElement;
        this.accountFilterSelect = document.getElementById('accountFilter') as HTMLSelectElement;
        this.enableLlmSummaryCheckbox = document.getElementById('enableLlmSummary') as HTMLInputElement;
        this.enableLlmSuggestionsCheckbox = document.getElementById('enableLlmSuggestions') as HTMLInputElement;
        
        // Display elements
        this.statusDiv = document.getElementById('status') as HTMLDivElement;
        this.loadingDiv = document.getElementById('loadingMessage') as HTMLDivElement;
        this.emptyStateDiv = document.getElementById('emptyState') as HTMLDivElement;
        this.emailListDiv = document.getElementById('emailList') as HTMLDivElement;
        
        // Modal elements
        this.snoozeModal = document.getElementById('snoozeModal') as HTMLDivElement;
        this.settingsModal = document.getElementById('settingsModal') as HTMLDivElement;
        this.snoozeOptionsSelect = document.getElementById('snoozeOptions') as HTMLSelectElement;
        this.customSnoozeGroup = document.getElementById('customSnoozeGroup') as HTMLDivElement;
        this.customSnoozeDate = document.getElementById('customSnoozeDate') as HTMLInputElement;
        this.llmEndpointInput = document.getElementById('llmEndpoint') as HTMLInputElement;
        this.llmApiKeyInput = document.getElementById('llmApiKey') as HTMLInputElement;
        this.showSnoozedEmailsCheckbox = document.getElementById('showSnoozedEmails') as HTMLInputElement;
        this.showDismissedEmailsCheckbox = document.getElementById('showDismissedEmails') as HTMLInputElement;
        
        // New UI elements
        this.statsDashboard = document.getElementById('statsDashboard') as HTMLDivElement;
        this.showStatsButton = document.getElementById('showStatsButton') as HTMLButtonElement;
        this.toggleStatsButton = document.getElementById('toggleStats') as HTMLButtonElement;
        this.advancedFilters = document.getElementById('advancedFilters') as HTMLDivElement;
        this.toggleAdvancedFiltersButton = document.getElementById('toggleAdvancedFilters') as HTMLButtonElement;
        this.threadModal = document.getElementById('threadModal') as HTMLDivElement;
        this.threadSubject = document.getElementById('threadSubject') as HTMLHeadingElement;
        this.threadBody = document.getElementById('threadBody') as HTMLDivElement;
        
        // Statistics elements
        this.totalEmailsAnalyzedSpan = document.getElementById('totalEmailsAnalyzed') as HTMLSpanElement;
        this.needingFollowupSpan = document.getElementById('needingFollowup') as HTMLSpanElement;
        this.highPriorityCountSpan = document.getElementById('highPriorityCount') as HTMLSpanElement;
        this.avgResponseTimeSpan = document.getElementById('avgResponseTime') as HTMLSpanElement;
        
        // Filter elements
        this.priorityFilter = document.getElementById('priorityFilter') as HTMLSelectElement;
        this.responseTimeFilter = document.getElementById('responseTimeFilter') as HTMLSelectElement;
        this.subjectFilter = document.getElementById('subjectFilter') as HTMLInputElement;
        this.senderFilter = document.getElementById('senderFilter') as HTMLInputElement;
        this.aiSuggestionFilter = document.getElementById('aiSuggestionFilter') as HTMLSelectElement;
        this.clearFiltersButton = document.getElementById('clearFilters') as HTMLButtonElement;
        
        // Progress elements
        this.loadingStep = document.getElementById('loadingStep') as HTMLSpanElement;
        this.loadingDetail = document.getElementById('loadingDetail') as HTMLDivElement;
        this.progressFill = document.getElementById('progressFill') as HTMLDivElement;
    }

    private attachEventListeners(): void {
        // Main controls
        this.analyzeButton.addEventListener('click', () => this.analyzeEmails());
        this.refreshButton.addEventListener('click', () => this.refreshEmails());
        this.settingsButton.addEventListener('click', () => this.showSettingsModal());
        
        // Save configuration when changed
        this.emailCountSelect.addEventListener('change', () => this.saveConfiguration());
        this.daysBackSelect.addEventListener('change', () => this.saveConfiguration());
        this.accountFilterSelect.addEventListener('change', () => this.saveConfiguration());
        this.enableLlmSummaryCheckbox.addEventListener('change', () => this.saveConfiguration());
        this.enableLlmSuggestionsCheckbox.addEventListener('change', () => this.saveConfiguration());
        
        // New UI event listeners
        this.showStatsButton.addEventListener('click', () => this.toggleStatsDashboard(true));
        this.toggleStatsButton.addEventListener('click', () => this.toggleStatsDashboard(false));
        this.toggleAdvancedFiltersButton.addEventListener('click', () => this.toggleAdvancedFilters());
        this.clearFiltersButton.addEventListener('click', () => this.clearAllFilters());
        
        // Filter event listeners
        this.priorityFilter.addEventListener('change', () => this.applyFilters());
        this.responseTimeFilter.addEventListener('change', () => this.applyFilters());
        this.subjectFilter.addEventListener('input', () => this.debounceApplyFilters());
        this.senderFilter.addEventListener('input', () => this.debounceApplyFilters());
        this.aiSuggestionFilter.addEventListener('change', () => this.applyFilters());
        
        // Modal controls
        this.attachModalEventListeners();
        
        // Thread modal event listeners
        this.threadModal.addEventListener('click', (e) => {
            if (e.target === this.threadModal) this.hideThreadModal();
        });
    }

    private attachModalEventListeners(): void {
        // Snooze modal
        document.getElementById('confirmSnooze')?.addEventListener('click', () => this.confirmSnooze());
        document.getElementById('cancelSnooze')?.addEventListener('click', () => this.hideSnoozeModal());
        this.snoozeOptionsSelect.addEventListener('change', () => this.handleSnoozeOptionChange());
        
        // Settings modal
        document.getElementById('saveSettings')?.addEventListener('click', () => this.saveSettings());
        document.getElementById('cancelSettings')?.addEventListener('click', () => this.hideSettingsModal());
        
        // Close modals when clicking outside or on close button
        document.querySelectorAll('.modal-close').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const modal = (e.target as Element).closest('.modal') as HTMLDivElement;
                if (modal) modal.style.display = 'none';
            });
        });
        
        window.addEventListener('click', (e) => {
            if (e.target === this.snoozeModal) this.hideSnoozeModal();
            if (e.target === this.settingsModal) this.hideSettingsModal();
        });
    }

    private async loadConfiguration(): Promise<void> {
        try {
            const config = await this.configurationService.getConfiguration();
            this.emailCountSelect.value = config.emailCount.toString();
            this.daysBackSelect.value = config.daysBack.toString();
            this.enableLlmSummaryCheckbox.checked = config.enableLlmSummary;
            this.enableLlmSuggestionsCheckbox.checked = config.enableLlmSuggestions;
            this.showSnoozedEmailsCheckbox.checked = config.showSnoozedEmails;
            this.showDismissedEmailsCheckbox.checked = config.showDismissedEmails;
            
            if (config.llmApiEndpoint) this.llmEndpointInput.value = config.llmApiEndpoint;
            if (config.llmApiKey) this.llmApiKeyInput.value = config.llmApiKey;
            
            // Load accounts and populate filter
            await this.loadAvailableAccounts();
            this.populateAccountFilter(config.selectedAccounts);
            
            // Initialize LLM service if configured
            if (config.llmApiEndpoint && config.llmApiKey) {
                this.llmService = new LlmService(config, this.retryService);
                this.emailAnalysisService.setLlmService(this.llmService);
            }
            
            // Populate snooze options
            this.populateSnoozeOptions(config.snoozeOptions);
        } catch (error) {
            console.error('Error loading configuration:', error);
        }
    }

    private async loadAvailableAccounts(): Promise<void> {
        try {
            this.availableAccounts = await this.configurationService.getAvailableAccounts();
        } catch (error) {
            console.error('Error loading available accounts:', error);
            this.availableAccounts = [];
        }
    }

    private populateAccountFilter(selectedAccounts: string[]): void {
        this.accountFilterSelect.innerHTML = '';
        
        this.availableAccounts.forEach(account => {
            const option = document.createElement('option');
            option.value = account;
            option.textContent = account;
            option.selected = selectedAccounts.includes(account);
            this.accountFilterSelect.appendChild(option);
        });
    }

    private populateSnoozeOptions(options: SnoozeOption[]): void {
        this.snoozeOptionsSelect.innerHTML = '';
        
        options.forEach(option => {
            const optionElement = document.createElement('option');
            optionElement.value = option.value.toString();
            optionElement.textContent = option.label;
            optionElement.dataset.isCustom = option.isCustom?.toString() || 'false';
            this.snoozeOptionsSelect.appendChild(optionElement);
        });
    }

    private async saveConfiguration(): Promise<void> {
        try {
            const config: Configuration = {
                emailCount: parseInt(this.emailCountSelect.value),
                daysBack: parseInt(this.daysBackSelect.value),
                lastAnalysisDate: new Date(),
                enableLlmSummary: this.enableLlmSummaryCheckbox.checked,
                enableLlmSuggestions: this.enableLlmSuggestionsCheckbox.checked,
                llmApiEndpoint: this.llmEndpointInput.value.trim(),
                llmApiKey: this.llmApiKeyInput.value.trim(),
                showSnoozedEmails: this.showSnoozedEmailsCheckbox.checked,
                showDismissedEmails: this.showDismissedEmailsCheckbox.checked,
                selectedAccounts: Array.from(this.accountFilterSelect.selectedOptions).map(option => option.value),
                snoozeOptions: [] // Populate this if needed
            };
            await this.configurationService.saveConfiguration(config);
        } catch (error) {
            console.error('Error saving configuration:', error);
        }
    }

    // Enhanced analyzeEmails with progress tracking
    private async analyzeEmails(): Promise<void> {
        try {
            this.setLoadingState(true);
            this.hideStatus();
            this.updateProgress(0, 'Initializing analysis...', 'Getting ready to analyze your emails');

            const emailCount = parseInt(this.emailCountSelect.value);
            const daysBack = parseInt(this.daysBackSelect.value);
            const selectedAccounts = Array.from(this.accountFilterSelect.selectedOptions).map(option => option.value);

            this.updateProgress(20, 'Fetching emails...', `Looking for emails from the last ${daysBack} days`);
            
            // Simulate progress updates during analysis
            const progressInterval = setInterval(() => {
                const currentWidth = parseInt(this.progressFill.style.width) || 20;
                if (currentWidth < 80) {
                    this.updateProgress(currentWidth + 10, 'Analyzing emails...', 'Processing email content and threads');
                }
            }, 500);

            const followupEmails = await this.emailAnalysisService.analyzeEmails(emailCount, daysBack, selectedAccounts);
            
            clearInterval(progressInterval);
            this.updateProgress(100, 'Analysis complete!', 'Preparing results');
            
            this.allEmails = followupEmails;
            this.filteredEmails = [...followupEmails];
            
            await this.saveConfiguration();
            this.updateStatistics();
            this.displayEmails(this.filteredEmails);
            
            if (followupEmails.length > 0) {
                this.showStatus(`Found ${followupEmails.length} email(s) that may need follow-up`, 'success');
            } else {
                this.showStatus('No emails requiring follow-up found', 'success');
            }
        } catch (error) {
            console.error('Error analyzing emails:', (error as Error).message);
            this.showStatus(`Error analyzing emails: ${(error as Error).message}`, 'error');
            this.displayEmails([]);
        } finally {
            this.setLoadingState(false);
        }
    }

    // Progress tracking methods
    private updateProgress(percentage: number, step: string, detail: string): void {
        this.progressFill.style.width = `${percentage}%`;
        this.loadingStep.textContent = step;
        this.loadingDetail.textContent = detail;
    }

    // Statistics methods
    private updateStatistics(): void {
        const total = this.allEmails.length;
        const needingFollowup = this.filteredEmails.length;
        const highPriority = this.allEmails.filter(email => email.priority === 'high').length;
        const avgDays = total > 0 ? Math.round(
            this.allEmails.reduce((sum, email) => sum + email.daysWithoutResponse, 0) / total
        ) : 0;

        this.totalEmailsAnalyzedSpan.textContent = total.toString();
        this.needingFollowupSpan.textContent = needingFollowup.toString();
        this.highPriorityCountSpan.textContent = highPriority.toString();
        this.avgResponseTimeSpan.textContent = avgDays.toString();
    }

    private toggleStatsDashboard(show: boolean): void {
        if (show) {
            this.statsDashboard.classList.add('show');
            this.showStatsButton.style.display = 'none';
        } else {
            this.statsDashboard.classList.remove('show');
            this.showStatsButton.style.display = 'inline-block';
        }
    }

    private toggleAdvancedFilters(): void {
        this.advancedFilters.classList.toggle('show');
        const isShown = this.advancedFilters.classList.contains('show');
        this.toggleAdvancedFiltersButton.textContent = isShown ? 'Hide Advanced Filters' : 'Advanced Filters';
    }

    // Filter methods
    private debounceTimer?: number;
    
    private debounceApplyFilters(): void {
        if (this.debounceTimer) {
            clearTimeout(this.debounceTimer);
        }
        this.debounceTimer = window.setTimeout(() => this.applyFilters(), 300);
    }

    private applyFilters(): void {
        this.filteredEmails = this.allEmails.filter(email => {
            // Priority filter
            if (this.priorityFilter.value && email.priority !== this.priorityFilter.value) {
                return false;
            }

            // Response time filter
            if (this.responseTimeFilter.value) {
                const days = email.daysWithoutResponse;
                const range = this.responseTimeFilter.value;
                
                if (range === '1-3' && (days < 1 || days > 3)) return false;
                if (range === '4-7' && (days < 4 || days > 7)) return false;
                if (range === '8-14' && (days < 8 || days > 14)) return false;
                if (range === '15+' && days < 15) return false;
            }

            // Subject filter
            if (this.subjectFilter.value && 
                !email.subject.toLowerCase().includes(this.subjectFilter.value.toLowerCase())) {
                return false;
            }

            // Sender filter (check recipients for sent emails)
            if (this.senderFilter.value) {
                const senderText = this.senderFilter.value.toLowerCase();
                const hasMatchingRecipient = email.recipients.some(recipient => 
                    recipient.toLowerCase().includes(senderText)
                );
                if (!hasMatchingRecipient) return false;
            }

            // AI suggestion filter
            if (this.aiSuggestionFilter.value) {
                const hasAI = !!email.llmSuggestion;
                if (this.aiSuggestionFilter.value === 'with-ai' && !hasAI) return false;
                if (this.aiSuggestionFilter.value === 'without-ai' && hasAI) return false;
            }

            return true;
        });

        this.displayEmails(this.filteredEmails);
        this.updateStatistics();
    }

    private clearAllFilters(): void {
        this.priorityFilter.value = '';
        this.responseTimeFilter.value = '';
        this.subjectFilter.value = '';
        this.senderFilter.value = '';
        this.aiSuggestionFilter.value = '';
        this.applyFilters();
    }

    // Enhanced email display with priority classes and metadata
    private createEmailElement(email: FollowupEmail): HTMLDivElement {
        const emailDiv = document.createElement('div');
        emailDiv.className = `email-item priority-${email.priority}`;

        const priorityBadge = email.priority === 'high' ? '🔴' : email.priority === 'medium' ? '🟡' : '🟢';
        const accountBadge = email.accountEmail ? `📧 ${email.accountEmail}` : '';
        const llmIndicator = email.llmSummary ? '🤖' : '';

        // Calculate confidence score (mock implementation)
        const confidence = Math.min(100, Math.max(0, 
            (email.daysWithoutResponse * 10) + 
            (email.priority === 'high' ? 30 : email.priority === 'medium' ? 15 : 0) +
            (email.llmSuggestion ? 20 : 0)
        ));

        emailDiv.innerHTML = `
            <div class="email-header">
                <span class="priority-badge">${priorityBadge}</span>
                <span class="account-badge">${accountBadge}</span>
                <span class="llm-indicator">${llmIndicator}</span>
            </div>
            <div class="email-subject">${this.escapeHtml(email.subject)}</div>
            <div class="email-metadata">
                <div class="metadata-item">
                    <span>📤 To: ${this.escapeHtml(email.recipients.join(', '))}</span>
                </div>
                <div class="metadata-item">
                    <span>📅 ${email.sentDate.toLocaleDateString()}</span>
                </div>
                <div class="metadata-item">
                    <span>⏱️ ${email.daysWithoutResponse} days</span>
                </div>
            </div>
            <div class="confidence-meter">
                <span>Priority Score:</span>
                <div class="confidence-bar">
                    <div class="confidence-fill" style="width: ${confidence}%"></div>
                </div>
                <span>${confidence}%</span>
            </div>
            <div class="email-summary">${this.escapeHtml(email.summary)}</div>
            ${email.llmSuggestion ? `<div class="llm-suggestion"><strong>AI Suggestion:</strong> ${this.escapeHtml(email.llmSuggestion)}</div>` : ''}
            <div class="email-actions">
                <button class="action-button" data-email-id="${email.id}" data-action="reply">Reply</button>
                <button class="action-button" data-email-id="${email.id}" data-action="forward">Forward</button>
                <button class="action-button" data-email-id="${email.id}" data-action="snooze">Snooze</button>
                <button class="action-button" data-email-id="${email.id}" data-action="dismiss">Dismiss</button>
                ${email.threadMessages.length > 1 ? `<button class="action-button" data-email-id="${email.id}" data-action="view-thread">View Thread (${email.threadMessages.length})</button>` : ''}
            </div>
        `;

        // Attach event listeners to action buttons
        const actionButtons = emailDiv.querySelectorAll('.action-button');
        actionButtons.forEach(button => {
            button.addEventListener('click', (e) => this.handleEmailAction(e));
        });

        return emailDiv;
    }

    // Handle email action buttons
    private handleEmailAction(event: Event): void {
        const button = event.target as HTMLButtonElement;
        const emailId = button.dataset.emailId!;
        const action = button.dataset.action!;

        switch (action) {
            case 'reply':
                this.replyToEmail(emailId);
                break;
            case 'forward':
                this.forwardEmail(emailId);
                break;
            case 'snooze':
                this.showSnoozeModal(emailId);
                break;
            case 'dismiss':
                this.dismissEmail(emailId);
                break;
            case 'view-thread':
                this.showThreadView(emailId);
                break;
        }
    }

    // Thread view methods
    private showThreadView(emailId: string): void {
        const email = this.allEmails.find(e => e.id === emailId);
        if (!email) return;

        this.threadSubject.textContent = `Thread: ${email.subject}`;
        this.threadBody.innerHTML = '';

        email.threadMessages.forEach((message, index) => {
            const messageDiv = document.createElement('div');
            messageDiv.className = 'thread-message';
            
            messageDiv.innerHTML = `
                <div class="thread-message-header">
                    <strong>Message ${index + 1}</strong> - ${message.sentDate.toLocaleDateString()} ${message.sentDate.toLocaleTimeString()}
                    <br>From: ${this.escapeHtml(message.from)} | To: ${this.escapeHtml(message.to.join(', '))}
                </div>
                <div class="thread-message-body">
                    ${this.escapeHtml(message.body)}
                </div>
            `;
            
            this.threadBody.appendChild(messageDiv);
        });

        this.threadModal.style.display = 'block';
    }

    private hideThreadModal(): void {
        this.threadModal.style.display = 'none';
    }

    // Load cached results from ribbon commands
    private async loadCachedResults(): Promise<void> {
        try {
            const cachedEmails = await this.configurationService.getCachedAnalysisResults();
            if (cachedEmails.length > 0) {
                this.allEmails = cachedEmails;
                this.filteredEmails = [...cachedEmails];
                this.updateStatistics();
                this.displayEmails(this.filteredEmails);
                this.showStatus(`Loaded ${cachedEmails.length} cached results`, 'success');
            }
        } catch (error) {
            console.error('Error loading cached results:', error);
        }
    }

    private async replyToEmail(_emailId: string): Promise<void> {
        // Use Office.js to compose a reply
        Office.context.mailbox.displayNewMessageForm({
            toRecipients: [], // Will be populated based on original email
            subject: 'Re: ',
            htmlBody: '<br><br>',
            attachments: []
        });
    }

    private async forwardEmail(_emailId: string): Promise<void> {
        // Use Office.js to compose a forward
        Office.context.mailbox.displayNewMessageForm({
            toRecipients: [],
            subject: 'Fwd: ',
            htmlBody: '<br><br>',
            attachments: []
        });
    }

    // Modal management methods
    private showSnoozeModal(emailId: string): void {
        this.currentEmailForSnooze = emailId;
        this.snoozeModal.style.display = 'block';
    }

    private hideSnoozeModal(): void {
        this.snoozeModal.style.display = 'none';
        this.currentEmailForSnooze = '';
        this.customSnoozeGroup.style.display = 'none';
    }

    private showSettingsModal(): void {
        this.settingsModal.style.display = 'block';
    }

    private hideSettingsModal(): void {
        this.settingsModal.style.display = 'none';
    }

    private handleSnoozeOptionChange(): void {
        const selectedOption = this.snoozeOptionsSelect.selectedOptions[0];
        const isCustom = selectedOption?.dataset.isCustom === 'true';
        
        if (isCustom) {
            this.customSnoozeGroup.style.display = 'block';
            // Set default to 1 hour from now
            const defaultDate = new Date();
            defaultDate.setHours(defaultDate.getHours() + 1);
            this.customSnoozeDate.value = defaultDate.toISOString().slice(0, 16);
        } else {
            this.customSnoozeGroup.style.display = 'none';
        }
    }

    private async confirmSnooze(): Promise<void> {
        if (!this.currentEmailForSnooze) return;

        const selectedOption = this.snoozeOptionsSelect.selectedOptions[0];
        const isCustom = selectedOption?.dataset.isCustom === 'true';
        
        try {
            if (isCustom) {
                const customDate = new Date(this.customSnoozeDate.value);
                if (customDate <= new Date()) {
                    this.showStatus('Snooze time must be in the future', 'error');
                    return;
                }
                this.emailAnalysisService.snoozeEmailUntil(this.currentEmailForSnooze, customDate);
            } else {
                const minutes = parseInt(selectedOption.value);
                this.emailAnalysisService.snoozeEmail(this.currentEmailForSnooze, minutes);
            }

            // Remove email from current display
            this.removeEmailFromDisplay(this.currentEmailForSnooze);
            this.showStatus('Email snoozed successfully', 'success');
            this.hideSnoozeModal();
        } catch (error) {
            console.error('Error snoozing email:', (error as Error).message);
            this.showStatus(`Error snoozing email: ${(error as Error).message}`, 'error');
        }
    }

    private async dismissEmail(emailId: string): Promise<void> {
        this.emailAnalysisService.dismissEmail(emailId);
        this.removeEmailFromDisplay(emailId);
        this.showStatus('Email dismissed', 'success');
    }

    private removeEmailFromDisplay(emailId: string): void {
        const emailElement = document.querySelector(`[data-email-id="${emailId}"]`)?.closest('.email-item');
        if (emailElement) {
            emailElement.remove();
            
            // Check if list is now empty
            if (this.emailListDiv.children.length === 0) {
                this.hideAllStates();
                this.emptyStateDiv.style.display = 'block';
            }
        }
    }

    private async saveSettings(): Promise<void> {
        try {
            const endpoint = this.llmEndpointInput.value.trim();
            const apiKey = this.llmApiKeyInput.value.trim();
            const enableSummary = this.showSnoozedEmailsCheckbox.checked;
            const enableSuggestions = this.showDismissedEmailsCheckbox.checked;

            if (endpoint && apiKey) {
                await this.configurationService.updateLlmSettings(endpoint, apiKey, enableSummary, enableSuggestions);
                
                // Update LLM service with proper configuration object
                const config = await this.configurationService.getConfiguration();
                this.llmService = new LlmService(config, this.retryService);
                this.emailAnalysisService.setLlmService(this.llmService);
                
                this.showStatus('Settings saved successfully', 'success');
            } else if (endpoint || apiKey) {
                this.showStatus('Both endpoint and API key are required for LLM integration', 'error');
                return;
            }

            this.hideSettingsModal();
        } catch (error) {
            console.error('Error saving settings:', (error as Error).message);
            this.showStatus(`Error saving settings: ${(error as Error).message}`, 'error');
        }
    }

    private showStatus(message: string, type: 'success' | 'error'): void {
        this.statusDiv.textContent = message;
        this.statusDiv.className = `status ${type}`;
        this.statusDiv.style.display = 'block';
        
        // Auto-hide success messages after 5 seconds
        if (type === 'success') {
            setTimeout(() => this.hideStatus(), 5000);
        }
    }

    private hideStatus(): void {
        this.statusDiv.style.display = 'none';
    }

    private escapeHtml(text: string): string {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    // Refresh emails (re-run analysis)
    private async refreshEmails(): Promise<void> {
        await this.analyzeEmails();
    }

    // Loading state management
    private setLoadingState(isLoading: boolean): void {
        if (isLoading) {
            this.hideAllStates();
            this.loadingDiv.style.display = 'block';
            this.analyzeButton.disabled = true;
            this.refreshButton.disabled = true;
        } else {
            this.loadingDiv.style.display = 'none';
            this.analyzeButton.disabled = false;
            this.refreshButton.disabled = false;
        }
    }

    // Hide all display states
    private hideAllStates(): void {
        this.loadingDiv.style.display = 'none';
        this.emptyStateDiv.style.display = 'none';
        this.emailListDiv.style.display = 'none';
    }

    // Initialize the task pane
    public async initialize(): Promise<void> {
        await this.loadConfiguration();
        this.hideAllStates();
        this.emptyStateDiv.style.display = 'block';
    }

    // Display emails with enhanced UI
    private displayEmails(emails: FollowupEmail[]): void {
        this.hideAllStates();
        
        if (emails.length === 0) {
            this.emptyStateDiv.innerHTML = `
                <h3>No emails need follow-up</h3>
                <p>Great! You're all caught up with your email responses.</p>
            `;
            this.emptyStateDiv.style.display = 'block';
            return;
        }

        this.emailListDiv.innerHTML = '';
        emails.forEach(email => {
            const emailElement = this.createEmailElement(email);
            this.emailListDiv.appendChild(emailElement);
        });
        
        this.emailListDiv.style.display = 'block';
    }
}

// Initialize the task pane when Office is ready
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        const taskpaneManager = new TaskpaneManager();
        taskpaneManager.initialize().catch(console.error);
    }
});