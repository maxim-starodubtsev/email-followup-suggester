import { EmailAnalysisService } from '../services/EmailAnalysisService';
import { ConfigurationService } from '../services/ConfigurationService';
import { LlmService } from '../services/LlmService';
import { RetryService } from '../services/RetryService';
import { Configuration } from '../models/Configuration';
import { FollowupEmail } from '../models/FollowupEmail';

interface SnoozeOption {
    label: string;
    value: number;
    isCustom?: boolean;
}

export class TaskpaneManager {
    private emailAnalysisService: EmailAnalysisService;
    private configurationService: ConfigurationService;
    private llmService?: LlmService;
    private retryService: RetryService;
    
    // Main controls
    private analyzeButton!: HTMLButtonElement;
    private refreshButton!: HTMLButtonElement;
    private settingsButton!: HTMLButtonElement;
    private emailCountSelect!: HTMLSelectElement;
    private daysBackSelect!: HTMLSelectElement;
    private accountFilterSelect!: HTMLSelectElement;
    private enableLlmSummaryCheckbox!: HTMLInputElement;
    private enableLlmSuggestionsCheckbox!: HTMLInputElement;
    
    // Display elements
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
    private aiStatusDiv!: HTMLDivElement;
    private aiStatusText!: HTMLSpanElement;
    private testAiConnectionButton!: HTMLButtonElement;
    private disableAiFeaturesButton!: HTMLButtonElement;
    private enableAiFeaturesCheckbox!: HTMLInputElement;
    private llmProviderSelect!: HTMLSelectElement;
    private llmModelInput!: HTMLInputElement;
    private llmDeploymentNameInput!: HTMLInputElement;
    private llmApiVersionInput!: HTMLInputElement;
    private azureSpecificOptions!: HTMLDivElement;
    
    // Enhanced UI elements
    private statsDashboard!: HTMLDivElement;
    private showStatsButton!: HTMLButtonElement;
    private toggleStatsButton!: HTMLButtonElement;
    
    // Debug logging
    private debugEnabled: boolean = true;
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
        this.initializeDebugLogging();
    }

    private initializeDebugLogging(): void {
        if (this.debugEnabled) {
            console.log('%c[Followup Suggester] Debug logging enabled', 'color: #2196F3; font-weight: bold;');
            console.log('%cTo see detailed debug output, click "Analyze Emails" and watch this console', 'color: #666;');
        }
    }

    private debugLog(message: string, data?: any): void {
        if (this.debugEnabled) {
            console.log(`%c[Followup Suggester] ${message}`, 'color: #2196F3;', data || '');
        }
    }

    private initializeElements(): void {
        const safeElement = (id: string): any => {
            const el = document.getElementById(id);
            if (el) return el;
            // Create a lightweight stub element for unit tests / headless environments
            return {
                id,
                style: {},
                children: [],
                innerHTML: '',
                value: '',
                disabled: false,
                checked: false,
                textContent: '',
                addEventListener: () => {},
                removeEventListener: () => {},
                appendChild: () => {},
                querySelector: () => null,
                querySelectorAll: () => [],
                classList: { add: () => {}, remove: () => {}, toggle: () => {}, contains: () => false }
            } as any;
        };
        // Main controls
    this.analyzeButton = safeElement('analyzeButton') as HTMLButtonElement;
    this.refreshButton = safeElement('refreshButton') as HTMLButtonElement;
    this.settingsButton = safeElement('settingsButton') as HTMLButtonElement;
    this.emailCountSelect = safeElement('emailCount') as HTMLSelectElement;
    this.daysBackSelect = safeElement('daysBack') as HTMLSelectElement;
    this.accountFilterSelect = safeElement('accountFilter') as HTMLSelectElement;
    this.enableLlmSummaryCheckbox = safeElement('enableLlmSummary') as HTMLInputElement;
    this.enableLlmSuggestionsCheckbox = safeElement('enableLlmSuggestions') as HTMLInputElement;
        
        // Display elements
    this.statusDiv = safeElement('status') as HTMLDivElement;
    this.loadingDiv = safeElement('loadingMessage') as HTMLDivElement;
    this.emptyStateDiv = safeElement('emptyState') as HTMLDivElement;
    this.emailListDiv = safeElement('emailList') as HTMLDivElement;
        
        // Modal elements
    this.snoozeModal = safeElement('snoozeModal') as HTMLDivElement;
    this.settingsModal = safeElement('settingsModal') as HTMLDivElement;
    this.snoozeOptionsSelect = safeElement('snoozeOptions') as HTMLSelectElement;
    this.customSnoozeGroup = safeElement('customSnoozeGroup') as HTMLDivElement;
    this.customSnoozeDate = safeElement('customSnoozeDate') as HTMLInputElement;
    this.llmEndpointInput = safeElement('llmEndpoint') as HTMLInputElement;
    this.llmApiKeyInput = safeElement('llmApiKey') as HTMLInputElement;
    this.showSnoozedEmailsCheckbox = safeElement('showSnoozedEmails') as HTMLInputElement;
    this.showDismissedEmailsCheckbox = safeElement('showDismissedEmails') as HTMLInputElement;
    this.aiStatusDiv = safeElement('aiStatus') as HTMLDivElement;
    this.aiStatusText = safeElement('aiStatusText') as HTMLSpanElement;
    this.testAiConnectionButton = safeElement('testAiConnection') as HTMLButtonElement;
    this.disableAiFeaturesButton = safeElement('disableAiFeatures') as HTMLButtonElement;
    this.enableAiFeaturesCheckbox = safeElement('enableAiFeatures') as HTMLInputElement;
    this.llmProviderSelect = safeElement('llmProvider') as HTMLSelectElement;
    this.llmModelInput = safeElement('llmModel') as HTMLInputElement;
    this.llmDeploymentNameInput = safeElement('llmDeploymentName') as HTMLInputElement;
    this.llmApiVersionInput = safeElement('llmApiVersion') as HTMLInputElement;
    this.azureSpecificOptions = safeElement('azureSpecificOptions') as HTMLDivElement;
        
        // New UI elements
    this.statsDashboard = safeElement('statsDashboard') as HTMLDivElement;
    this.showStatsButton = safeElement('showStatsButton') as HTMLButtonElement;
    this.toggleStatsButton = safeElement('toggleStats') as HTMLButtonElement;
    this.advancedFilters = safeElement('advancedFilters') as HTMLDivElement;
    this.toggleAdvancedFiltersButton = safeElement('toggleAdvancedFilters') as HTMLButtonElement;
    this.threadModal = safeElement('threadModal') as HTMLDivElement;
    this.threadSubject = safeElement('threadSubject') as HTMLHeadingElement;
    this.threadBody = safeElement('threadBody') as HTMLDivElement;
        
        // Statistics elements
    this.totalEmailsAnalyzedSpan = safeElement('totalEmailsAnalyzed') as HTMLSpanElement;
    this.needingFollowupSpan = safeElement('needingFollowup') as HTMLSpanElement;
    this.highPriorityCountSpan = safeElement('highPriorityCount') as HTMLSpanElement;
    this.avgResponseTimeSpan = safeElement('avgResponseTime') as HTMLSpanElement;
        
        // Filter elements
    this.priorityFilter = safeElement('priorityFilter') as HTMLSelectElement;
    this.responseTimeFilter = safeElement('responseTimeFilter') as HTMLSelectElement;
    this.subjectFilter = safeElement('subjectFilter') as HTMLInputElement;
    this.senderFilter = safeElement('senderFilter') as HTMLInputElement;
    this.aiSuggestionFilter = safeElement('aiSuggestionFilter') as HTMLSelectElement;
    this.clearFiltersButton = safeElement('clearFilters') as HTMLButtonElement;
        
        // Progress elements
    this.loadingStep = safeElement('loadingStep') as HTMLSpanElement;
    this.loadingDetail = safeElement('loadingDetail') as HTMLDivElement;
    this.progressFill = safeElement('progressFill') as HTMLDivElement;
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
        this.testAiConnectionButton.addEventListener('click', () => this.testAiConnection());
        this.disableAiFeaturesButton.addEventListener('click', () => this.disableAiFeatures());
        this.enableAiFeaturesCheckbox.addEventListener('change', () => this.toggleAiFeatures());
        this.llmProviderSelect.addEventListener('change', () => this.handleProviderChange());
        
        // Close modals when clicking outside or on close button
        document.addEventListener('click', (e) => {
            const target = e.target as Element;
            if (target.classList.contains('modal-close')) {
                const modal = target.closest('.modal') as HTMLDivElement;
                if (modal) modal.style.display = 'none';
            }
        });
        
        window.addEventListener('click', (e) => {
            if (e.target === this.snoozeModal) this.hideSnoozeModal();
            if (e.target === this.settingsModal) this.hideSettingsModal();
            // Handle diagnostic modal close on outside click
            const diagnosticModal = document.getElementById('diagnosticModal');
            if (diagnosticModal && e.target === diagnosticModal) {
                diagnosticModal.style.display = 'none';
            }
        });
        
        // Handle escape key for modal closing
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                const diagnosticModal = document.getElementById('diagnosticModal');
                if (diagnosticModal && diagnosticModal.style.display === 'block') {
                    diagnosticModal.style.display = 'none';
                }
            }
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
            if (config.llmModel) this.llmModelInput.value = config.llmModel;
            if (config.llmProvider) this.llmProviderSelect.value = config.llmProvider;
            if (config.llmDeploymentName) this.llmDeploymentNameInput.value = config.llmDeploymentName;
            if (config.llmApiVersion) this.llmApiVersionInput.value = config.llmApiVersion;
            
            // Show/hide Azure-specific options
            this.handleProviderChange();
            
            // Load accounts and populate filter
            await this.loadAvailableAccounts();
            this.populateAccountFilter(config.selectedAccounts);
            
            // Set configuration in EmailAnalysisService
            this.emailAnalysisService.setConfiguration(config);
            
            // Initialize LLM service if configured
            if (config.llmApiEndpoint && config.llmApiKey) {
                this.llmService = new LlmService(config, this.retryService);
                this.emailAnalysisService.setLlmService(this.llmService);
                // Perform availability check unless user manually disabled AI
                const aiDisabled = localStorage.getItem('aiDisabled') === 'true';
                if (!aiDisabled) {
                    try {
                        const healthy = await this.llmService.healthCheck();
                        if (!healthy) {
                            localStorage.setItem('aiDisabled', 'true');
                            this.showStatus('AI service unreachable – continuing without AI features. You can re-enable later in Settings.', 'error');
                        }
                    } catch (hcErr) {
                        console.warn('LLM health check error:', hcErr);
                        localStorage.setItem('aiDisabled', 'true');
                        this.showStatus('AI service check failed – AI features disabled for this session.', 'error');
                    }
                }
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
        console.log('populateSnoozeOptions called with:', options);
        this.snoozeOptionsSelect.innerHTML = '';
        
        if (!options || options.length === 0) {
            console.warn('No snooze options provided, using default options');
            // Fallback to default options
            options = [
                { label: '15 minutes', value: 15 },
                { label: '1 hour', value: 60 },
                { label: '4 hours', value: 240 },
                { label: '1 day', value: 1440 },
                { label: '3 days', value: 4320 },
                { label: '1 week', value: 10080 },
                { label: 'Custom...', value: 0, isCustom: true }
            ];
        }
        
        options.forEach(option => {
            const optionElement = document.createElement('option');
            optionElement.value = option.value.toString();
            optionElement.textContent = option.label;
            optionElement.dataset.isCustom = option.isCustom?.toString() || 'false';
            this.snoozeOptionsSelect.appendChild(optionElement);
            console.log(`Added snooze option: ${option.label} (${option.value} minutes)`);
        });
        
        console.log(`Total snooze options added: ${this.snoozeOptionsSelect.children.length}`);
    }

    private async saveConfiguration(): Promise<void> {
        try {
            // Get the current configuration to preserve snooze options and other settings
            const currentConfig = await this.configurationService.getConfiguration();
            
            const config: Configuration = {
                ...currentConfig, // Preserve existing settings
                emailCount: parseInt(this.emailCountSelect.value),
                daysBack: parseInt(this.daysBackSelect.value),
                lastAnalysisDate: new Date(),
                enableLlmSummary: this.enableLlmSummaryCheckbox.checked,
                enableLlmSuggestions: this.enableLlmSuggestionsCheckbox.checked,
                llmApiEndpoint: this.llmEndpointInput.value.trim(),
                llmApiKey: this.llmApiKeyInput.value.trim(),
                llmModel: this.llmModelInput.value.trim(),
                llmProvider: this.llmProviderSelect.value as 'azure' | 'dial' | 'openai' || undefined,
                llmDeploymentName: this.llmDeploymentNameInput.value.trim(),
                llmApiVersion: this.llmApiVersionInput.value.trim(),
                showSnoozedEmails: this.showSnoozedEmailsCheckbox.checked,
                showDismissedEmails: this.showDismissedEmailsCheckbox.checked,
                selectedAccounts: Array.from(this.accountFilterSelect.selectedOptions).map(option => option.value)
            };
            await this.configurationService.saveConfiguration(config);
        } catch (error) {
            console.error('Error saving configuration:', error);
        }
    }

    // Enhanced analyzeEmails with progress tracking
    private async analyzeEmails(): Promise<void> {
        try {
            this.debugLog('Starting email analysis');
            this.setLoadingState(true);
            this.hideStatus();
            this.updateProgress(0, 'Initializing analysis...', 'Getting ready to analyze your emails');

            const emailCount = parseInt(this.emailCountSelect.value);
            const daysBack = parseInt(this.daysBackSelect.value);
            const selectedAccounts = Array.from(this.accountFilterSelect.selectedOptions).map(option => option.value);

            this.debugLog('Analysis parameters', { emailCount, daysBack, selectedAccounts });
            
            // Get current configuration and pass it to the EmailAnalysisService
            const config = await this.configurationService.getConfiguration();
            this.emailAnalysisService.setConfiguration(config);
            
            this.updateProgress(20, 'Fetching emails...', `Looking for emails from the last ${daysBack} days`);
            
            // Simulate progress updates during analysis
            const progressInterval = setInterval(() => {
                const currentWidth = parseInt(this.progressFill.style.width) || 20;
                if (currentWidth < 80) {
                    this.updateProgress(currentWidth + 10, 'Analyzing emails...', 'Processing email content and threads');
                }
            }, 500);

            const followupEmails = await this.emailAnalysisService.analyzeEmails(emailCount, daysBack, selectedAccounts);
            
            this.debugLog('Analysis completed', { 
                foundEmails: followupEmails.length,
                emails: followupEmails.map(e => ({ subject: e.subject, sentDate: e.sentDate, recipients: e.recipients }))
            });
            
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
            button.addEventListener('click', (e) => {
                this.handleEmailAction(e).catch(error => {
                    console.error('Error handling email action:', error);
                });
            });
        });

        return emailDiv;
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

    // Requirement: always display most recent emails first (newest -> oldest)
    // Even if analysis service internally sorts by priority, we enforce recency ordering for UI display.
    const recentFirst = [...emails].sort((a, b) => b.sentDate.getTime() - a.sentDate.getTime());

    this.emailListDiv.innerHTML = '';
    recentFirst.forEach(email => {
            const emailElement = this.createEmailElement(email);
            this.emailListDiv.appendChild(emailElement);
        });
        
        this.emailListDiv.style.display = 'block';
    }

    // Handle email action buttons
    private async handleEmailAction(event: Event): Promise<void> {
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
                await this.showSnoozeModal(emailId);
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
        try {
            const email = this.allEmails.find(e => e.id === _emailId);
            if (!email) {
                console.warn('Email not found for reply', _emailId);
                return;
            }
            if (!email.threadMessages || email.threadMessages.length === 0) {
                console.warn('No thread messages available for reply', _emailId);
                return;
            }

            // Assume last element is most recent; if not, sort by sentDate
            const lastMessage = email.threadMessages[email.threadMessages.length - 1];
            const currentUser = Office?.context?.mailbox?.userProfile?.emailAddress?.toLowerCase();

            // Build Reply All recipient lists (approximation; model lacks explicit CC list)
            const unique = new Set<string>();
            const toRecipients: string[] = [];
            const ccRecipients: string[] = [];

            const addRecipient = (addr: string | undefined | null, list: string[]) => {
                if (!addr) return;
                const norm = addr.trim();
                if (!norm) return;
                const lower = norm.toLowerCase();
                if (currentUser && lower === currentUser) return; // exclude self
                if (unique.has(lower)) return;
                unique.add(lower);
                list.push(norm);
            };

            // Put original sender in To if it's not current user
            addRecipient(lastMessage.from, toRecipients);

            // Other original recipients go to CC (excluding current user & sender)
            lastMessage.to.forEach(addr => {
                const lower = addr.toLowerCase();
                if (lastMessage.from && lower === lastMessage.from.toLowerCase()) return;
                addRecipient(addr, ccRecipients);
            });

            // Fallback: if no To recipients (e.g., user was original sender), try using first non-user original recipient
            if (toRecipients.length === 0 && ccRecipients.length > 0) {
                toRecipients.push(ccRecipients.shift()!);
            }

            // Subject handling – prefix Re: if not already present
            let subject = lastMessage.subject || email.subject || '';
            if (!/^re:/i.test(subject)) {
                subject = `Re: ${subject}`;
            }

            // Preserve original HTML formatting when available.
            const rawBody = lastMessage.body || '';
            const looksHtml = /<\w+[^>]*>/.test(rawBody); // crude check for existing HTML tags
            let sanitizedBody = rawBody;
            if (!looksHtml) {
                // Convert plain text newlines to <br> for readability
                sanitizedBody = this.escapeHtml(rawBody).replace(/\n/g, '<br>');
            } else {
                // Minimal sanitization: strip any script/style tags for safety
                sanitizedBody = rawBody.replace(/<script[\s\S]*?<\/script>/gi, '')
                                       .replace(/<style[\s\S]*?<\/style>/gi, '');
            }
            const quotedSeparator = '<br><br>----- Original Message -----<br>';
            const htmlBody = `<br><br>${quotedSeparator}${sanitizedBody}`;

            Office.context.mailbox.displayNewMessageForm({
                toRecipients,
                ccRecipients,
                subject,
                htmlBody,
                attachments: []
            });
        } catch (err) {
            console.error('Error composing reply all form', err);
            this.showStatus('Failed to open reply window', 'error');
        }
    }

    private async forwardEmail(_emailId: string): Promise<void> {
        try {
            const email = this.allEmails.find(e => e.id === _emailId);
            if (!email) {
                console.warn('Email not found for forward', _emailId);
                return;
            }
            if (!email.threadMessages || email.threadMessages.length === 0) {
                console.warn('No thread messages available for forward', _emailId);
                return;
            }

            const lastMessage = email.threadMessages[email.threadMessages.length - 1];
            let subject = lastMessage.subject || email.subject || '';
            if (!/^fw:|^fwd:/i.test(subject)) {
                subject = `FW: ${subject}`; // Outlook commonly uses FW:
            }
            const rawBody = lastMessage.body || '';
            const looksHtml = /<\w+[^>]*>/.test(rawBody);
            let sanitizedBody = rawBody;
            if (!looksHtml) {
                sanitizedBody = this.escapeHtml(rawBody).replace(/\n/g, '<br>');
            } else {
                sanitizedBody = rawBody.replace(/<script[\s\S]*?<\/script>/gi, '')
                                       .replace(/<style[\s\S]*?<\/style>/gi, '');
            }
            const quotedSeparator = '<br><br>----- Forwarded Message -----<br>';
            const htmlBody = `<br><br>${quotedSeparator}${sanitizedBody}`;

            Office.context.mailbox.displayNewMessageForm({
                toRecipients: [], // user will choose
                subject,
                htmlBody,
                attachments: []
            });
        } catch (err) {
            console.error('Error composing forward form', err);
            this.showStatus('Failed to open forward window', 'error');
        }
    }

    // Modal management methods
    private async showSnoozeModal(emailId: string): Promise<void> {
        this.currentEmailForSnooze = emailId;
        
        // Ensure snooze options are populated
        try {
            const config = await this.configurationService.getConfiguration();
            this.populateSnoozeOptions(config.snoozeOptions);
            console.log('Snooze options populated:', config.snoozeOptions);
        } catch (error) {
            console.error('Error loading snooze options:', error);
            // Fallback to default options if configuration fails
            const defaultOptions: SnoozeOption[] = [
                { label: '15 minutes', value: 15 },
                { label: '1 hour', value: 60 },
                { label: '4 hours', value: 240 },
                { label: '1 day', value: 1440 },
                { label: '3 days', value: 4320 },
                { label: '1 week', value: 10080 },
                { label: 'Custom...', value: 0, isCustom: true }
            ];
            this.populateSnoozeOptions(defaultOptions);
        }
        
        this.snoozeModal.style.display = 'block';
    }

    private hideSnoozeModal(): void {
        this.snoozeModal.style.display = 'none';
        this.currentEmailForSnooze = '';
        this.customSnoozeGroup.style.display = 'none';
    }

    private showSettingsModal(): void {
        this.settingsModal.style.display = 'block';
        this.updateAiStatus(); // Update AI status when opening settings
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

            // Update AI status after saving
            this.updateAiStatus();
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
        
        // Load AI disable state
        const aiDisabled = localStorage.getItem('aiDisabled') === 'true';
        this.enableAiFeaturesCheckbox.checked = !aiDisabled;
        
        this.updateAiStatus(); // Check AI status on initialization
        
        // Add diagnostic button event listener
        const diagnosticButton = document.getElementById('diagnosticButton');
        if (diagnosticButton) {
            diagnosticButton.addEventListener('click', () => this.openDiagnosticModal());
        }
        
        // Add close diagnostic modal event listener
        const closeDiagnosticButton = document.querySelector('#diagnosticModal .modal-close');
        if (closeDiagnosticButton) {
            closeDiagnosticButton.addEventListener('click', () => this.closeDiagnosticModal());
        }
        
        // Add event listener for the "Close" button in the modal footer
        const closeDiagnosticFooterButton = document.getElementById('closeDiagnostic');
        if (closeDiagnosticFooterButton) {
            closeDiagnosticFooterButton.addEventListener('click', () => this.closeDiagnosticModal());
        }
        
        // Add event listeners for diagnostic tests
        const testOfficeButton = document.getElementById('testOfficeContext');
        const testMailboxButton = document.getElementById('testMailboxAccess');
        const testAccountsButton = document.getElementById('testAccountDetection');
        const testEmailsButton = document.getElementById('testEmailReading');
        
        if (testOfficeButton) {
            testOfficeButton.addEventListener('click', () => this.testOfficeContext());
        }
        if (testMailboxButton) {
            testMailboxButton.addEventListener('click', () => this.testMailboxAccess());
        }
        if (testAccountsButton) {
            testAccountsButton.addEventListener('click', () => this.testAccountDetection());
        }
        if (testEmailsButton) {
            testEmailsButton.addEventListener('click', () => this.testEmailReading());
        }
    }

    // Diagnostic modal methods
    private openDiagnosticModal(): void {
        const modal = document.getElementById('diagnosticModal');
        if (modal) {
            modal.style.display = 'block';
        }
    }

    private closeDiagnosticModal(): void {
        const modal = document.getElementById('diagnosticModal');
        if (modal) {
            modal.style.display = 'none';
        }
    }

    // Diagnostic test methods
    private async testOfficeContext(): Promise<void> {
        const output = document.getElementById('diagnosticOutput');
        if (!output) return;

        try {
            output.innerHTML = '<strong>Testing Office.js Context...</strong><br>';
            
            // Test basic Office.js availability
            if (typeof Office === 'undefined') {
                output.innerHTML += '❌ Office.js is not available<br>';
                return;
            }
            output.innerHTML += '✅ Office.js is available<br>';

            // Test host type
            if (Office.context && Office.context.host) {
                output.innerHTML += `✅ Host: ${Office.context.host}<br>`;
            } else {
                output.innerHTML += '❌ Office.context.host not available<br>';
            }

            // Test platform
            if (Office.context && Office.context.platform) {
                output.innerHTML += `✅ Platform: ${Office.context.platform}<br>`;
            } else {
                output.innerHTML += '❌ Office.context.platform not available<br>';
            }

            // Test mailbox availability
            if (Office.context && Office.context.mailbox) {
                output.innerHTML += '✅ Mailbox context available<br>';
            } else {
                output.innerHTML += '❌ Mailbox context not available<br>';
            }

            output.innerHTML += '<br><strong>Office Context Test Complete</strong><br><br>';
        } catch (error) {
            output.innerHTML += `❌ Error: ${(error as Error).message}<br>`;
        }
    }

    private async testMailboxAccess(): Promise<void> {
        const output = document.getElementById('diagnosticOutput');
        if (!output) return;

        try {
            output.innerHTML += '<strong>Testing Mailbox Access...</strong><br>';

            if (!Office.context.mailbox) {
                output.innerHTML += '❌ Mailbox not available<br>';
                return;
            }

            // Test user profile
            if (Office.context.mailbox.userProfile) {
                const profile = Office.context.mailbox.userProfile;
                output.innerHTML += `✅ User email: ${profile.emailAddress}<br>`;
                output.innerHTML += `✅ Display name: ${profile.displayName}<br>`;
                output.innerHTML += `✅ Time zone: ${profile.timeZone}<br>`;
            } else {
                output.innerHTML += '❌ User profile not available<br>';
            }

            // Test diagnostics
            if (Office.context.mailbox.diagnostics) {
                const diag = Office.context.mailbox.diagnostics;
                output.innerHTML += `✅ Host name: ${diag.hostName}<br>`;
                output.innerHTML += `✅ Host version: ${diag.hostVersion}<br>`;
                output.innerHTML += `✅ OWA view: ${diag.OWAView}<br>`;
            } else {
                output.innerHTML += '❌ Diagnostics not available<br>';
            }

            output.innerHTML += '<br><strong>Mailbox Access Test Complete</strong><br><br>';
        } catch (error) {
            output.innerHTML += `❌ Error: ${(error as Error).message}<br>`;
        }
    }

    private async testAccountDetection(): Promise<void> {
        const output = document.getElementById('diagnosticOutput');
        if (!output) return;

        try {
            output.innerHTML += '<strong>Testing Account Detection...</strong><br>';

            // Try to get available accounts using our service
            try {
                const accounts = await this.configurationService.getAvailableAccounts();
                if (accounts.length > 0) {
                    output.innerHTML += `✅ Found ${accounts.length} account(s):<br>`;
                    accounts.forEach(account => {
                        output.innerHTML += `&nbsp;&nbsp;• ${account}<br>`;
                    });
                } else {
                    output.innerHTML += '⚠️ No accounts detected<br>';
                }
            } catch (error) {
                output.innerHTML += `❌ Account detection error: ${(error as Error).message}<br>`;
            }

            // Test EWS availability (this is often the issue on macOS)
            if (Office.context.mailbox.makeEwsRequestAsync) {
                output.innerHTML += '✅ EWS (Exchange Web Services) available<br>';
                
                // Test a simple EWS request
                try {
                    const simpleEwsRequest = `<?xml version="1.0" encoding="utf-8"?>
                        <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                                       xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                                       xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                                       xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                          <soap:Header>
                            <t:RequestServerVersion Version="Exchange2013" />
                          </soap:Header>
                          <soap:Body>
                            <m:GetFolder>
                              <m:FolderShape>
                                <t:BaseShape>IdOnly</t:BaseShape>
                              </m:FolderShape>
                              <m:FolderIds>
                                <t:DistinguishedFolderId Id="sentitems" />
                              </m:FolderIds>
                            </m:GetFolder>
                          </soap:Body>
                        </soap:Envelope>`;

                    await new Promise<void>((resolve, reject) => {
                        Office.context.mailbox.makeEwsRequestAsync(simpleEwsRequest, (result) => {
                            if (result.status === Office.AsyncResultStatus.Succeeded) {
                                output.innerHTML += '✅ EWS test request successful<br>';
                                resolve();
                            } else {
                                output.innerHTML += `❌ EWS test failed: ${result.error.message}<br>`;
                                reject(new Error(result.error.message));
                            }
                        });
                    });
                } catch (ewsError) {
                    output.innerHTML += `❌ EWS test request failed: ${(ewsError as Error).message}<br>`;
                }
            } else {
                output.innerHTML += '❌ EWS not available (this is common on macOS)<br>';
            }

            output.innerHTML += '<br><strong>Account Detection Test Complete</strong><br><br>';
        } catch (error) {
            output.innerHTML += `❌ Error: ${(error as Error).message}<br>`;
        }
    }

    private async testEmailReading(): Promise<void> {
        const output = document.getElementById('diagnosticOutput');
        if (!output) return;

        try {
            output.innerHTML += '<strong>Testing Email Reading...</strong><br>';

            // Try to read emails using our email analysis service
            try {
                output.innerHTML += 'Attempting to read 5 emails from last 7 days...<br>';
                const emails = await this.emailAnalysisService.analyzeEmails(5, 7, []);
                
                if (emails.length > 0) {
                    output.innerHTML += `✅ Successfully read ${emails.length} email(s)<br>`;
                    emails.forEach((email, index) => {
                        output.innerHTML += `&nbsp;&nbsp;${index + 1}. ${email.subject} (${email.sentDate.toLocaleDateString()})<br>`;
                    });
                } else {
                    output.innerHTML += '⚠️ No emails found (this could be normal if no sent emails in timeframe)<br>';
                }
            } catch (emailError) {
                output.innerHTML += `❌ Email reading failed: ${(emailError as Error).message}<br>`;
                
                // Provide specific guidance based on error
                const errorMsg = (emailError as Error).message.toLowerCase();
                if (errorMsg.includes('ews') || errorMsg.includes('exchange')) {
                    output.innerHTML += '💡 This appears to be an EWS (Exchange) issue. On macOS, EWS is often limited.<br>';
                    output.innerHTML += '💡 Consider switching to REST API for better macOS compatibility.<br>';
                } else if (errorMsg.includes('permission') || errorMsg.includes('unauthorized')) {
                    output.innerHTML += '💡 This appears to be a permissions issue. Check manifest permissions.<br>';
                } else if (errorMsg.includes('network') || errorMsg.includes('timeout')) {
                    output.innerHTML += '💡 This appears to be a network connectivity issue.<br>';
                }
            }

            output.innerHTML += '<br><strong>Email Reading Test Complete</strong><br><br>';
        } catch (error) {
            output.innerHTML += `❌ Error: ${(error as Error).message}<br>`;
        }
    }

    // AI Service Management Methods
    private updateAiStatus(): void {
        const aiDisabled = localStorage.getItem('aiDisabled') === 'true';
        
        if (aiDisabled) {
            this.setAiStatus('warning', '⚠️ AI features manually disabled');
            this.enableAiFeaturesCheckbox.checked = false;
            return;
        }

        if (!this.llmService || !this.llmEndpointInput.value.trim() || !this.llmApiKeyInput.value.trim()) {
            this.setAiStatus('warning', '⚠️ AI features disabled - No API configuration');
            return;
        }

        // Check if circuit breaker is open for llm-api
        const circuitStates = this.retryService.getCircuitBreakerStates();
        if (circuitStates['llm-api'] === 'OPEN') {
            this.setAiStatus('error', '❌ AI service temporarily unavailable - Too many failures detected');
            return;
        }

        this.setAiStatus('success', '✅ AI service configured and ready');
    }

    private toggleAiFeatures(): void {
        const isEnabled = this.enableAiFeaturesCheckbox.checked;
        localStorage.setItem('aiDisabled', (!isEnabled).toString());
        
        if (!isEnabled) {
            this.setAiStatus('warning', '⚠️ AI features manually disabled');
            this.showStatus('AI features disabled - errors will stop appearing', 'success');
        } else {
            this.updateAiStatus();
            this.showStatus('AI features re-enabled', 'success');
        }
    }

    private disableAiFeatures(): void {
        localStorage.setItem('aiDisabled', 'true');
        this.enableAiFeaturesCheckbox.checked = false;
        this.setAiStatus('warning', '⚠️ AI features manually disabled');
        this.showStatus('AI features disabled - no more AI-related errors will appear', 'success');
    }

    private handleProviderChange(): void {
        const provider = this.llmProviderSelect.value;
        const isAzure = provider === 'azure' || 
                       (provider === '' && this.llmEndpointInput.value.includes('openai.azure.com'));
        
        this.azureSpecificOptions.style.display = isAzure ? 'block' : 'none';
        
        // Set default values based on provider
        if (provider === 'dial' && !this.llmEndpointInput.value) {
            this.llmEndpointInput.value = 'https://ai-proxy.lab.epam.com/openai/chat/completions';
            this.llmModelInput.value = 'gpt-35-turbo';
        } else if (provider === 'azure' && !this.llmEndpointInput.value) {
            this.llmEndpointInput.placeholder = 'https://your-resource.openai.azure.com';
            this.llmModelInput.value = 'gpt-35-turbo';
            this.llmApiVersionInput.value = '2023-12-01-preview';
        } else if (provider === 'openai' && !this.llmEndpointInput.value) {
            this.llmEndpointInput.value = 'https://api.openai.com/v1/chat/completions';
            this.llmModelInput.value = 'gpt-3.5-turbo';
        }
    }

    private setAiStatus(type: 'success' | 'warning' | 'error', message: string): void {
        this.aiStatusDiv.style.display = 'block';
        this.aiStatusDiv.className = `ai-status ${type}`;
        this.aiStatusText.textContent = message;
    }

    private async testAiConnection(): Promise<void> {
        if (!this.llmEndpointInput.value.trim() || !this.llmApiKeyInput.value.trim()) {
            this.setAiStatus('warning', '⚠️ Please enter both API endpoint and API key');
            return;
        }

        this.setAiStatus('warning', '🔄 Testing AI connection...');
        this.testAiConnectionButton.disabled = true;
        this.testAiConnectionButton.textContent = 'Testing...';

        try {
            // Create a temporary configuration for testing
            const testConfig: Configuration = {
                emailCount: 10,
                daysBack: 7,
                lastAnalysisDate: new Date(),
                enableLlmSummary: true,
                enableLlmSuggestions: true,
                llmApiEndpoint: this.llmEndpointInput.value.trim(),
                llmApiKey: this.llmApiKeyInput.value.trim(),
                llmModel: this.llmModelInput.value.trim() || 'gpt-35-turbo',
                llmProvider: this.llmProviderSelect.value as 'azure' | 'dial' | 'openai' || undefined,
                llmDeploymentName: this.llmDeploymentNameInput.value.trim(),
                llmApiVersion: this.llmApiVersionInput.value.trim() || '2023-12-01-preview',
                showSnoozedEmails: false,
                showDismissedEmails: false,
                selectedAccounts: [],
                snoozeOptions: []
            };

            // Create a temporary LLM service for testing
            const testLlmService = new LlmService(testConfig, this.retryService);
            
            // Test with a simple prompt
            const testPrompt = "Test connection. Please respond with 'Connection successful'.";
            const response = await testLlmService.generateFollowupSuggestions(testPrompt);
            
            if (response && response.length > 0) {
                this.setAiStatus('success', '✅ AI connection test successful!');
            } else {
                this.setAiStatus('error', '❌ AI service responded but with empty result');
            }
        } catch (error) {
            const errorMessage = (error as Error).message;
            if (errorMessage.includes('502')) {
                this.setAiStatus('error', '❌ Connection failed: AI service not available (502 error)');
            } else if (errorMessage.includes('403') || errorMessage.includes('401')) {
                this.setAiStatus('error', '❌ Connection failed: Invalid API key or unauthorized');
            } else if (errorMessage.includes('429')) {
                this.setAiStatus('error', '❌ Connection failed: Rate limit exceeded');
            } else {
                this.setAiStatus('error', `❌ Connection failed: ${errorMessage}`);
            }
        } finally {
            this.testAiConnectionButton.disabled = false;
            this.testAiConnectionButton.textContent = 'Test AI Connection';
        }
    }
}

// Initialize the task pane when Office is ready (skip in test/non-DOM environments)
if (typeof Office !== 'undefined' && typeof Office.onReady === 'function') {
    Office.onReady((info) => {
        try {
            // Only auto-initialize if running in Outlook host and expected root element exists
            if (info.host === Office.HostType.Outlook && typeof document !== 'undefined' && document.getElementById('analyzeButton')) {
                const taskpaneManager = new TaskpaneManager();
                taskpaneManager.initialize().catch(console.error);
            }
        } catch (e) {
            // Swallow errors in headless/unit test environments
            console.warn('Taskpane auto-initialization skipped:', e);
        }
    });
}
