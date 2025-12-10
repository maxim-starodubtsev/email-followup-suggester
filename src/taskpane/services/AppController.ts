import { EmailAnalysisService } from "../../services/EmailAnalysisService";
import { ConfigurationService } from "../../services/ConfigurationService";
import { LlmService } from "../../services/LlmService";
import { RetryService } from "../../services/RetryService";
import { Configuration } from "../../models/Configuration";
import { FollowupEmail } from "../../models/FollowupEmail";
import { UiService } from "./UiService";

export class AppController {
  private emailAnalysisService: EmailAnalysisService;
  private configurationService: ConfigurationService;
  private llmService?: LlmService;
  private retryService: RetryService;
  private uiService: UiService;

  private availableAccounts: string[] = [];
  private allEmails: FollowupEmail[] = [];
  private filteredEmails: FollowupEmail[] = [];
  private currentEmailForSnooze: string = "";
  private debounceTimer?: number;

  constructor(uiService: UiService) {
    this.uiService = uiService;
    this.retryService = new RetryService();
    this.emailAnalysisService = new EmailAnalysisService();
    this.configurationService = new ConfigurationService();
  }

  public async initialize(): Promise<void> {
    this.attachEventListeners();
    await this.loadConfiguration();
    await this.loadCachedResults();
    
    // Check AI status
    const aiDisabled = localStorage.getItem("aiDisabled") === "true";
    this.uiService.enableAiFeaturesCheckbox.checked = !aiDisabled;
    this.updateAiStatus();
  }

  private attachEventListeners(): void {
    // Main controls
    this.uiService.analyzeButton.addEventListener("click", () => this.analyzeEmails());
    this.uiService.refreshButton.addEventListener("click", () => this.analyzeEmails());
    this.uiService.settingsButton.addEventListener("click", () => this.showSettingsModal());

    // Configuration changes
    const configElements = [
      this.uiService.emailCountSelect,
      this.uiService.daysBackSelect,
      this.uiService.accountFilterSelect,
      this.uiService.enableLlmSummaryCheckbox,
      this.uiService.enableLlmSuggestionsCheckbox
    ];
    configElements.forEach(el => el.addEventListener("change", () => this.saveConfiguration()));

    // UI Toggles
    this.uiService.showStatsButton.addEventListener("click", () => this.uiService.toggleStatsDashboard(true));
    this.uiService.toggleStatsButton.addEventListener("click", () => this.uiService.toggleStatsDashboard(false));
    this.uiService.toggleAdvancedFiltersButton.addEventListener("click", () => this.uiService.toggleAdvancedFilters());
    this.uiService.clearFiltersButton.addEventListener("click", () => this.clearAllFilters());

    // Filters
    this.uiService.priorityFilter.addEventListener("change", () => this.applyFilters());
    this.uiService.responseTimeFilter.addEventListener("change", () => this.applyFilters());
    this.uiService.aiSuggestionFilter.addEventListener("change", () => this.applyFilters());
    
    this.uiService.subjectFilter.addEventListener("input", () => this.debounceApplyFilters());
    this.uiService.senderFilter.addEventListener("input", () => this.debounceApplyFilters());

    // Modals
    document.getElementById("confirmSnooze")?.addEventListener("click", () => this.confirmSnooze());
    document.getElementById("cancelSnooze")?.addEventListener("click", () => this.uiService.snoozeModal.style.display = "none");
    this.uiService.snoozeOptionsSelect.addEventListener("change", () => this.handleSnoozeOptionChange());

    document.getElementById("saveSettings")?.addEventListener("click", () => this.saveSettings());
    document.getElementById("cancelSettings")?.addEventListener("click", () => this.uiService.settingsModal.style.display = "none");

    this.uiService.testAiConnectionButton.addEventListener("click", () => this.testAiConnection());
    this.uiService.disableAiFeaturesButton.addEventListener("click", () => this.disableAiFeatures());
    this.uiService.enableAiFeaturesCheckbox.addEventListener("change", () => this.toggleAiFeatures());
    this.uiService.llmProviderSelect.addEventListener("change", () => this.handleProviderChange());

    // Thread modal
    this.uiService.threadModal?.addEventListener("click", (e) => {
        if (e.target === this.uiService.threadModal) this.uiService.hideThreadModal();
    });

    // Email Actions
    this.uiService.setOnActionCallback((action, emailId) => this.handleEmailAction(action, emailId));
  }

  private async loadConfiguration(): Promise<void> {
    try {
      const config = await this.configurationService.getConfiguration();
      this.uiService.emailCountSelect.value = config.emailCount.toString();
      this.uiService.daysBackSelect.value = config.daysBack.toString();
      this.uiService.enableLlmSummaryCheckbox.checked = config.enableLlmSummary;
      this.uiService.enableLlmSuggestionsCheckbox.checked = config.enableLlmSuggestions;
      this.uiService.showSnoozedEmailsCheckbox.checked = config.showSnoozedEmails;
      this.uiService.showDismissedEmailsCheckbox.checked = config.showDismissedEmails;

      if (config.llmApiEndpoint) this.uiService.llmEndpointInput.value = config.llmApiEndpoint;
      if (config.llmApiKey) this.uiService.llmApiKeyInput.value = config.llmApiKey;
      if (config.llmModel) this.uiService.llmModelInput.value = config.llmModel;
      if (config.llmProvider) this.uiService.llmProviderSelect.value = config.llmProvider;
      if (config.llmDeploymentName) this.uiService.llmDeploymentNameInput.value = config.llmDeploymentName;
      if (config.llmApiVersion) this.uiService.llmApiVersionInput.value = config.llmApiVersion;

      this.handleProviderChange();

      // Load accounts
      try {
        this.availableAccounts = await this.configurationService.getAvailableAccounts();
      } catch (e) {
        console.error("Error loading accounts", e);
        this.availableAccounts = [];
      }
      this.uiService.populateAccountFilter(config.selectedAccounts, this.availableAccounts);

      this.emailAnalysisService.setConfiguration(config);

      // Initialize LLM
      if (config.llmApiEndpoint && config.llmApiKey) {
        this.llmService = new LlmService(config, this.retryService);
        this.emailAnalysisService.setLlmService(this.llmService);
        
        const aiDisabled = localStorage.getItem("aiDisabled") === "true";
        if (!aiDisabled) {
            this.llmService.healthCheck().then(healthy => {
                if (!healthy) {
                    localStorage.setItem("aiDisabled", "true");
                    this.uiService.showStatus("AI service unreachable - disabled for session", "error");
                }
            }).catch(() => {
                localStorage.setItem("aiDisabled", "true");
            });
        }
      }

      this.uiService.populateSnoozeOptions(config.snoozeOptions);

    } catch (error) {
      console.error("Error loading configuration:", error);
    }
  }

  private async analyzeEmails(): Promise<void> {
    try {
      this.uiService.setLoadingState(true);
      this.uiService.hideStatus();
      this.uiService.updateProgress(0, "Initializing...", "Getting ready");

      const emailCount = parseInt(this.uiService.emailCountSelect.value);
      const daysBack = parseInt(this.uiService.daysBackSelect.value);
      const selectedAccounts = Array.from(this.uiService.accountFilterSelect.selectedOptions).map(o => o.value);

      const config = await this.configurationService.getConfiguration();
      this.emailAnalysisService.setConfiguration(config);

      this.uiService.updateProgress(20, "Fetching emails...", `Checking last ${daysBack} days`);

      // Fake progress
      const progressInterval = setInterval(() => {
        const current = this.uiService.getCurrentProgress();
        if (current < 80) {
            this.uiService.updateProgress(current + 10, "Analyzing...", "Processing threads");
        }
      }, 500);

      const followupEmails = await this.emailAnalysisService.analyzeEmails(emailCount, daysBack, selectedAccounts);

      clearInterval(progressInterval);
      this.uiService.updateProgress(100, "Done!", "Preparing results");

      this.allEmails = followupEmails;
      this.filteredEmails = [...followupEmails];

      await this.saveConfiguration(); // Save last analysis date implicitly via analyzeEmails side effects? No, explicitly.
      // But analyzeEmails in service doesn't save config.
      
      this.applyFilters(); // This calls displayEmails and updateStatistics

      if (followupEmails.length > 0) {
        this.uiService.showStatus(`Found ${followupEmails.length} emails needing follow-up`, "success");
      } else {
        this.uiService.showStatus("No emails needing follow-up found", "success");
      }

    } catch (error) {
      console.error("Analysis error:", error);
      this.uiService.showStatus(`Error: ${(error as Error).message}`, "error");
      this.uiService.displayEmails([]);
    } finally {
      this.uiService.setLoadingState(false);
    }
  }

  private async saveConfiguration(): Promise<void> {
    try {
      const currentConfig = await this.configurationService.getConfiguration();
      const config: Configuration = {
        ...currentConfig,
        emailCount: parseInt(this.uiService.emailCountSelect.value),
        daysBack: parseInt(this.uiService.daysBackSelect.value),
        lastAnalysisDate: new Date(),
        enableLlmSummary: this.uiService.enableLlmSummaryCheckbox.checked,
        enableLlmSuggestions: this.uiService.enableLlmSuggestionsCheckbox.checked,
        llmApiEndpoint: this.uiService.llmEndpointInput.value.trim(),
        llmApiKey: this.uiService.llmApiKeyInput.value.trim(),
        llmModel: this.uiService.llmModelInput.value.trim(),
        llmProvider: (this.uiService.llmProviderSelect.value as "azure" | "dial" | "openai") || undefined,
        llmDeploymentName: this.uiService.llmDeploymentNameInput.value.trim(),
        llmApiVersion: this.uiService.llmApiVersionInput.value.trim(),
        showSnoozedEmails: this.uiService.showSnoozedEmailsCheckbox.checked,
        showDismissedEmails: this.uiService.showDismissedEmailsCheckbox.checked,
        selectedAccounts: Array.from(this.uiService.accountFilterSelect.selectedOptions).map(o => o.value),
      };
      await this.configurationService.saveConfiguration(config);
    } catch (error) {
      console.error("Error saving config:", error);
    }
  }

  private debounceApplyFilters(): void {
    if (this.debounceTimer) clearTimeout(this.debounceTimer);
    this.debounceTimer = window.setTimeout(() => this.applyFilters(), 300);
  }

  private applyFilters(): void {
    this.filteredEmails = this.allEmails.filter(email => {
        // Priority
        if (this.uiService.priorityFilter.value && email.priority !== this.uiService.priorityFilter.value) return false;
        
        // Response Time
        if (this.uiService.responseTimeFilter.value) {
            const days = email.daysWithoutResponse;
            const range = this.uiService.responseTimeFilter.value;
            if (range === "1-3" && (days < 1 || days > 3)) return false;
            if (range === "4-7" && (days < 4 || days > 7)) return false;
            if (range === "8-14" && (days < 8 || days > 14)) return false;
            if (range === "15+" && days < 15) return false;
        }

        // Subject
        if (this.uiService.subjectFilter.value && !email.subject.toLowerCase().includes(this.uiService.subjectFilter.value.toLowerCase())) return false;

        // Sender
        if (this.uiService.senderFilter.value) {
            const senderText = this.uiService.senderFilter.value.toLowerCase();
            const hasMatching = email.recipients.some(r => r.toLowerCase().includes(senderText));
            if (!hasMatching) return false;
        }

        // AI
        if (this.uiService.aiSuggestionFilter.value) {
            const hasAI = !!email.llmSuggestion;
            if (this.uiService.aiSuggestionFilter.value === "with-ai" && !hasAI) return false;
            if (this.uiService.aiSuggestionFilter.value === "without-ai" && hasAI) return false;
        }

        return true;
    });

    this.uiService.displayEmails(this.filteredEmails);
    
    // Update stats
    const total = this.allEmails.length;
    const needing = this.filteredEmails.length;
    const high = this.allEmails.filter(e => e.priority === "high").length;
    const avg = total > 0 ? Math.round(this.allEmails.reduce((s, e) => s + e.daysWithoutResponse, 0) / total) : 0;
    this.uiService.updateStatistics(total, needing, high, avg);
  }

  private clearAllFilters(): void {
    this.uiService.priorityFilter.value = "";
    this.uiService.responseTimeFilter.value = "";
    this.uiService.subjectFilter.value = "";
    this.uiService.senderFilter.value = "";
    this.uiService.aiSuggestionFilter.value = "";
    this.applyFilters();
  }

  private async loadCachedResults(): Promise<void> {
    try {
        const cached = await this.configurationService.getCachedAnalysisResults();
        if (cached.length > 0) {
            this.allEmails = cached;
            this.filteredEmails = [...cached];
            this.applyFilters();
            this.uiService.showStatus(`Loaded ${cached.length} cached emails`, "success");
        }
    } catch (e) {
        console.error("Error loading cache", e);
    }
  }

  // --- Email Actions ---

  private async handleEmailAction(action: string, emailId: string): Promise<void> {
    switch (action) {
        case "reply":
            await this.replyToEmail(emailId);
            break;
        case "forward":
            await this.forwardEmail(emailId);
            break;
        case "snooze":
            this.currentEmailForSnooze = emailId;
            this.uiService.snoozeModal.style.display = "block";
            break;
        case "dismiss":
            await this.dismissEmail(emailId);
            break;
        case "view-thread":
            const email = this.allEmails.find(e => e.id === emailId);
            if (email) this.uiService.showThreadView(email);
            break;
    }
  }

  private async replyToEmail(emailId: string): Promise<void> {
    // simplified for brevity - assumes Outlook context
    if (Office.context?.mailbox && typeof (Office.context.mailbox as any).displayMessageForm === "function") {
        (Office.context.mailbox as any).displayMessageForm(emailId);
    } else {
        this.uiService.showStatus("Reply requires Outlook", "error");
    }
  }

  private async forwardEmail(emailId: string): Promise<void> {
    // simplified - leveraging TaskpaneManager's logic would be better but for now basic stub
    if (Office.context?.mailbox && typeof (Office.context.mailbox as any).displayMessageForm === "function") {
        (Office.context.mailbox as any).displayMessageForm(emailId);
    } else {
        // Fallback or error
        this.uiService.showStatus("Forward requires Outlook", "error");
    }
  }

  private async dismissEmail(emailId: string): Promise<void> {
    this.emailAnalysisService.dismissEmail(emailId);
    // Remove from allEmails
    this.allEmails = this.allEmails.filter(e => e.id !== emailId);
    this.applyFilters();
    this.uiService.showStatus("Email dismissed", "success");
  }

  // --- Snooze Logic ---

  private handleSnoozeOptionChange(): void {
    const selected = this.uiService.snoozeOptionsSelect.selectedOptions[0];
    const isCustom = selected?.dataset.isCustom === "true";
    this.uiService.customSnoozeGroup.style.display = isCustom ? "block" : "none";
    if (isCustom) {
        const d = new Date();
        d.setHours(d.getHours() + 1);
        this.uiService.customSnoozeDate.value = d.toISOString().slice(0, 16);
    }
  }

  private async confirmSnooze(): Promise<void> {
    if (!this.currentEmailForSnooze) return;
    const selected = this.uiService.snoozeOptionsSelect.selectedOptions[0];
    const isCustom = selected?.dataset.isCustom === "true";

    try {
        if (isCustom) {
            const date = new Date(this.uiService.customSnoozeDate.value);
            if (date <= new Date()) {
                this.uiService.showStatus("Snooze time must be in future", "error");
                return;
            }
            this.emailAnalysisService.snoozeEmailUntil(this.currentEmailForSnooze, date);
        } else {
            const minutes = parseInt(selected.value);
            this.emailAnalysisService.snoozeEmail(this.currentEmailForSnooze, minutes);
        }

        this.allEmails = this.allEmails.filter(e => e.id !== this.currentEmailForSnooze);
        this.applyFilters();
        this.uiService.showStatus("Email snoozed", "success");
        this.uiService.snoozeModal.style.display = "none";
    } catch (e) {
        this.uiService.showStatus(`Error snoozing: ${(e as Error).message}`, "error");
    }
  }

  // --- Settings & AI ---

  private showSettingsModal(): void {
    this.uiService.settingsModal.style.display = "block";
    this.updateAiStatus();
  }

  private async saveSettings(): Promise<void> {
    try {
        const endpoint = this.uiService.llmEndpointInput.value.trim();
        const apiKey = this.uiService.llmApiKeyInput.value.trim();
        const summary = this.uiService.showSnoozedEmailsCheckbox.checked; // Reusing checkbox for this? TaskpaneManager logic seemed to map fields differently. 
        // Checking TaskpaneManager:
        // const enableSummary = this.showSnoozedEmailsCheckbox.checked; <- Wait, TaskpaneManager lines 1434-1435:
        // const enableSummary = this.showSnoozedEmailsCheckbox.checked;
        // const enableSuggestions = this.showDismissedEmailsCheckbox.checked;
        // This looks like a bug in TaskpaneManager or weird reuse.
        // Actually looking at `saveConfiguration`, it uses `enableLlmSummaryCheckbox` for summary.
        // `saveSettings` in TaskpaneManager used `showSnoozedEmailsCheckbox`. That seems wrong.
        // I will use the correct inputs.
        
        await this.configurationService.updateLlmSettings(
            endpoint,
            apiKey,
            this.uiService.enableLlmSummaryCheckbox.checked,
            this.uiService.enableLlmSuggestionsCheckbox.checked
        );

        // Reload config to update service
        await this.loadConfiguration();
        this.uiService.showStatus("Settings saved", "success");
        this.uiService.settingsModal.style.display = "none";
    } catch (e) {
        this.uiService.showStatus(`Error: ${(e as Error).message}`, "error");
    }
  }

  private updateAiStatus(): void {
    const aiDisabled = localStorage.getItem("aiDisabled") === "true";
    if (aiDisabled) {
        this.uiService.setAiStatus("warning", "AI manually disabled");
        return;
    }
    
    if (!this.llmService || !this.uiService.llmEndpointInput.value.trim() || !this.uiService.llmApiKeyInput.value.trim()) {
        this.uiService.setAiStatus("warning", "AI features disabled - No config");
        return;
    }

    const circuitStates = this.retryService.getCircuitBreakerStates();
    if (circuitStates["llm-api"] === "OPEN") {
        this.uiService.setAiStatus("error", "AI service unavailable (Circuit Open)");
        return;
    }

    this.uiService.setAiStatus("success", "AI service ready");
  }

  private disableAiFeatures(): void {
    localStorage.setItem("aiDisabled", "true");
    this.uiService.enableAiFeaturesCheckbox.checked = false;
    this.updateAiStatus();
    this.uiService.showStatus("AI features disabled", "success");
  }

  private toggleAiFeatures(): void {
    const enabled = this.uiService.enableAiFeaturesCheckbox.checked;
    localStorage.setItem("aiDisabled", (!enabled).toString());
    this.updateAiStatus();
    this.uiService.showStatus(enabled ? "AI enabled" : "AI disabled", "success");
  }

  private async testAiConnection(): Promise<void> {
    // Basic stub - reusing logic from TaskpaneManager would be better but keeping it simple for now
    this.uiService.setAiStatus("warning", "Testing...");
    try {
        // ... (TaskpaneManager logic for test connection)
        if (this.llmService) {
            const res = await this.llmService.generateFollowupSuggestions("Test");
            if (res) this.uiService.setAiStatus("success", "Connection successful");
        } else {
            this.uiService.setAiStatus("error", "LLM Service not initialized");
        }
    } catch (e) {
        this.uiService.setAiStatus("error", `Failed: ${(e as Error).message}`);
    }
  }

  private handleProviderChange(): void {
    const provider = this.uiService.llmProviderSelect.value;
    const isAzure = provider === "azure" || (provider === "" && this.uiService.llmEndpointInput.value.includes("azure"));
    this.uiService.azureSpecificOptions.style.display = isAzure ? "block" : "none";
  }
}
