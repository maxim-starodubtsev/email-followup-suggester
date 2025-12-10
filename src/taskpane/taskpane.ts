import { EmailAnalysisService } from "../services/EmailAnalysisService";
import { ConfigurationService } from "../services/ConfigurationService";
import { LlmService } from "../services/LlmService";
import { RetryService } from "../services/RetryService";
import { Configuration } from "../models/Configuration";
import { FollowupEmail } from "../models/FollowupEmail";

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

  private currentEmailForSnooze: string = "";
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
      console.log(
        "%c[Followup Suggester] Debug logging enabled",
        "color: #2196F3; font-weight: bold;",
      );
      console.log(
        '%cTo see detailed debug output, click "Analyze Emails" and watch this console',
        "color: #666;",
      );
    }
  }

  private debugLog(message: string, data?: any): void {
    if (this.debugEnabled) {
      console.log(
        `%c[Followup Suggester] ${message}`,
        "color: #2196F3;",
        data || "",
      );
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
        innerHTML: "",
        value: "",
        disabled: false,
        checked: false,
        textContent: "",
        addEventListener: () => {},
        removeEventListener: () => {},
        appendChild: () => {},
        querySelector: () => null,
        querySelectorAll: () => [],
        classList: {
          add: () => {},
          remove: () => {},
          toggle: () => {},
          contains: () => false,
        },
      } as any;
    };
    // Main controls
    this.analyzeButton = safeElement("analyzeButton") as HTMLButtonElement;
    this.refreshButton = safeElement("refreshButton") as HTMLButtonElement;
    this.settingsButton = safeElement("settingsButton") as HTMLButtonElement;
    this.emailCountSelect = safeElement("emailCount") as HTMLSelectElement;
    this.daysBackSelect = safeElement("daysBack") as HTMLSelectElement;
    this.accountFilterSelect = safeElement(
      "accountFilter",
    ) as HTMLSelectElement;
    this.enableLlmSummaryCheckbox = safeElement(
      "enableLlmSummary",
    ) as HTMLInputElement;
    this.enableLlmSuggestionsCheckbox = safeElement(
      "enableLlmSuggestions",
    ) as HTMLInputElement;

    // Display elements
    this.statusDiv = safeElement("status") as HTMLDivElement;
    this.loadingDiv = safeElement("loadingMessage") as HTMLDivElement;
    this.emptyStateDiv = safeElement("emptyState") as HTMLDivElement;
    this.emailListDiv = safeElement("emailList") as HTMLDivElement;

    // Modal elements
    this.snoozeModal = safeElement("snoozeModal") as HTMLDivElement;
    this.settingsModal = safeElement("settingsModal") as HTMLDivElement;
    this.snoozeOptionsSelect = safeElement(
      "snoozeOptions",
    ) as HTMLSelectElement;
    this.customSnoozeGroup = safeElement("customSnoozeGroup") as HTMLDivElement;
    this.customSnoozeDate = safeElement("customSnoozeDate") as HTMLInputElement;
    this.llmEndpointInput = safeElement("llmEndpoint") as HTMLInputElement;
    this.llmApiKeyInput = safeElement("llmApiKey") as HTMLInputElement;
    this.showSnoozedEmailsCheckbox = safeElement(
      "showSnoozedEmails",
    ) as HTMLInputElement;
    this.showDismissedEmailsCheckbox = safeElement(
      "showDismissedEmails",
    ) as HTMLInputElement;
    this.aiStatusDiv = safeElement("aiStatus") as HTMLDivElement;
    this.aiStatusText = safeElement("aiStatusText") as HTMLSpanElement;
    this.testAiConnectionButton = safeElement(
      "testAiConnection",
    ) as HTMLButtonElement;
    this.disableAiFeaturesButton = safeElement(
      "disableAiFeatures",
    ) as HTMLButtonElement;
    this.enableAiFeaturesCheckbox = safeElement(
      "enableAiFeatures",
    ) as HTMLInputElement;
    this.llmProviderSelect = safeElement("llmProvider") as HTMLSelectElement;
    this.llmModelInput = safeElement("llmModel") as HTMLInputElement;
    this.llmDeploymentNameInput = safeElement(
      "llmDeploymentName",
    ) as HTMLInputElement;
    this.llmApiVersionInput = safeElement("llmApiVersion") as HTMLInputElement;
    this.azureSpecificOptions = safeElement(
      "azureSpecificOptions",
    ) as HTMLDivElement;

    // New UI elements
    this.statsDashboard = safeElement("statsDashboard") as HTMLDivElement;
    this.showStatsButton = safeElement("showStatsButton") as HTMLButtonElement;
    this.toggleStatsButton = safeElement("toggleStats") as HTMLButtonElement;
    this.advancedFilters = safeElement("advancedFilters") as HTMLDivElement;
    this.toggleAdvancedFiltersButton = safeElement(
      "toggleAdvancedFilters",
    ) as HTMLButtonElement;
    this.threadModal = safeElement("threadModal") as HTMLDivElement;
    this.threadSubject = safeElement("threadSubject") as HTMLHeadingElement;
    this.threadBody = safeElement("threadBody") as HTMLDivElement;

    // Statistics elements
    this.totalEmailsAnalyzedSpan = safeElement(
      "totalEmailsAnalyzed",
    ) as HTMLSpanElement;
    this.needingFollowupSpan = safeElement(
      "needingFollowup",
    ) as HTMLSpanElement;
    this.highPriorityCountSpan = safeElement(
      "highPriorityCount",
    ) as HTMLSpanElement;
    this.avgResponseTimeSpan = safeElement(
      "avgResponseTime",
    ) as HTMLSpanElement;

    // Filter elements
    this.priorityFilter = safeElement("priorityFilter") as HTMLSelectElement;
    this.responseTimeFilter = safeElement(
      "responseTimeFilter",
    ) as HTMLSelectElement;
    this.subjectFilter = safeElement("subjectFilter") as HTMLInputElement;
    this.senderFilter = safeElement("senderFilter") as HTMLInputElement;
    this.aiSuggestionFilter = safeElement(
      "aiSuggestionFilter",
    ) as HTMLSelectElement;
    this.clearFiltersButton = safeElement("clearFilters") as HTMLButtonElement;

    // Progress elements
    this.loadingStep = safeElement("loadingStep") as HTMLSpanElement;
    this.loadingDetail = safeElement("loadingDetail") as HTMLDivElement;
    this.progressFill = safeElement("progressFill") as HTMLDivElement;
  }

  private attachEventListeners(): void {
    // Main controls
    this.analyzeButton.addEventListener("click", () => this.analyzeEmails());
    this.refreshButton.addEventListener("click", () => this.refreshEmails());
    this.settingsButton.addEventListener("click", () =>
      this.showSettingsModal(),
    );

    // Save configuration when changed
    this.emailCountSelect.addEventListener("change", () =>
      this.saveConfiguration(),
    );
    this.daysBackSelect.addEventListener("change", () =>
      this.saveConfiguration(),
    );
    this.accountFilterSelect.addEventListener("change", () =>
      this.saveConfiguration(),
    );
    this.enableLlmSummaryCheckbox.addEventListener("change", () =>
      this.saveConfiguration(),
    );
    this.enableLlmSuggestionsCheckbox.addEventListener("change", () =>
      this.saveConfiguration(),
    );

    // New UI event listeners
    this.showStatsButton.addEventListener("click", () =>
      this.toggleStatsDashboard(true),
    );
    this.toggleStatsButton.addEventListener("click", () =>
      this.toggleStatsDashboard(false),
    );
    this.toggleAdvancedFiltersButton.addEventListener("click", () =>
      this.toggleAdvancedFilters(),
    );
    this.clearFiltersButton.addEventListener("click", () =>
      this.clearAllFilters(),
    );

    // Filter event listeners
    this.priorityFilter.addEventListener("change", () => this.applyFilters());
    this.responseTimeFilter.addEventListener("change", () =>
      this.applyFilters(),
    );
    this.subjectFilter.addEventListener("input", () =>
      this.debounceApplyFilters(),
    );
    this.senderFilter.addEventListener("input", () =>
      this.debounceApplyFilters(),
    );
    this.aiSuggestionFilter.addEventListener("change", () =>
      this.applyFilters(),
    );

    // Modal controls
    this.attachModalEventListeners();

    // Thread modal event listeners
    this.threadModal.addEventListener("click", (e) => {
      if (e.target === this.threadModal) this.hideThreadModal();
    });
  }

  private attachModalEventListeners(): void {
    // Snooze modal
    document
      .getElementById("confirmSnooze")
      ?.addEventListener("click", () => this.confirmSnooze());
    document
      .getElementById("cancelSnooze")
      ?.addEventListener("click", () => this.hideSnoozeModal());
    this.snoozeOptionsSelect.addEventListener("change", () =>
      this.handleSnoozeOptionChange(),
    );

    // Settings modal
    document
      .getElementById("saveSettings")
      ?.addEventListener("click", () => this.saveSettings());
    document
      .getElementById("cancelSettings")
      ?.addEventListener("click", () => this.hideSettingsModal());
    this.testAiConnectionButton.addEventListener("click", () =>
      this.testAiConnection(),
    );
    this.disableAiFeaturesButton.addEventListener("click", () =>
      this.disableAiFeatures(),
    );
    this.enableAiFeaturesCheckbox.addEventListener("change", () =>
      this.toggleAiFeatures(),
    );
    this.llmProviderSelect.addEventListener("change", () =>
      this.handleProviderChange(),
    );

    // Close modals when clicking outside or on close button
    document.addEventListener("click", (e) => {
      const target = e.target as Element;
      if (target.classList.contains("modal-close")) {
        const modal = target.closest(".modal") as HTMLDivElement;
        if (modal) modal.style.display = "none";
      }
    });

    window.addEventListener("click", (e) => {
      if (e.target === this.snoozeModal) this.hideSnoozeModal();
      if (e.target === this.settingsModal) this.hideSettingsModal();
      // Handle diagnostic modal close on outside click
      const diagnosticModal = document.getElementById("diagnosticModal");
      if (diagnosticModal && e.target === diagnosticModal) {
        diagnosticModal.style.display = "none";
      }
    });

    // Handle escape key for modal closing
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape") {
        const diagnosticModal = document.getElementById("diagnosticModal");
        if (diagnosticModal && diagnosticModal.style.display === "block") {
          diagnosticModal.style.display = "none";
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

      if (config.llmApiEndpoint)
        this.llmEndpointInput.value = config.llmApiEndpoint;
      if (config.llmApiKey) this.llmApiKeyInput.value = config.llmApiKey;
      if (config.llmModel) this.llmModelInput.value = config.llmModel;
      if (config.llmProvider) this.llmProviderSelect.value = config.llmProvider;
      if (config.llmDeploymentName)
        this.llmDeploymentNameInput.value = config.llmDeploymentName;
      if (config.llmApiVersion)
        this.llmApiVersionInput.value = config.llmApiVersion;

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
        const aiDisabled = localStorage.getItem("aiDisabled") === "true";
        if (!aiDisabled) {
          try {
            const healthy = await this.llmService.healthCheck();
            if (!healthy) {
              localStorage.setItem("aiDisabled", "true");
              this.showStatus(
                "AI service unreachable – continuing without AI features. You can re-enable later in Settings.",
                "error",
              );
            }
          } catch (hcErr) {
            console.warn("LLM health check error:", hcErr);
            localStorage.setItem("aiDisabled", "true");
            this.showStatus(
              "AI service check failed – AI features disabled for this session.",
              "error",
            );
          }
        }
      }

      // Populate snooze options
      this.populateSnoozeOptions(config.snoozeOptions);
    } catch (error) {
      console.error("Error loading configuration:", error);
    }
  }

  private async loadAvailableAccounts(): Promise<void> {
    try {
      this.availableAccounts =
        await this.configurationService.getAvailableAccounts();
    } catch (error) {
      console.error("Error loading available accounts:", error);
      this.availableAccounts = [];
    }
  }

  private populateAccountFilter(selectedAccounts: string[]): void {
    this.accountFilterSelect.innerHTML = "";

    // If only one account is available (common in Office.js context), simplify the UI
    if (this.availableAccounts.length === 1) {
      const account = this.availableAccounts[0];
      const option = document.createElement("option");
      option.value = account;
      option.textContent = `${account} (Current)`;
      option.selected = true;
      this.accountFilterSelect.appendChild(option);
      // Disable to indicate single-context scope
      this.accountFilterSelect.disabled = true;
      return;
    }

    this.accountFilterSelect.disabled = false;
    this.availableAccounts.forEach((account) => {
      const option = document.createElement("option");
      option.value = account;
      option.textContent = account;
      option.selected = selectedAccounts.includes(account);
      this.accountFilterSelect.appendChild(option);
    });
  }

  private populateSnoozeOptions(options: SnoozeOption[]): void {
    console.log("populateSnoozeOptions called with:", options);
    this.snoozeOptionsSelect.innerHTML = "";

    if (!options || options.length === 0) {
      console.warn("No snooze options provided, using default options");
      // Fallback to default options
      options = [
        { label: "15 minutes", value: 15 },
        { label: "1 hour", value: 60 },
        { label: "4 hours", value: 240 },
        { label: "1 day", value: 1440 },
        { label: "3 days", value: 4320 },
        { label: "1 week", value: 10080 },
        { label: "Custom...", value: 0, isCustom: true },
      ];
    }

    options.forEach((option) => {
      const optionElement = document.createElement("option");
      optionElement.value = option.value.toString();
      optionElement.textContent = option.label;
      optionElement.dataset.isCustom = option.isCustom?.toString() || "false";
      this.snoozeOptionsSelect.appendChild(optionElement);
      console.log(
        `Added snooze option: ${option.label} (${option.value} minutes)`,
      );
    });

    console.log(
      `Total snooze options added: ${this.snoozeOptionsSelect.children.length}`,
    );
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
        llmProvider:
          (this.llmProviderSelect.value as "azure" | "dial" | "openai") ||
          undefined,
        llmDeploymentName: this.llmDeploymentNameInput.value.trim(),
        llmApiVersion: this.llmApiVersionInput.value.trim(),
        showSnoozedEmails: this.showSnoozedEmailsCheckbox.checked,
        showDismissedEmails: this.showDismissedEmailsCheckbox.checked,
        selectedAccounts: Array.from(
          this.accountFilterSelect.selectedOptions,
        ).map((option) => option.value),
      };
      await this.configurationService.saveConfiguration(config);
    } catch (error) {
      console.error("Error saving configuration:", error);
    }
  }

  // Enhanced analyzeEmails with progress tracking
  private async analyzeEmails(): Promise<void> {
    try {
      this.debugLog("Starting email analysis");
      this.setLoadingState(true);
      this.hideStatus();
      this.updateProgress(
        0,
        "Initializing analysis...",
        "Getting ready to analyze your emails",
      );

      const emailCount = parseInt(this.emailCountSelect.value);
      const daysBack = parseInt(this.daysBackSelect.value);
      const selectedAccounts = Array.from(
        this.accountFilterSelect.selectedOptions,
      ).map((option) => option.value);

      this.debugLog("Analysis parameters", {
        emailCount,
        daysBack,
        selectedAccounts,
      });

      // Get current configuration and pass it to the EmailAnalysisService
      const config = await this.configurationService.getConfiguration();
      this.emailAnalysisService.setConfiguration(config);

      this.updateProgress(
        20,
        "Fetching emails...",
        `Looking for emails from the last ${daysBack} days`,
      );

      // Simulate progress updates during analysis
      const progressInterval = setInterval(() => {
        const currentWidth = parseInt(this.progressFill.style.width) || 20;
        if (currentWidth < 80) {
          this.updateProgress(
            currentWidth + 10,
            "Analyzing emails...",
            "Processing email content and threads",
          );
        }
      }, 500);

      const followupEmails = await this.emailAnalysisService.analyzeEmails(
        emailCount,
        daysBack,
        selectedAccounts,
      );

      this.debugLog("Analysis completed", {
        foundEmails: followupEmails.length,
        emails: followupEmails.map((e) => ({
          subject: e.subject,
          sentDate: e.sentDate,
          recipients: e.recipients,
        })),
      });

      clearInterval(progressInterval);
      this.updateProgress(100, "Analysis complete!", "Preparing results");

      this.allEmails = followupEmails;
      this.filteredEmails = [...followupEmails];

      await this.saveConfiguration();
      this.updateStatistics();
      this.displayEmails(this.filteredEmails);

      if (followupEmails.length > 0) {
        this.showStatus(
          `Found ${followupEmails.length} email(s) that may need follow-up`,
          "success",
        );
      } else {
        this.showStatus("No emails requiring follow-up found", "success");
      }
    } catch (error) {
      console.error("Error analyzing emails:", (error as Error).message);
      this.showStatus(
        `Error analyzing emails: ${(error as Error).message}`,
        "error",
      );
      this.displayEmails([]);
    } finally {
      this.setLoadingState(false);
    }
  }

  // Progress tracking methods
  private updateProgress(
    percentage: number,
    step: string,
    detail: string,
  ): void {
    this.progressFill.style.width = `${percentage}%`;
    this.loadingStep.textContent = step;
    this.loadingDetail.textContent = detail;
  }

  // Statistics methods
  private updateStatistics(): void {
    const total = this.allEmails.length;
    const needingFollowup = this.filteredEmails.length;
    const highPriority = this.allEmails.filter(
      (email) => email.priority === "high",
    ).length;
    const avgDays =
      total > 0
        ? Math.round(
            this.allEmails.reduce(
              (sum, email) => sum + email.daysWithoutResponse,
              0,
            ) / total,
          )
        : 0;

    this.totalEmailsAnalyzedSpan.textContent = total.toString();
    this.needingFollowupSpan.textContent = needingFollowup.toString();
    this.highPriorityCountSpan.textContent = highPriority.toString();
    this.avgResponseTimeSpan.textContent = avgDays.toString();
  }

  private toggleStatsDashboard(show: boolean): void {
    if (show) {
      this.statsDashboard.classList.add("show");
      this.showStatsButton.style.display = "none";
    } else {
      this.statsDashboard.classList.remove("show");
      this.showStatsButton.style.display = "inline-block";
    }
  }

  private toggleAdvancedFilters(): void {
    this.advancedFilters.classList.toggle("show");
    const isShown = this.advancedFilters.classList.contains("show");
    this.toggleAdvancedFiltersButton.textContent = isShown
      ? "Hide Advanced Filters"
      : "Advanced Filters";
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
    this.filteredEmails = this.allEmails.filter((email) => {
      // Priority filter
      if (
        this.priorityFilter.value &&
        email.priority !== this.priorityFilter.value
      ) {
        return false;
      }

      // Response time filter
      if (this.responseTimeFilter.value) {
        const days = email.daysWithoutResponse;
        const range = this.responseTimeFilter.value;

        if (range === "1-3" && (days < 1 || days > 3)) return false;
        if (range === "4-7" && (days < 4 || days > 7)) return false;
        if (range === "8-14" && (days < 8 || days > 14)) return false;
        if (range === "15+" && days < 15) return false;
      }

      // Subject filter
      if (
        this.subjectFilter.value &&
        !email.subject
          .toLowerCase()
          .includes(this.subjectFilter.value.toLowerCase())
      ) {
        return false;
      }

      // Sender filter (check recipients for sent emails)
      if (this.senderFilter.value) {
        const senderText = this.senderFilter.value.toLowerCase();
        const hasMatchingRecipient = email.recipients.some((recipient) =>
          recipient.toLowerCase().includes(senderText),
        );
        if (!hasMatchingRecipient) return false;
      }

      // AI suggestion filter
      if (this.aiSuggestionFilter.value) {
        const hasAI = !!email.llmSuggestion;
        if (this.aiSuggestionFilter.value === "with-ai" && !hasAI) return false;
        if (this.aiSuggestionFilter.value === "without-ai" && hasAI)
          return false;
      }

      return true;
    });

    this.displayEmails(this.filteredEmails);
    this.updateStatistics();
  }

  private clearAllFilters(): void {
    this.priorityFilter.value = "";
    this.responseTimeFilter.value = "";
    this.subjectFilter.value = "";
    this.senderFilter.value = "";
    this.aiSuggestionFilter.value = "";
    this.applyFilters();
  }

  // Enhanced email display with priority classes and metadata
  private createEmailElement(email: FollowupEmail): HTMLDivElement {
    const emailDiv = document.createElement("div");
    emailDiv.className = `email-item priority-${email.priority}`;

    const priorityBadge =
      email.priority === "high"
        ? "🔴"
        : email.priority === "medium"
          ? "🟡"
          : "🟢";
    const accountBadge = email.accountEmail ? `📧 ${email.accountEmail}` : "";
    const llmIndicator = email.llmSummary ? "🤖" : "";

    // Calculate confidence score (mock implementation)
    const confidence = Math.min(
      100,
      Math.max(
        0,
        email.daysWithoutResponse * 10 +
          (email.priority === "high"
            ? 30
            : email.priority === "medium"
              ? 15
              : 0) +
          (email.llmSuggestion ? 20 : 0),
      ),
    );

    const toLine = this.formatRecipientsAsEmails(email.recipients).join(", ");
    const summaryHtml = this.renderSummaryHtml(email.summary);

    emailDiv.innerHTML = `
            <div class="email-header">
                <span class="priority-badge">${priorityBadge}</span>
                <span class="account-badge">${accountBadge}</span>
                <span class="llm-indicator">${llmIndicator}</span>
            </div>
            <div class="email-subject">${this.escapeHtml(email.subject)}</div>
            <div class="email-metadata">
                <div class="metadata-item">
            <span>📤 To: ${this.escapeHtml(toLine)}</span>
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
        <div class="email-summary">${summaryHtml}</div>
            ${email.llmSuggestion ? `<div class="llm-suggestion"><strong>AI Suggestion:</strong> ${this.escapeHtml(email.llmSuggestion)}</div>` : ""}
            <div class="email-actions">
                <button class="action-button" data-email-id="${email.id}" data-action="reply">Reply</button>
                <button class="action-button" data-email-id="${email.id}" data-action="forward">Forward</button>
                <button class="action-button" data-email-id="${email.id}" data-action="snooze">Snooze</button>
                <button class="action-button" data-email-id="${email.id}" data-action="dismiss">Dismiss</button>
                ${email.threadMessages.length > 1 ? `<button class="action-button" data-email-id="${email.id}" data-action="view-thread">View Thread (${email.threadMessages.length})</button>` : ""}
            </div>
        `;

    // Attach event listeners to action buttons
    const actionButtons = emailDiv.querySelectorAll(".action-button");
    actionButtons.forEach((button) => {
      button.addEventListener("click", (e) => {
        this.handleEmailAction(e).catch((error) => {
          console.error("Error handling email action:", error);
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
      this.emptyStateDiv.style.display = "block";
      return;
    }

    // Requirement: always display most recent emails first (newest -> oldest)
    // Even if analysis service internally sorts by priority, we enforce recency ordering for UI display.
    const recentFirst = [...emails].sort(
      (a, b) => b.sentDate.getTime() - a.sentDate.getTime(),
    );

    this.emailListDiv.innerHTML = "";
    recentFirst.forEach((email) => {
      const emailElement = this.createEmailElement(email);
      this.emailListDiv.appendChild(emailElement);
    });

    this.emailListDiv.style.display = "block";
  }

  // Handle email action buttons
  private async handleEmailAction(event: Event): Promise<void> {
    const button = event.target as HTMLButtonElement;
    const emailId = button.dataset.emailId!;
    const action = button.dataset.action!;

    switch (action) {
      case "reply":
        this.replyToEmail(emailId);
        break;
      case "forward":
        this.forwardEmail(emailId);
        break;
      case "snooze":
        await this.showSnoozeModal(emailId);
        break;
      case "dismiss":
        this.dismissEmail(emailId);
        break;
      case "view-thread":
        this.showThreadView(emailId);
        break;
    }
  }

  // Thread view methods
  private showThreadView(emailId: string): void {
    const email = this.allEmails.find((e) => e.id === emailId);
    if (!email) return;

    this.threadSubject.textContent = `Thread: ${email.subject}`;
    this.threadBody.innerHTML = "";

    email.threadMessages.forEach((message, index) => {
      const messageDiv = document.createElement("div");
      messageDiv.className = "thread-message";

      const toHeader = this.formatRecipientsAsEmails(message.to).join(", ");
      messageDiv.innerHTML = `
                <div class="thread-message-header">
                    <strong>Message ${index + 1}</strong> - ${message.sentDate.toLocaleDateString()} ${message.sentDate.toLocaleTimeString()}
            <br>From: ${this.escapeHtml(message.from)} | To: ${this.escapeHtml(toHeader)}
                </div>
                <div class="thread-message-body">
                    ${this.escapeHtml(message.body)}
                </div>
            `;

      this.threadBody.appendChild(messageDiv);
    });

    this.threadModal.style.display = "block";
  }

  private hideThreadModal(): void {
    this.threadModal.style.display = "none";
  }

  // Load cached results from ribbon commands
  private async loadCachedResults(): Promise<void> {
    try {
      const cachedEmails =
        await this.configurationService.getCachedAnalysisResults();
      if (cachedEmails.length > 0) {
        this.allEmails = cachedEmails;
        this.filteredEmails = [...cachedEmails];
        this.updateStatistics();
        this.displayEmails(this.filteredEmails);
        this.showStatus(
          `Loaded ${cachedEmails.length} cached results`,
          "success",
        );
      }
    } catch (error) {
      console.error("Error loading cached results:", error);
    }
  }

  private async replyToEmail(_emailId: string): Promise<void> {
    try {
      if (!(window as any).Office || !Office.context?.mailbox) {
        this.showStatus("Reply requires Outlook mailbox context.", "error");
        return;
      }
      const email = this.allEmails.find((e) => e.id === _emailId);
      if (!email) {
        console.warn("Email not found for reply", _emailId);
        return;
      }

      // PRIMARY STRATEGY: Open the email item so the user can use the native "Reply All" button.
      // This ensures thread context is preserved perfectly and works without EWS write permissions.
      const hasDisplayMessage =
        typeof (Office.context?.mailbox as any)?.displayMessageForm ===
        "function";

      if (hasDisplayMessage && this.isLikelyEwsItemId(_emailId)) {
        try {
          Office.context.mailbox.displayMessageForm(_emailId);
          // Optional: Show a toast/status telling the user what to do
          this.showStatus(
            "Opened email. Please click 'Reply All' in Outlook to respond.",
            "success",
          );
          return;
        } catch (e) {
          console.warn("Failed to display message form:", e);
          this.showStatus(
            "Failed to open email. Please find it in your Sent Items to reply.",
            "error",
          );
        }
      } else {
        this.showStatus(
          "Cannot open email (ID format not supported or API unavailable).",
          "error",
        );
      }
    } catch (err) {
      console.error("Error opening message for reply", err);
      this.showStatus("Failed to open message in Outlook", "error");
    }
  }

  private async forwardEmail(_emailId: string): Promise<void> {
    try {
      if (!(window as any).Office || !Office.context?.mailbox) {
        this.showStatus("Forward requires Outlook mailbox context.", "error");
        return;
      }
      const email = this.allEmails.find((e) => e.id === _emailId);
      if (!email) {
        console.warn("Email not found for forward", _emailId);
        return;
      }
      // Preferred: create a Forward draft via EWS and open it using native compose
      const isTestEnv = typeof (globalThis as any).jest !== "undefined";
      const hasDisplay =
        typeof (Office.context?.mailbox as any)?.displayMessageForm ===
        "function";
      const hasEws =
        typeof (Office.context?.mailbox as any)?.makeEwsRequestAsync ===
        "function";
      if (!isTestEnv && hasDisplay && hasEws) {
        try {
          const lastMsg =
            email.threadMessages && email.threadMessages.length > 0
              ? email.threadMessages[email.threadMessages.length - 1]
              : undefined;
          const referenceId = lastMsg?.id || email.id;
          const referenceChangeKey = lastMsg?.changeKey;
          if (this.isLikelyEwsItemId(referenceId)) {
            const draftId = await this.createForwardDraft(
              referenceId,
              referenceChangeKey,
            );
            if (draftId) {
              Office.context.mailbox.displayMessageForm(draftId);
              return;
            }
          } else {
            console.warn(
              "Reference ID does not look like an EWS ItemId. Skipping EWS path and using compose fallback.",
            );
          }
        } catch (e) {
          console.warn(
            "Forward draft creation failed, falling back to compose builder:",
            e,
          );
        }
      }
      // Fallback (e.g., unit tests or non-Outlook host): synthesize a Forward compose window
      const lastMessage =
        email.threadMessages && email.threadMessages.length > 0
          ? email.threadMessages[email.threadMessages.length - 1]
          : undefined;
      if (!lastMessage) {
        console.warn("No thread messages available for forward", _emailId);
        return;
      }
      let subject = lastMessage.subject || email.subject || "";
      if (!/^fw:|^fwd:/i.test(subject)) subject = `FW: ${subject}`;
      const rawBody = lastMessage.body || "";
      const looksHtml = /<\w+[^>]*>/.test(rawBody);
      let sanitizedBody = rawBody;
      if (!looksHtml) {
        sanitizedBody = this.escapeHtml(rawBody).replace(/\n/g, "<br>");
      } else {
        sanitizedBody = rawBody
          .replace(/<script[\s\S]*?<\/script>/gi, "")
          .replace(/<style[\s\S]*?<\/style>/gi, "");
      }
      const quotedSeparator = "<br><br>----- Forwarded Message -----<br>";
      // Trim to avoid OWA htmlBody parameter range issues
      const MAX_BODY = 15000;
      const bodyTrimmed =
        sanitizedBody.length > MAX_BODY
          ? sanitizedBody.slice(0, MAX_BODY) + "…"
          : sanitizedBody;
      const htmlBody = `<br><br>${quotedSeparator}${bodyTrimmed}`;
      Office.context.mailbox.displayNewMessageForm({
        toRecipients: [],
        subject,
        htmlBody,
        attachments: [],
      });
    } catch (err) {
      console.error("Error opening message for forward", err);
      this.showStatus("Failed to open message in Outlook", "error");
    }
  }

  // EWS helpers to create native drafts for Reply All and Forward
  private async createReplyAllDraft(
    itemId: string,
    changeKey?: string,
  ): Promise<string> {
    // OWA often requires ChangeKey for write operations; fetch if not provided
    let ck = changeKey;
    if (!ck) {
      try {
        const info = await this.getItemIdWithChangeKey(itemId);
        if (info?.changeKey) ck = info.changeKey;
      } catch (e) {
        console.warn(
          "Failed to fetch ChangeKey via GetItem; proceeding without it.",
          e,
        );
      }
    }
    const envelope = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
                             xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" 
                             xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" 
                             xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
    <soap:Header>
        <t:RequestServerVersion Version="Exchange2013" />
    </soap:Header>
    <soap:Body>
        <m:CreateItem MessageDisposition="SaveOnly">
            <m:Items>
                <t:ReplyAllToItem>
            <t:ReferenceItemId Id="${itemId}"${ck ? ` ChangeKey="${ck}"` : ""} />
                    <t:NewBodyContent BodyType="HTML"></t:NewBodyContent>
                </t:ReplyAllToItem>
            </m:Items>
        </m:CreateItem>
    </soap:Body>
</soap:Envelope>`;
    return this.createDraftViaEws(envelope);
  }

  private async createForwardDraft(
    itemId: string,
    changeKey?: string,
  ): Promise<string> {
    // OWA often requires ChangeKey for write operations; fetch if not provided
    let ck = changeKey;
    if (!ck) {
      try {
        const info = await this.getItemIdWithChangeKey(itemId);
        if (info?.changeKey) ck = info.changeKey;
      } catch (e) {
        console.warn(
          "Failed to fetch ChangeKey via GetItem; proceeding without it.",
          e,
        );
      }
    }
    const envelope = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
                             xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" 
                             xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" 
                             xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
    <soap:Header>
        <t:RequestServerVersion Version="Exchange2013" />
    </soap:Header>
    <soap:Body>
        <m:CreateItem MessageDisposition="SaveOnly">
            <m:Items>
                <t:ForwardItem>
            <t:ReferenceItemId Id="${itemId}"${ck ? ` ChangeKey="${ck}"` : ""} />
                    <t:NewBodyContent BodyType="HTML"></t:NewBodyContent>
                </t:ForwardItem>
            </m:Items>
        </m:CreateItem>
    </soap:Body>
</soap:Envelope>`;
    return this.createDraftViaEws(envelope);
  }

  // Fetch ItemId+ChangeKey for a given item via EWS GetItem
  private getItemIdWithChangeKey(
    itemId: string,
  ): Promise<{ id: string; changeKey?: string } | null> {
    const envelope = `<?xml version="1.0" encoding="utf-8"?>
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
      </m:ItemShape>
      <m:ItemIds>
    <t:ItemId Id="${itemId}" />
      </m:ItemIds>
    </m:GetItem>
  </soap:Body>
</soap:Envelope>`;
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.makeEwsRequestAsync(envelope, (res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded) {
            try {
              const parser = new DOMParser();
              const doc = parser.parseFromString(res.value, "text/xml");
              // If EWS returned an error class, bail out
              const errorResp = doc.querySelector('[ResponseClass="Error"]');
              if (errorResp) {
                const code =
                  errorResp.querySelector("ResponseCode")?.textContent ||
                  "Unknown error";
                console.warn("EWS GetItem error:", code);
                return resolve(null);
              }
              const els = doc.getElementsByTagName("*");
              for (let i = 0; i < els.length; i++) {
                const el = els[i];
                if (el.localName === "ItemId") {
                  const id = el.getAttribute("Id") || "";
                  const ck = el.getAttribute("ChangeKey") || undefined;
                  return resolve({ id, changeKey: ck });
                }
              }
              console.warn("GetItem returned no ItemId. Raw response follows.");
              console.warn(res.value);
              resolve(null);
            } catch (e) {
              reject(e);
            }
          } else {
            reject(new Error(res.error?.message || "GetItem failed"));
          }
        });
      } catch (err) {
        reject(err);
      }
    });
  }

  // Heuristic: EWS ItemIds are typically long, base64-like strings; filter out obvious synthetic ids
  private isLikelyEwsItemId(id: string | undefined | null): boolean {
    if (!id) return false;
    if (id.length < 24) return false;
    // Disallow clearly synthetic short ids used in artificial threads
    const looksBase64ish = /^[A-Za-z0-9+/=_-]+$/.test(id);
    return looksBase64ish;
  }

  private createDraftViaEws(envelope: string): Promise<string> {
    return new Promise((resolve, reject) => {
      try {
        Office.context.mailbox.makeEwsRequestAsync(envelope, (res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded) {
            try {
              const validation = this.xmlValidate(res.value);
              if (!validation.isValid) {
                console.warn(
                  "EWS CreateItem response indicates error:",
                  validation.error,
                );
              }
              const id = this.parseCreateItemResponseForId(res.value);
              if (id) resolve(id);
              else reject(new Error("Draft ItemId not found"));
            } catch (e) {
              reject(e);
            }
          } else {
            reject(new Error(res.error?.message || "CreateItem failed"));
          }
        });
      } catch (err) {
        reject(err);
      }
    });
  }

  private parseCreateItemResponseForId(xml: string): string | null {
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(xml, "text/xml");
      // Find ItemId Id attribute in a namespace-agnostic way
      const els = doc.getElementsByTagName("*");
      for (let i = 0; i < els.length; i++) {
        const el = els[i];
        if (el.localName === "ItemId") {
          const id = el.getAttribute("Id");
          if (id) return id;
        }
      }
      console.warn(
        "EWS CreateItem response did not contain ItemId. Raw response follows for diagnostics.",
      );
      console.warn(xml);
      return null;
    } catch (e) {
      console.error("Failed to parse CreateItem response:", e);
      return null;
    }
  }

  // Light wrapper around XmlParsingService.validateEwsResponse without importing to keep bundle minimal
  private xmlValidate(xml: string): { isValid: boolean; error?: string } {
    try {
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(xml, "text/xml");
      const parserError = xmlDoc.querySelector("parsererror");
      if (parserError) {
        return {
          isValid: false,
          error: `XML Parse Error: ${parserError.textContent}`,
        };
      }
      const nsSOAP = "http://schemas.xmlsoap.org/soap/envelope/";
      const fault = xmlDoc.getElementsByTagNameNS(nsSOAP, "Fault")[0];
      if (fault) {
        const faultString =
          fault.querySelector("faultstring")?.textContent ||
          "Unknown SOAP fault";
        return { isValid: false, error: `SOAP Fault: ${faultString}` };
      }
      const nsMsg =
        "http://schemas.microsoft.com/exchange/services/2006/messages";
      const respMsgs = xmlDoc.getElementsByTagNameNS(
        nsMsg,
        "ResponseMessages",
      )[0];
      if (respMsgs) {
        const errorResp = respMsgs.querySelector('[ResponseClass="Error"]');
        if (errorResp) {
          const code =
            errorResp.querySelector("ResponseCode")?.textContent ||
            "Unknown error";
          const text =
            errorResp.querySelector("MessageText")?.textContent || "";
          return { isValid: false, error: `EWS Error ${code}: ${text}` };
        }
      }
      return { isValid: true };
    } catch (err) {
      return {
        isValid: false,
        error: `Validation error: ${(err as Error).message}`,
      };
    }
  }

  // Modal management methods
  private async showSnoozeModal(emailId: string): Promise<void> {
    this.currentEmailForSnooze = emailId;

    // Ensure snooze options are populated
    try {
      const config = await this.configurationService.getConfiguration();
      this.populateSnoozeOptions(config.snoozeOptions);
      console.log("Snooze options populated:", config.snoozeOptions);
    } catch (error) {
      console.error("Error loading snooze options:", error);
      // Fallback to default options if configuration fails
      const defaultOptions: SnoozeOption[] = [
        { label: "15 minutes", value: 15 },
        { label: "1 hour", value: 60 },
        { label: "4 hours", value: 240 },
        { label: "1 day", value: 1440 },
        { label: "3 days", value: 4320 },
        { label: "1 week", value: 10080 },
        { label: "Custom...", value: 0, isCustom: true },
      ];
      this.populateSnoozeOptions(defaultOptions);
    }

    this.snoozeModal.style.display = "block";
  }

  private hideSnoozeModal(): void {
    this.snoozeModal.style.display = "none";
    this.currentEmailForSnooze = "";
    this.customSnoozeGroup.style.display = "none";
  }

  private showSettingsModal(): void {
    this.settingsModal.style.display = "block";
    this.updateAiStatus(); // Update AI status when opening settings
  }

  private hideSettingsModal(): void {
    this.settingsModal.style.display = "none";
  }

  private handleSnoozeOptionChange(): void {
    const selectedOption = this.snoozeOptionsSelect.selectedOptions[0];
    const isCustom = selectedOption?.dataset.isCustom === "true";

    if (isCustom) {
      this.customSnoozeGroup.style.display = "block";
      // Set default to 1 hour from now
      const defaultDate = new Date();
      defaultDate.setHours(defaultDate.getHours() + 1);
      this.customSnoozeDate.value = defaultDate.toISOString().slice(0, 16);
    } else {
      this.customSnoozeGroup.style.display = "none";
    }
  }

  private async confirmSnooze(): Promise<void> {
    if (!this.currentEmailForSnooze) return;

    const selectedOption = this.snoozeOptionsSelect.selectedOptions[0];
    const isCustom = selectedOption?.dataset.isCustom === "true";

    try {
      if (isCustom) {
        const customDate = new Date(this.customSnoozeDate.value);
        if (customDate <= new Date()) {
          this.showStatus("Snooze time must be in the future", "error");
          return;
        }
        this.emailAnalysisService.snoozeEmailUntil(
          this.currentEmailForSnooze,
          customDate,
        );
      } else {
        const minutes = parseInt(selectedOption.value);
        this.emailAnalysisService.snoozeEmail(
          this.currentEmailForSnooze,
          minutes,
        );
      }

      // Remove email from current display
      this.removeEmailFromDisplay(this.currentEmailForSnooze);
      this.showStatus("Email snoozed successfully", "success");
      this.hideSnoozeModal();
    } catch (error) {
      console.error("Error snoozing email:", (error as Error).message);
      this.showStatus(
        `Error snoozing email: ${(error as Error).message}`,
        "error",
      );
    }
  }

  private async dismissEmail(emailId: string): Promise<void> {
    this.emailAnalysisService.dismissEmail(emailId);
    this.removeEmailFromDisplay(emailId);
    this.showStatus("Email dismissed", "success");
  }

  private removeEmailFromDisplay(emailId: string): void {
    const emailElement = document
      .querySelector(`[data-email-id="${emailId}"]`)
      ?.closest(".email-item");
    if (emailElement) {
      emailElement.remove();

      // Check if list is now empty
      if (this.emailListDiv.children.length === 0) {
        this.hideAllStates();
        this.emptyStateDiv.style.display = "block";
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
        await this.configurationService.updateLlmSettings(
          endpoint,
          apiKey,
          enableSummary,
          enableSuggestions,
        );

        // Update LLM service with proper configuration object
        const config = await this.configurationService.getConfiguration();
        this.llmService = new LlmService(config, this.retryService);
        this.emailAnalysisService.setLlmService(this.llmService);

        this.showStatus("Settings saved successfully", "success");
      } else if (endpoint || apiKey) {
        this.showStatus(
          "Both endpoint and API key are required for LLM integration",
          "error",
        );
        return;
      }

      // Update AI status after saving
      this.updateAiStatus();
      this.hideSettingsModal();
    } catch (error) {
      console.error("Error saving settings:", (error as Error).message);
      this.showStatus(
        `Error saving settings: ${(error as Error).message}`,
        "error",
      );
    }
  }

  private showStatus(message: string, type: "success" | "error"): void {
    this.statusDiv.textContent = message;
    this.statusDiv.className = `status ${type}`;
    this.statusDiv.style.display = "block";

    // Auto-hide success messages after 5 seconds
    if (type === "success") {
      setTimeout(() => this.hideStatus(), 5000);
    }
  }

  private hideStatus(): void {
    this.statusDiv.style.display = "none";
  }

  private escapeHtml(text: string): string {
    const div = document.createElement("div");
    div.textContent = text;
    return div.innerHTML;
  }

  // Decode common HTML entities to plain text
  private decodeHtmlEntities(text: string): string {
    const parser = new DOMParser();
    const doc = parser.parseFromString(text, "text/html");
    return doc.documentElement.textContent || "";
  }

  // Render a short summary as sanitized HTML: decode entities, remove empty lines, preserve new lines (as <br>), no &nbsp;, and cap font size via CSS
  private renderSummaryHtml(raw: string): string {
    if (!raw) return "";
    // Normalize non-breaking spaces and decode entities
    let normalized = raw.replace(/&nbsp;/gi, " ").replace(/\u00A0/g, " ");
    normalized = this.decodeHtmlEntities(normalized);
    // Preserve new lines and remove empty lines
    const lines = normalized
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter((l) => l.length > 0);
    const safe = lines.map((l) => this.escapeHtml(l));
    return safe.join("<br>");
  }

  // Ensure recipients are shown as email addresses
  private formatRecipientsAsEmails(recipients: string[]): string[] {
    return (recipients || [])
      .map((r) => this.extractEmail(r))
      .filter((r) => !!r);
  }

  private extractEmail(text: string): string {
    if (!text) return "";
    const angle = text.match(/<([^>]+)>/);
    if (angle && angle[1]) return angle[1].trim();
    const match = text.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
    return match ? match[0] : text;
  }

  // Refresh emails (re-run analysis)
  private async refreshEmails(): Promise<void> {
    await this.analyzeEmails();
  }

  // Loading state management
  private setLoadingState(isLoading: boolean): void {
    if (isLoading) {
      this.hideAllStates();
      this.loadingDiv.style.display = "block";
      this.analyzeButton.disabled = true;
      this.refreshButton.disabled = true;
    } else {
      this.loadingDiv.style.display = "none";
      this.analyzeButton.disabled = false;
      this.refreshButton.disabled = false;
    }
  }

  // Hide all display states
  private hideAllStates(): void {
    this.loadingDiv.style.display = "none";
    this.emptyStateDiv.style.display = "none";
    this.emailListDiv.style.display = "none";
  }

  // Initialize the task pane
  // API status tracking
  private apiStatus = {
    ewsAvailable: false,
    graphAvailable: false,
    currentEmailAvailable: false,
    bulkAccessBlocked: false,
  };

  // Current analysis mode (used for tracking state and future features)
  public getCurrentAnalysisMode(): "bulk" | "current" | "manual" {
    return this._currentAnalysisMode;
  }
  private _currentAnalysisMode: "bulk" | "current" | "manual" = "bulk";

  public async initialize(): Promise<void> {
    await this.loadConfiguration();
    this.hideAllStates();
    this.emptyStateDiv.style.display = "block";

    // Load AI disable state
    const aiDisabled = localStorage.getItem("aiDisabled") === "true";
    this.enableAiFeaturesCheckbox.checked = !aiDisabled;

    this.updateAiStatus(); // Check AI status on initialization

    // Check API availability and set up fallback mode UI
    await this.checkApiAvailability();
    this.setupFallbackModeUI();

    // Add diagnostic button event listener
    const diagnosticButton = document.getElementById("diagnosticButton");
    if (diagnosticButton) {
      diagnosticButton.addEventListener("click", () =>
        this.openDiagnosticModal(),
      );
    }

    // Add close diagnostic modal event listener
    const closeDiagnosticButton = document.querySelector(
      "#diagnosticModal .modal-close",
    );
    if (closeDiagnosticButton) {
      closeDiagnosticButton.addEventListener("click", () =>
        this.closeDiagnosticModal(),
      );
    }

    // Add event listener for the "Close" button in the modal footer
    const closeDiagnosticFooterButton =
      document.getElementById("closeDiagnostic");
    if (closeDiagnosticFooterButton) {
      closeDiagnosticFooterButton.addEventListener("click", () =>
        this.closeDiagnosticModal(),
      );
    }

    // Add event listeners for diagnostic tests
    const testOfficeButton = document.getElementById("testOfficeContext");
    const testMailboxButton = document.getElementById("testMailboxAccess");
    const testAccountsButton = document.getElementById("testAccountDetection");
    const testEmailsButton = document.getElementById("testEmailReading");
    const testGraphApiButton = document.getElementById("testGraphApi");

    if (testOfficeButton) {
      testOfficeButton.addEventListener("click", () =>
        this.testOfficeContext(),
      );
    }
    if (testMailboxButton) {
      testMailboxButton.addEventListener("click", () =>
        this.testMailboxAccess(),
      );
    }
    if (testAccountsButton) {
      testAccountsButton.addEventListener("click", () =>
        this.testAccountDetection(),
      );
    }
    if (testEmailsButton) {
      testEmailsButton.addEventListener("click", () => this.testEmailReading());
    }
    if (testGraphApiButton) {
      testGraphApiButton.addEventListener("click", () => this.testGraphApi());
    }
  }

  /**
   * Check API availability silently on startup
   */
  private async checkApiAvailability(): Promise<void> {
    // Check if Office.context is available
    if (
      typeof Office === "undefined" ||
      !Office.context ||
      !Office.context.mailbox
    ) {
      this.apiStatus.bulkAccessBlocked = true;
      return;
    }

    // Check EWS availability
    if (typeof Office.context.mailbox.makeEwsRequestAsync === "function") {
      // EWS API exists, but we need to test if it actually works
      try {
        const ewsWorks = await this.testEwsQuietly();
        this.apiStatus.ewsAvailable = ewsWorks;
      } catch {
        this.apiStatus.ewsAvailable = false;
      }
    }

    // Check current email availability
    if (Office.context.mailbox.item) {
      this.apiStatus.currentEmailAvailable = true;
    }

    // If EWS and Graph are both unavailable, mark bulk access as blocked
    if (!this.apiStatus.ewsAvailable && !this.apiStatus.graphAvailable) {
      this.apiStatus.bulkAccessBlocked = true;
    }
  }

  /**
   * Test EWS quietly without showing UI
   */
  private async testEwsQuietly(): Promise<boolean> {
    return new Promise((resolve) => {
      const timeout = setTimeout(() => resolve(false), 5000);

      const testRequest = `<?xml version="1.0" encoding="utf-8"?>
        <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
                       xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                       xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
          <soap:Header>
            <t:RequestServerVersion Version="Exchange2013" />
          </soap:Header>
          <soap:Body>
            <m:GetFolder>
              <m:FolderShape><t:BaseShape>IdOnly</t:BaseShape></m:FolderShape>
              <m:FolderIds><t:DistinguishedFolderId Id="inbox" /></m:FolderIds>
            </m:GetFolder>
          </soap:Body>
        </soap:Envelope>`;

      try {
        Office.context.mailbox.makeEwsRequestAsync(testRequest, (result) => {
          clearTimeout(timeout);
          resolve(result.status === Office.AsyncResultStatus.Succeeded);
        });
      } catch {
        clearTimeout(timeout);
        resolve(false);
      }
    });
  }

  /**
   * Set up the fallback mode UI based on API availability
   */
  private setupFallbackModeUI(): void {
    const banner = document.getElementById("apiStatusBanner");
    const fallbackSection = document.getElementById("fallbackModeSection");
    const bulkAnalysisStatus = document.getElementById("bulkAnalysisStatus");
    const currentEmailStatus = document.getElementById("currentEmailStatus");
    const ewsStatusIcon = document.getElementById("ewsStatusIcon");
    const ewsStatusText = document.getElementById("ewsStatusText");
    const graphStatusIcon = document.getElementById("graphStatusIcon");
    const graphStatusText = document.getElementById("graphStatusText");
    const currentEmailStatusIcon = document.getElementById(
      "currentEmailStatusIcon",
    );
    const currentEmailStatusText = document.getElementById(
      "currentEmailStatusText",
    );
    const toggleApiDetails = document.getElementById("toggleApiDetails");
    const apiStatusDetails = document.getElementById("apiStatusDetails");

    // Update API status display
    if (ewsStatusIcon && ewsStatusText) {
      if (this.apiStatus.ewsAvailable) {
        ewsStatusIcon.textContent = "✅";
        ewsStatusText.textContent = "EWS (Exchange Web Services): Available";
      } else {
        ewsStatusIcon.textContent = "❌";
        ewsStatusText.textContent =
          "EWS (Exchange Web Services): Blocked by organization";
      }
    }

    if (graphStatusIcon && graphStatusText) {
      if (this.apiStatus.graphAvailable) {
        graphStatusIcon.textContent = "✅";
        graphStatusText.textContent = "Microsoft Graph API: Available";
      } else {
        graphStatusIcon.textContent = "❌";
        graphStatusText.textContent =
          "Microsoft Graph API: Requires Azure AD setup";
      }
    }

    if (currentEmailStatusIcon && currentEmailStatusText) {
      if (this.apiStatus.currentEmailAvailable) {
        currentEmailStatusIcon.textContent = "✅";
        currentEmailStatusText.textContent = "Current Email Access: Available";
      } else {
        currentEmailStatusIcon.textContent = "⚠️";
        currentEmailStatusText.textContent =
          "Current Email Access: Open an email to use";
      }
    }

    // Toggle API details
    if (toggleApiDetails && apiStatusDetails) {
      toggleApiDetails.addEventListener("click", () => {
        const isHidden = apiStatusDetails.style.display === "none";
        apiStatusDetails.style.display = isHidden ? "block" : "none";
        toggleApiDetails.textContent = isHidden
          ? "Hide Details"
          : "Show Details";
      });
    }

    // Update bulk analysis status
    if (bulkAnalysisStatus) {
      if (this.apiStatus.ewsAvailable || this.apiStatus.graphAvailable) {
        bulkAnalysisStatus.className = "fallback-option-status available";
        bulkAnalysisStatus.textContent = "✅ Available";
      } else {
        bulkAnalysisStatus.className = "fallback-option-status unavailable";
        bulkAnalysisStatus.textContent =
          "❌ Requires EWS or Graph API (blocked by organization)";
      }
    }

    // Update current email status
    if (currentEmailStatus) {
      if (this.apiStatus.currentEmailAvailable) {
        currentEmailStatus.className = "fallback-option-status available";
        currentEmailStatus.textContent =
          "✅ Available - Works with current email";
      } else {
        currentEmailStatus.className = "fallback-option-status unavailable";
        currentEmailStatus.textContent = "⚠️ Open an email first";
      }
    }

    // Show banner and fallback section if bulk access is blocked
    if (this.apiStatus.bulkAccessBlocked) {
      if (banner) {
        banner.style.display = "block";
        banner.className = "api-status-banner error";
      }
      if (fallbackSection) {
        fallbackSection.style.display = "block";
      }
      // Hide the main controls since bulk analysis won't work
      const mainControls = document.querySelector(".controls") as HTMLElement;
      if (mainControls) {
        mainControls.style.display = "none";
      }
    }

    // Set up mode selection handlers
    this.setupModeSelectionHandlers();
  }

  /**
   * Set up event handlers for mode selection
   */
  private setupModeSelectionHandlers(): void {
    const optionBulk = document.getElementById("optionBulkAnalysis");
    const optionCurrent = document.getElementById("optionCurrentEmail");
    const optionManual = document.getElementById("optionManualInput");

    const fallbackSection = document.getElementById("fallbackModeSection");
    const currentEmailSection = document.getElementById("currentEmailSection");
    const manualInputSection = document.getElementById("manualInputSection");
    const mainControls = document.querySelector(".controls") as HTMLElement;

    // Bulk analysis option
    if (optionBulk) {
      optionBulk.addEventListener("click", () => {
        if (this.apiStatus.ewsAvailable || this.apiStatus.graphAvailable) {
          this._currentAnalysisMode = "bulk";
          this.selectOption(optionBulk);
          if (fallbackSection) fallbackSection.style.display = "none";
          if (currentEmailSection) currentEmailSection.style.display = "none";
          if (manualInputSection) manualInputSection.style.display = "none";
          if (mainControls) mainControls.style.display = "block";
        }
      });

      // Disable if not available
      if (!this.apiStatus.ewsAvailable && !this.apiStatus.graphAvailable) {
        optionBulk.classList.add("disabled");
      }
    }

    // Current email option
    if (optionCurrent) {
      optionCurrent.addEventListener("click", () => {
        this._currentAnalysisMode = "current";
        this.selectOption(optionCurrent);
        if (fallbackSection) fallbackSection.style.display = "none";
        if (currentEmailSection) currentEmailSection.style.display = "block";
        if (manualInputSection) manualInputSection.style.display = "none";
        if (mainControls) mainControls.style.display = "none";
        this.loadCurrentEmailInfo();
      });
    }

    // Manual input option
    if (optionManual) {
      optionManual.addEventListener("click", () => {
        this._currentAnalysisMode = "manual";
        this.selectOption(optionManual);
        if (fallbackSection) fallbackSection.style.display = "none";
        if (currentEmailSection) currentEmailSection.style.display = "none";
        if (manualInputSection) manualInputSection.style.display = "block";
        if (mainControls) mainControls.style.display = "none";
      });
    }

    // Back buttons
    const backToModeSelection = document.getElementById("backToModeSelection");
    const backToModeSelectionManual = document.getElementById(
      "backToModeSelectionManual",
    );

    const showFallbackSection = () => {
      if (fallbackSection) fallbackSection.style.display = "block";
      if (currentEmailSection) currentEmailSection.style.display = "none";
      if (manualInputSection) manualInputSection.style.display = "none";
      if (mainControls) mainControls.style.display = "none";
    };

    if (backToModeSelection) {
      backToModeSelection.addEventListener("click", showFallbackSection);
    }
    if (backToModeSelectionManual) {
      backToModeSelectionManual.addEventListener("click", showFallbackSection);
    }

    // Analyze current email button
    const analyzeCurrentBtn = document.getElementById("analyzeCurrentEmailBtn");
    if (analyzeCurrentBtn) {
      analyzeCurrentBtn.addEventListener("click", () =>
        this.analyzeCurrentEmail(),
      );
    }

    // Manual input buttons
    const analyzeManualBtn = document.getElementById("analyzeManualEmailBtn");
    const clearManualBtn = document.getElementById("clearManualInput");

    if (analyzeManualBtn) {
      analyzeManualBtn.addEventListener("click", () =>
        this.analyzeManualEmail(),
      );
    }
    if (clearManualBtn) {
      clearManualBtn.addEventListener("click", () => {
        const textarea = document.getElementById(
          "manualEmailInput",
        ) as HTMLTextAreaElement;
        if (textarea) textarea.value = "";
      });
    }
  }

  /**
   * Select a fallback option visually
   */
  private selectOption(selectedElement: HTMLElement): void {
    document.querySelectorAll(".fallback-option").forEach((el) => {
      el.classList.remove("selected");
    });
    selectedElement.classList.add("selected");
  }

  /**
   * Load current email information
   */
  private loadCurrentEmailInfo(): void {
    const subjectEl = document.getElementById("currentEmailSubject");
    const fromEl = document.getElementById("currentEmailFrom");
    const dateEl = document.getElementById("currentEmailDate");
    const toEl = document.getElementById("currentEmailTo");
    const analyzeBtn = document.getElementById(
      "analyzeCurrentEmailBtn",
    ) as HTMLButtonElement;

    if (!Office.context?.mailbox?.item) {
      if (subjectEl) subjectEl.textContent = "No email selected";
      if (fromEl) fromEl.textContent = "-";
      if (dateEl) dateEl.textContent = "-";
      if (toEl) toEl.textContent = "-";
      if (analyzeBtn) analyzeBtn.disabled = true;
      return;
    }

    const item = Office.context.mailbox.item;

    // Get subject
    if (item.subject) {
      if (subjectEl) subjectEl.textContent = item.subject;
    }

    // Get from
    if (item.from) {
      if (fromEl) {
        fromEl.textContent = item.from.displayName
          ? `${item.from.displayName} <${item.from.emailAddress}>`
          : item.from.emailAddress;
      }
    }

    // Get date
    if (item.dateTimeCreated) {
      if (dateEl) dateEl.textContent = item.dateTimeCreated.toLocaleString();
    }

    // Get to recipients
    if (item.to && item.to.length > 0) {
      if (toEl) {
        const recipients = item.to
          .slice(0, 3)
          .map((r: any) => r.emailAddress)
          .join(", ");
        toEl.textContent = item.to.length > 3 ? `${recipients}...` : recipients;
      }
    }

    // Enable analyze button if we have subject
    if (analyzeBtn && item.subject) {
      analyzeBtn.disabled = false;
    }
  }

  /**
   * Analyze the currently viewed email
   */
  private async analyzeCurrentEmail(): Promise<void> {
    const analyzeBtn = document.getElementById(
      "analyzeCurrentEmailBtn",
    ) as HTMLButtonElement;
    const originalText = analyzeBtn?.textContent || "Analyze This Email";

    try {
      if (analyzeBtn) {
        analyzeBtn.disabled = true;
        analyzeBtn.textContent = "Analyzing...";
      }

      if (!Office.context?.mailbox?.item) {
        this.showStatus("No email is currently open", "error");
        return;
      }

      const item = Office.context.mailbox.item;

      // Get email body
      const bodyPromise = new Promise<string>((resolve) => {
        item.body.getAsync(Office.CoercionType.Text, (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(result.value);
          } else {
            resolve("");
          }
        });
      });

      const body = await bodyPromise;
      const daysWithoutResponse = this.calculateDaysSince(
        item.dateTimeCreated || new Date(),
      );

      // Create a follow-up email object
      const email: FollowupEmail = {
        id: item.itemId || `current-${Date.now()}`,
        subject: item.subject || "(No subject)",
        recipients: item.to?.map((r: any) => r.emailAddress) || [],
        sentDate: item.dateTimeCreated || new Date(),
        body: body,
        summary: body.substring(0, 200) + (body.length > 200 ? "..." : ""),
        priority: "medium",
        daysWithoutResponse: daysWithoutResponse,
        conversationId: (item as any).conversationId || "",
        hasAttachments: false,
        accountEmail: Office.context.mailbox.userProfile?.emailAddress || "",
        threadMessages: [],
        isSnoozed: false,
        isDismissed: false,
      };

      // Calculate priority based on days
      if (daysWithoutResponse >= 7) {
        email.priority = "high";
      } else if (daysWithoutResponse >= 3) {
        email.priority = "medium";
      } else {
        email.priority = "low";
      }

      // If AI is enabled, enhance with LLM
      if (this.llmService && this.enableLlmSummaryCheckbox?.checked) {
        try {
          const llmSummary = await this.llmService.summarizeEmail(body);
          email.llmSummary = llmSummary;
        } catch (llmError) {
          console.warn("LLM summary failed:", llmError);
        }
      }

      if (this.llmService && this.enableLlmSuggestionsCheckbox?.checked) {
        try {
          const suggestions = await this.llmService.generateFollowupSuggestions(
            body,
            `Subject: ${email.subject}, Days without response: ${daysWithoutResponse}`,
          );
          email.llmSuggestion = suggestions.join("\n");
        } catch (llmError) {
          console.warn("LLM suggestion failed:", llmError);
        }
      }

      // Display the result
      this.displaySingleEmailResult(email);
      this.showStatus("Email analyzed successfully!", "success");
    } catch (error) {
      this.showStatus(`Analysis failed: ${(error as Error).message}`, "error");
    } finally {
      if (analyzeBtn) {
        analyzeBtn.disabled = false;
        analyzeBtn.textContent = originalText;
      }
    }
  }

  /**
   * Analyze manually pasted email
   */
  private async analyzeManualEmail(): Promise<void> {
    const textarea = document.getElementById(
      "manualEmailInput",
    ) as HTMLTextAreaElement;
    const analyzeBtn = document.getElementById(
      "analyzeManualEmailBtn",
    ) as HTMLButtonElement;
    const originalText = analyzeBtn?.textContent || "Analyze Email";

    if (!textarea || !textarea.value.trim()) {
      this.showStatus("Please paste email content first", "error");
      return;
    }

    try {
      if (analyzeBtn) {
        analyzeBtn.disabled = true;
        analyzeBtn.textContent = "Analyzing...";
      }

      const content = textarea.value;

      // Parse the pasted content
      const parsed = this.parseManualEmailContent(content);
      const daysWithoutResponse = this.calculateDaysSince(
        parsed.date || new Date(),
      );

      // Create a follow-up email object
      const email: FollowupEmail = {
        id: `manual-${Date.now()}`,
        subject: parsed.subject || "(No subject)",
        recipients: parsed.to ? [parsed.to] : [],
        sentDate: parsed.date || new Date(),
        body: parsed.body,
        summary:
          parsed.body.substring(0, 200) +
          (parsed.body.length > 200 ? "..." : ""),
        priority: "medium",
        daysWithoutResponse: daysWithoutResponse,
        conversationId: "",
        hasAttachments: false,
        accountEmail:
          Office.context?.mailbox?.userProfile?.emailAddress || "manual",
        threadMessages: [],
        isSnoozed: false,
        isDismissed: false,
      };

      // Calculate priority
      if (daysWithoutResponse >= 7) {
        email.priority = "high";
      } else if (daysWithoutResponse >= 3) {
        email.priority = "medium";
      } else {
        email.priority = "low";
      }

      // If AI is enabled, enhance with LLM
      if (this.llmService && this.enableLlmSummaryCheckbox?.checked) {
        try {
          const llmSummary = await this.llmService.summarizeEmail(parsed.body);
          email.llmSummary = llmSummary;
        } catch (llmError) {
          console.warn("LLM summary failed:", llmError);
        }
      }

      if (this.llmService && this.enableLlmSuggestionsCheckbox?.checked) {
        try {
          const suggestions = await this.llmService.generateFollowupSuggestions(
            parsed.body,
            `Subject: ${email.subject}, Days without response: ${daysWithoutResponse}`,
          );
          email.llmSuggestion = suggestions.join("\n");
        } catch (llmError) {
          console.warn("LLM suggestion failed:", llmError);
        }
      }

      // Display the result
      this.displaySingleEmailResult(email);
      this.showStatus("Email analyzed successfully!", "success");
    } catch (error) {
      this.showStatus(`Analysis failed: ${(error as Error).message}`, "error");
    } finally {
      if (analyzeBtn) {
        analyzeBtn.disabled = false;
        analyzeBtn.textContent = originalText;
      }
    }
  }

  /**
   * Parse manually pasted email content
   */
  private parseManualEmailContent(content: string): {
    subject: string;
    from: string;
    fromName: string;
    to: string;
    date: Date | null;
    body: string;
  } {
    const lines = content.split("\n");
    let subject = "";
    let from = "";
    let fromName = "";
    let to = "";
    let date: Date | null = null;
    let bodyStartIndex = 0;

    // Parse headers
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();

      if (line.toLowerCase().startsWith("subject:")) {
        subject = line.substring(8).trim();
        bodyStartIndex = i + 1;
      } else if (line.toLowerCase().startsWith("from:")) {
        const fromValue = line.substring(5).trim();
        // Try to extract name and email
        const emailMatch = fromValue.match(/<([^>]+)>/);
        if (emailMatch) {
          from = emailMatch[1];
          fromName = fromValue.replace(/<[^>]+>/, "").trim();
        } else {
          from = fromValue;
          fromName = fromValue;
        }
        bodyStartIndex = i + 1;
      } else if (line.toLowerCase().startsWith("to:")) {
        to = line.substring(3).trim();
        bodyStartIndex = i + 1;
      } else if (line.toLowerCase().startsWith("date:")) {
        const dateStr = line.substring(5).trim();
        const parsedDate = new Date(dateStr);
        if (!isNaN(parsedDate.getTime())) {
          date = parsedDate;
        }
        bodyStartIndex = i + 1;
      } else if (line === "" && bodyStartIndex > 0) {
        // Empty line after headers indicates start of body
        bodyStartIndex = i + 1;
        break;
      }
    }

    // Extract body
    const body = lines.slice(bodyStartIndex).join("\n").trim();

    return { subject, from, fromName, to, date, body };
  }

  /**
   * Calculate days since a date
   */
  private calculateDaysSince(date: Date): number {
    const now = new Date();
    const diffTime = Math.abs(now.getTime() - date.getTime());
    return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  }

  /**
   * Display a single email analysis result
   */
  private displaySingleEmailResult(email: FollowupEmail): void {
    // Add to allEmails for action button handlers
    this.allEmails = [email];

    this.hideAllStates();
    this.emailListDiv.style.display = "block";
    this.emailListDiv.innerHTML = "";

    const emailElement = this.createEmailElement(email);
    this.emailListDiv.appendChild(emailElement);
  }

  // Diagnostic modal methods
  private openDiagnosticModal(): void {
    const modal = document.getElementById("diagnosticModal");
    if (modal) {
      modal.style.display = "block";
    }
  }

  private closeDiagnosticModal(): void {
    const modal = document.getElementById("diagnosticModal");
    if (modal) {
      modal.style.display = "none";
    }
  }

  // Diagnostic test methods
  private async testOfficeContext(): Promise<void> {
    const output = document.getElementById("diagnosticOutput");
    if (!output) return;

    try {
      output.innerHTML = "<strong>Testing Office.js Context...</strong><br>";

      // Test basic Office.js availability
      if (typeof Office === "undefined") {
        output.innerHTML += "❌ Office.js is not available<br>";
        return;
      }
      output.innerHTML += "✅ Office.js is available<br>";

      // Test Office.context availability
      if (!Office.context) {
        output.innerHTML += "❌ Office.context is not available<br>";
        output.innerHTML +=
          "<br><strong>Office Context Test Complete</strong><br><br>";
        return;
      }
      output.innerHTML += "✅ Office.context is available<br>";

      // Test host type
      if (Office.context.host) {
        output.innerHTML += `✅ Host: ${Office.context.host}<br>`;
      } else {
        output.innerHTML += "⚠️ Office.context.host not available<br>";
      }

      // Test platform
      if (Office.context.platform) {
        output.innerHTML += `✅ Platform: ${Office.context.platform}<br>`;
      } else {
        output.innerHTML += "⚠️ Office.context.platform not available<br>";
      }

      // Test mailbox availability
      if (Office.context.mailbox) {
        output.innerHTML += "✅ Mailbox context available<br>";
      } else {
        output.innerHTML += "❌ Mailbox context not available<br>";
      }

      // Test roaming settings availability
      if (Office.context.roamingSettings) {
        output.innerHTML += "✅ Roaming settings available<br>";
      } else {
        output.innerHTML += "⚠️ Roaming settings not available<br>";
      }

      // Test content language
      if (Office.context.contentLanguage) {
        output.innerHTML += `✅ Content language: ${Office.context.contentLanguage}<br>`;
      }

      // Test display language
      if (Office.context.displayLanguage) {
        output.innerHTML += `✅ Display language: ${Office.context.displayLanguage}<br>`;
      }

      // Test mailbox item type if available
      if (
        Office.context.mailbox &&
        Office.context.mailbox.item &&
        Office.context.mailbox.item.itemType
      ) {
        const itemType = Office.context.mailbox.item.itemType;
        output.innerHTML += `✅ Current item type: ${itemType}<br>`;
      }

      output.innerHTML +=
        "<br><strong>Office Context Test Complete</strong><br><br>";
    } catch (error) {
      output.innerHTML += `❌ Error: ${(error as Error).message}<br>`;
    }
  }

  private async testMailboxAccess(): Promise<void> {
    const output = document.getElementById("diagnosticOutput");
    if (!output) return;

    try {
      output.innerHTML += "<strong>Testing Mailbox Access...</strong><br>";

      if (!Office.context.mailbox) {
        output.innerHTML += "❌ Mailbox not available<br>";
        return;
      }

      // Test user profile
      if (Office.context.mailbox.userProfile) {
        const profile = Office.context.mailbox.userProfile;
        output.innerHTML += `✅ User email: ${this.escapeHtml(profile.emailAddress || "N/A")}<br>`;
        output.innerHTML += `✅ Display name: ${this.escapeHtml(profile.displayName || "N/A")}<br>`;
        output.innerHTML += `✅ Time zone: ${this.escapeHtml(profile.timeZone || "N/A")}<br>`;
      } else {
        output.innerHTML += "❌ User profile not available<br>";
      }

      // Test diagnostics
      if (Office.context.mailbox.diagnostics) {
        const diag = Office.context.mailbox.diagnostics;
        output.innerHTML += `✅ Host name: ${this.escapeHtml(diag.hostName || "N/A")}<br>`;
        output.innerHTML += `✅ Host version: ${this.escapeHtml(diag.hostVersion || "N/A")}<br>`;
        // OWAView is only available in Outlook on the web, not in desktop Outlook
        const owaView = diag.OWAView
          ? diag.OWAView
          : diag.hostName === "Outlook"
            ? "N/A (Desktop Outlook)"
            : "N/A";
        output.innerHTML += `✅ OWA view: ${this.escapeHtml(owaView)}<br>`;
      } else {
        output.innerHTML += "❌ Diagnostics not available<br>";
      }

      output.innerHTML +=
        "<br><strong>Mailbox Access Test Complete</strong><br><br>";
    } catch (error) {
      output.innerHTML += `❌ Error: ${this.escapeHtml((error as Error).message)}<br>`;
    }
  }

  private async testAccountDetection(): Promise<void> {
    const output = document.getElementById("diagnosticOutput");
    if (!output) return;

    const button = document.getElementById("testAccountDetection");
    const originalButtonText = button?.textContent || "Test Account Detection";

    try {
      // Clear previous results and show initial status
      output.innerHTML = "<strong>🔍 Testing Account Detection...</strong><br>";
      output.innerHTML += "⏳ This may take a few seconds, please wait...<br>";
      output.innerHTML +=
        "<div id='accountDetectionSpinner' style='display: inline-block; margin-left: 10px; color: #0078d4;'>";
      output.innerHTML +=
        "<span class='spinner'>⟳</span> <span id='spinnerText'>Checking Office context...</span></div><br><br>";

      // Disable button and show loading state
      if (button) {
        (button as HTMLButtonElement).disabled = true;
        (button as HTMLButtonElement).textContent = "Testing...";
      }

      // Check if Office.context is available
      if (!Office.context || !Office.context.mailbox) {
        output.innerHTML =
          "<strong>🔍 Testing Account Detection...</strong><br>";
        output.innerHTML += "❌ Office.context.mailbox is not available<br>";
        output.innerHTML +=
          "ℹ️ This is expected when testing in a browser. The add-in needs to run inside Outlook.<br>";
        output.innerHTML +=
          "<br><strong>Account Detection Test Complete</strong><br><br>";
        if (button) {
          (button as HTMLButtonElement).disabled = false;
          (button as HTMLButtonElement).textContent = originalButtonText;
        }
        return;
      }

      // Update spinner text
      const spinnerText = document.getElementById("spinnerText");
      if (spinnerText)
        spinnerText.textContent = "Getting user identity token...";

      // Try to get available accounts using our service with timeout
      try {
        // First, try to get user profile info directly
        let userEmail: string | undefined;
        let userDisplayName: string | undefined;
        try {
          if (Office.context.mailbox.userProfile) {
            userEmail = Office.context.mailbox.userProfile.emailAddress;
            userDisplayName = Office.context.mailbox.userProfile.displayName;
          }
        } catch (profileError) {
          console.warn("Error accessing user profile:", profileError);
        }

        // Try to get identity token with detailed error handling
        const identityTokenPromise = new Promise<{
          success: boolean;
          error?: string;
          errorCode?: string;
        }>((resolve) => {
          try {
            Office.context.mailbox.getUserIdentityTokenAsync((result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve({ success: true });
              } else {
                const error = result.error;
                resolve({
                  success: false,
                  error: error?.message || "Unknown error",
                  errorCode: (error as any)?.code,
                });
              }
            });
          } catch (error) {
            resolve({
              success: false,
              error: (error as Error).message,
            });
          }
        });

        const timeoutPromise = new Promise<{
          success: boolean;
          error?: string;
        }>((_, reject) => {
          setTimeout(
            () => reject(new Error("Request timed out after 10 seconds")),
            10000,
          );
        });

        const identityResult = await Promise.race([
          identityTokenPromise,
          timeoutPromise,
        ]);

        // Update spinner text
        if (spinnerText) spinnerText.textContent = "Processing results...";

        // Display results
        output.innerHTML =
          "<strong>🔍 Testing Account Detection...</strong><br>";

        if (userEmail) {
          output.innerHTML += `✅ User Profile Email: ${this.escapeHtml(userEmail)}<br>`;
          if (userDisplayName) {
            output.innerHTML += `✅ Display Name: ${this.escapeHtml(userDisplayName)}<br>`;
          }
        } else {
          output.innerHTML += "⚠️ Could not access user profile email<br>";
        }

        if (identityResult.success) {
          output.innerHTML +=
            "✅ User identity token obtained successfully<br>";
          if (userEmail) {
            output.innerHTML += `✅ Account detected: ${this.escapeHtml(userEmail)}<br>`;
          }
        } else {
          output.innerHTML += "⚠️ No accounts detected via identity token<br>";
          if (identityResult.error) {
            output.innerHTML += `&nbsp;&nbsp;Error: ${identityResult.error}<br>`;
          }
          if ("errorCode" in identityResult && identityResult.errorCode) {
            output.innerHTML += `&nbsp;&nbsp;Error Code: ${identityResult.errorCode}<br>`;
          }
          output.innerHTML += "ℹ️ This might be normal if:<br>";
          output.innerHTML +=
            "&nbsp;&nbsp;• Office.context is not fully initialized<br>";
          output.innerHTML +=
            "&nbsp;&nbsp;• Add-in permissions don't include identity token access<br>";
          output.innerHTML +=
            "&nbsp;&nbsp;• Running in a restricted environment<br>";
        }
      } catch (error) {
        output.innerHTML =
          "<strong>🔍 Testing Account Detection...</strong><br>";
        const errorMsg = (error as Error).message;
        if (errorMsg.includes("timed out")) {
          output.innerHTML += `⏱️ ${errorMsg}<br>`;
          output.innerHTML +=
            "ℹ️ The request is taking longer than expected. This might indicate:<br>";
          output.innerHTML += "&nbsp;&nbsp;• Network connectivity issues<br>";
          output.innerHTML +=
            "&nbsp;&nbsp;• Office.js is not fully initialized<br>";
          output.innerHTML +=
            "&nbsp;&nbsp;• Running outside of Outlook context<br>";
          output.innerHTML +=
            "&nbsp;&nbsp;• Exchange server is slow to respond<br>";
        } else {
          output.innerHTML += `❌ Account detection error: ${errorMsg}<br>`;
        }
      }

      // Remove spinner
      const spinner = document.getElementById("accountDetectionSpinner");
      if (spinner) spinner.remove();

      // Test EWS availability (this is often the issue on macOS)
      output.innerHTML += "<br>📡 Testing EWS (Exchange Web Services)...<br>";

      if (Office.context.mailbox.makeEwsRequestAsync) {
        output.innerHTML += "✅ EWS API is available<br>";
        output.innerHTML +=
          "⏳ Testing EWS connection (this may take a few seconds)...<br>";

        // Test a simple EWS request with timeout
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

          const ewsPromise = new Promise<void>((resolve, reject) => {
            const timeout = setTimeout(() => {
              reject(new Error("EWS request timed out after 15 seconds"));
            }, 15000);

            Office.context.mailbox.makeEwsRequestAsync(
              simpleEwsRequest,
              (result) => {
                clearTimeout(timeout);
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                  output.innerHTML += "✅ EWS test request successful<br>";
                  resolve();
                } else {
                  // Extract detailed error information
                  const error = result.error;
                  let errorDetails = `❌ EWS test failed: ${error.message || "The operation failed"}<br>`;

                  // Add error code if available
                  if ((error as any).code) {
                    errorDetails += `&nbsp;&nbsp;Error Code: ${(error as any).code}<br>`;
                  }

                  // Add error name if available
                  if ((error as any).name) {
                    errorDetails += `&nbsp;&nbsp;Error Name: ${(error as any).name}<br>`;
                  }

                  // Try to parse response value if available (sometimes EWS returns error details in response)
                  if (result.value && typeof result.value === "string") {
                    try {
                      const parser = new DOMParser();
                      const xmlDoc = parser.parseFromString(
                        result.value,
                        "text/xml",
                      );

                      // Check for SOAP faults
                      const soapFault = xmlDoc.getElementsByTagNameNS(
                        "http://schemas.xmlsoap.org/soap/envelope/",
                        "Fault",
                      )[0];
                      if (soapFault) {
                        const faultString =
                          soapFault.querySelector("faultstring")?.textContent ||
                          soapFault.getElementsByTagName("faultstring")[0]
                            ?.textContent ||
                          "Unknown SOAP fault";
                        const faultCode =
                          soapFault.querySelector("faultcode")?.textContent ||
                          soapFault.getElementsByTagName("faultcode")[0]
                            ?.textContent;
                        errorDetails += `&nbsp;&nbsp;SOAP Fault: ${faultString}<br>`;
                        if (faultCode) {
                          errorDetails += `&nbsp;&nbsp;Fault Code: ${faultCode}<br>`;
                        }
                      }

                      // Check for EWS response errors
                      const responseMessages =
                        xmlDoc.getElementsByTagNameNS(
                          "http://schemas.microsoft.com/exchange/services/2006/messages",
                          "ResponseMessages",
                        )[0] || xmlDoc.querySelector("ResponseMessages");

                      if (responseMessages) {
                        const errorResponse =
                          responseMessages.querySelector(
                            '[ResponseClass="Error"]',
                          ) ||
                          Array.from(
                            responseMessages.querySelectorAll("*"),
                          ).find(
                            (el) =>
                              el.getAttribute("ResponseClass") === "Error",
                          );
                        if (errorResponse) {
                          const responseCode =
                            errorResponse.querySelector("ResponseCode")
                              ?.textContent ||
                            Array.from(
                              errorResponse.querySelectorAll("*"),
                            ).find(
                              (el) =>
                                el.tagName === "ResponseCode" ||
                                el.localName === "ResponseCode",
                            )?.textContent ||
                            "Unknown error";
                          const messageText =
                            errorResponse.querySelector("MessageText")
                              ?.textContent ||
                            Array.from(
                              errorResponse.querySelectorAll("*"),
                            ).find(
                              (el) =>
                                el.tagName === "MessageText" ||
                                el.localName === "MessageText",
                            )?.textContent ||
                            "";
                          errorDetails += `&nbsp;&nbsp;EWS Error Code: ${responseCode}<br>`;
                          if (messageText) {
                            errorDetails += `&nbsp;&nbsp;EWS Message: ${messageText}<br>`;
                          }
                        }
                      }
                    } catch (parseError) {
                      // If parsing fails, just show the raw error message
                      console.warn(
                        "Failed to parse EWS error response:",
                        parseError,
                      );
                    }
                  }

                  output.innerHTML += errorDetails;

                  // Add troubleshooting information
                  output.innerHTML +=
                    "<br>ℹ️ <strong>Troubleshooting:</strong><br>";
                  output.innerHTML +=
                    "&nbsp;&nbsp;• Verify EWS is enabled on your Exchange server<br>";
                  output.innerHTML +=
                    "&nbsp;&nbsp;• Check if you have proper permissions for EWS access<br>";
                  output.innerHTML +=
                    "&nbsp;&nbsp;• Ensure you're connected to Exchange (not just IMAP/POP3)<br>";
                  output.innerHTML +=
                    "&nbsp;&nbsp;• Some organizations restrict EWS access for security reasons<br>";

                  reject(new Error(error.message || "The operation failed"));
                }
              },
            );
          });

          await ewsPromise;
        } catch (ewsError) {
          const errorMsg = (ewsError as Error).message;
          if (errorMsg.includes("timed out")) {
            output.innerHTML += `⏱️ ${errorMsg}<br>`;
            output.innerHTML +=
              "ℹ️ EWS requests can take time. This might be normal.<br>";
            output.innerHTML +=
              "&nbsp;&nbsp;• Check your network connection<br>";
            output.innerHTML +=
              "&nbsp;&nbsp;• The Exchange server might be slow to respond<br>";
          } else if (!errorMsg.includes("EWS test failed")) {
            // Only show this if we haven't already shown detailed error above
            output.innerHTML += `❌ EWS test request failed: ${errorMsg}<br>`;
          }
        }
      } else {
        output.innerHTML +=
          "❌ EWS not available (this is common on macOS or when running outside Outlook)<br>";
      }

      output.innerHTML +=
        "<br><strong>✅ Account Detection Test Complete</strong><br><br>";
    } catch (error) {
      output.innerHTML += `❌ Unexpected error: ${(error as Error).message}<br>`;
    } finally {
      // Re-enable button
      if (button) {
        (button as HTMLButtonElement).disabled = false;
        (button as HTMLButtonElement).textContent = originalButtonText;
      }
    }
  }

  private async testEmailReading(): Promise<void> {
    const output = document.getElementById("diagnosticOutput");
    if (!output) return;

    try {
      output.innerHTML += "<strong>Testing Email Reading...</strong><br>";

      // Try to read emails using our email analysis service
      try {
        output.innerHTML +=
          "Attempting to read 5 emails from last 7 days...<br>";
        const emails = await this.emailAnalysisService.analyzeEmails(5, 7, []);

        if (emails.length > 0) {
          output.innerHTML += `✅ Successfully read ${emails.length} email(s)<br>`;
          emails.forEach((email, index) => {
            output.innerHTML += `&nbsp;&nbsp;${index + 1}. ${this.escapeHtml(email.subject)} (${email.sentDate.toLocaleDateString()})<br>`;
          });
        } else {
          output.innerHTML +=
            "⚠️ No emails found (this could be normal if no sent emails in timeframe)<br>";
        }
      } catch (emailError) {
        output.innerHTML += `❌ Email reading failed: ${(emailError as Error).message}<br>`;

        // Provide specific guidance based on error
        const errorMsg = (emailError as Error).message.toLowerCase();
        if (errorMsg.includes("ews") || errorMsg.includes("exchange")) {
          output.innerHTML +=
            "💡 This appears to be an EWS (Exchange) issue. On macOS, EWS is often limited.<br>";
          output.innerHTML +=
            "💡 Consider switching to REST API for better macOS compatibility.<br>";
        } else if (
          errorMsg.includes("permission") ||
          errorMsg.includes("unauthorized")
        ) {
          output.innerHTML +=
            "💡 This appears to be a permissions issue. Check manifest permissions.<br>";
        } else if (
          errorMsg.includes("network") ||
          errorMsg.includes("timeout")
        ) {
          output.innerHTML +=
            "💡 This appears to be a network connectivity issue.<br>";
        }
      }

      output.innerHTML +=
        "<br><strong>Email Reading Test Complete</strong><br><br>";
    } catch (error) {
      output.innerHTML += `❌ Error: ${(error as Error).message}<br>`;
    }
  }

  private async testGraphApi(): Promise<void> {
    const output = document.getElementById("diagnosticOutput");
    if (!output) return;

    const button = document.getElementById("testGraphApi");
    const originalButtonText = button?.textContent || "Test Graph API";

    try {
      // Disable button and show loading state
      if (button) {
        (button as HTMLButtonElement).disabled = true;
        (button as HTMLButtonElement).textContent = "Testing...";
      }

      output.innerHTML =
        "<strong>🔗 Testing Microsoft Graph API...</strong><br>";
      output.innerHTML +=
        "⏳ This test checks if Graph API can be used as an alternative to EWS...<br><br>";

      // Step 1: Check if Office.auth is available
      output.innerHTML +=
        "<strong>Step 1: Checking Office.auth availability...</strong><br>";

      if (typeof Office === "undefined" || !Office.context) {
        output.innerHTML += "❌ Office.context is not available<br>";
        output.innerHTML +=
          "<br><strong>Graph API Test Complete</strong><br><br>";
        return;
      }

      // Check if Office.auth exists (SSO capability)
      const hasAuth = typeof (Office as any).auth !== "undefined";
      if (hasAuth) {
        output.innerHTML += "✅ Office.auth is available (SSO supported)<br>";
      } else {
        output.innerHTML += "⚠️ Office.auth is not available<br>";
        output.innerHTML +=
          "ℹ️ SSO requires manifest configuration and Azure AD app registration<br>";
      }

      // Step 2: Try to get access token
      output.innerHTML +=
        "<br><strong>Step 2: Attempting to get access token...</strong><br>";

      if (hasAuth && (Office as any).auth.getAccessToken) {
        try {
          const tokenOptions = {
            allowSignInPrompt: false,
            allowConsentPrompt: false,
            forMSGraphAccess: true,
          };

          const tokenPromise = (Office as any).auth.getAccessToken(
            tokenOptions,
          );
          const timeoutPromise = new Promise<never>((_, reject) => {
            setTimeout(
              () =>
                reject(new Error("Token request timed out after 10 seconds")),
              10000,
            );
          });

          const token = await Promise.race([tokenPromise, timeoutPromise]);

          if (token) {
            output.innerHTML += "✅ Access token obtained successfully<br>";
            output.innerHTML += `&nbsp;&nbsp;Token length: ${token.length} characters<br>`;

            // Try to decode token to get some info (JWT has 3 parts separated by .)
            const parts = token.split(".");
            if (parts.length === 3) {
              try {
                const payload = JSON.parse(atob(parts[1]));
                if (payload.aud) {
                  output.innerHTML += `&nbsp;&nbsp;Audience: ${payload.aud}<br>`;
                }
                if (payload.scp) {
                  output.innerHTML += `&nbsp;&nbsp;Scopes: ${payload.scp}<br>`;
                }
                if (payload.upn || payload.unique_name) {
                  output.innerHTML += `&nbsp;&nbsp;User: ${payload.upn || payload.unique_name}<br>`;
                }
              } catch {
                // Token parsing failed, that's okay
              }
            }

            // Step 3: Try a simple Graph API call
            output.innerHTML +=
              "<br><strong>Step 3: Testing Graph API call...</strong><br>";
            output.innerHTML += "⏳ Calling /me endpoint...<br>";

            try {
              const response = await fetch(
                "https://graph.microsoft.com/v1.0/me",
                {
                  headers: {
                    Authorization: `Bearer ${token}`,
                    "Content-Type": "application/json",
                  },
                },
              );

              if (response.ok) {
                const userData = await response.json();
                output.innerHTML += "✅ Graph API /me call successful<br>";
                output.innerHTML += `&nbsp;&nbsp;Display Name: ${this.escapeHtml(userData.displayName || "N/A")}<br>`;
                output.innerHTML += `&nbsp;&nbsp;Email: ${this.escapeHtml(userData.mail || userData.userPrincipalName || "N/A")}<br>`;

                // Step 4: Try to get mail folders
                output.innerHTML +=
                  "<br><strong>Step 4: Testing mail access...</strong><br>";
                output.innerHTML += "⏳ Calling /me/mailFolders...<br>";

                try {
                  const foldersResponse = await fetch(
                    "https://graph.microsoft.com/v1.0/me/mailFolders?$top=5",
                    {
                      headers: {
                        Authorization: `Bearer ${token}`,
                        "Content-Type": "application/json",
                      },
                    },
                  );

                  if (foldersResponse.ok) {
                    const foldersData = await foldersResponse.json();
                    output.innerHTML += "✅ Mail folders accessible<br>";
                    if (foldersData.value && foldersData.value.length > 0) {
                      output.innerHTML += "&nbsp;&nbsp;Folders found:<br>";
                      foldersData.value.slice(0, 5).forEach((folder: any) => {
                        output.innerHTML += `&nbsp;&nbsp;&nbsp;&nbsp;• ${this.escapeHtml(folder.displayName)} (${folder.totalItemCount || 0} items)<br>`;
                      });
                    }

                    // Step 5: Try to get sent items
                    output.innerHTML +=
                      "<br><strong>Step 5: Testing sent items access...</strong><br>";
                    output.innerHTML += "⏳ Getting recent sent emails...<br>";

                    try {
                      const sentResponse = await fetch(
                        "https://graph.microsoft.com/v1.0/me/mailFolders/sentItems/messages?$top=3&$select=subject,sentDateTime,toRecipients",
                        {
                          headers: {
                            Authorization: `Bearer ${token}`,
                            "Content-Type": "application/json",
                          },
                        },
                      );

                      if (sentResponse.ok) {
                        const sentData = await sentResponse.json();
                        output.innerHTML += "✅ Sent items accessible<br>";
                        if (sentData.value && sentData.value.length > 0) {
                          output.innerHTML +=
                            "&nbsp;&nbsp;Recent sent emails:<br>";
                          sentData.value.forEach(
                            (email: any, index: number) => {
                              const sentDate = new Date(
                                email.sentDateTime,
                              ).toLocaleDateString();
                              output.innerHTML += `&nbsp;&nbsp;&nbsp;&nbsp;${index + 1}. ${this.escapeHtml(email.subject || "(No subject)")} (${sentDate})<br>`;
                            },
                          );
                          output.innerHTML +=
                            "<br>🎉 <strong>Graph API can substitute EWS!</strong><br>";
                          output.innerHTML +=
                            "ℹ️ The add-in can be updated to use Microsoft Graph API instead of EWS.<br>";
                        } else {
                          output.innerHTML += "⚠️ No sent emails found<br>";
                        }
                      } else {
                        const errorText = await sentResponse.text();
                        output.innerHTML += `❌ Sent items access failed: ${sentResponse.status}<br>`;
                        output.innerHTML += `&nbsp;&nbsp;Error: ${errorText.substring(0, 200)}<br>`;
                        output.innerHTML +=
                          "ℹ️ Mail.Read permission may be missing from the app registration<br>";
                      }
                    } catch (sentError) {
                      output.innerHTML += `❌ Sent items request failed: ${(sentError as Error).message}<br>`;
                    }
                  } else {
                    const errorText = await foldersResponse.text();
                    output.innerHTML += `❌ Mail folders access failed: ${foldersResponse.status}<br>`;
                    output.innerHTML += `&nbsp;&nbsp;Error: ${errorText.substring(0, 200)}<br>`;
                    output.innerHTML +=
                      "ℹ️ Mail.Read permission may be missing from the app registration<br>";
                  }
                } catch (foldersError) {
                  output.innerHTML += `❌ Mail folders request failed: ${(foldersError as Error).message}<br>`;
                }
              } else {
                const errorText = await response.text();
                output.innerHTML += `❌ Graph API /me call failed: ${response.status}<br>`;
                output.innerHTML += `&nbsp;&nbsp;Error: ${errorText.substring(0, 200)}<br>`;

                if (response.status === 401) {
                  output.innerHTML +=
                    "ℹ️ The token may need to be exchanged for a Graph token (on-behalf-of flow)<br>";
                  output.innerHTML +=
                    "ℹ️ This requires a server-side component to exchange the token<br>";
                }
              }
            } catch (graphError) {
              output.innerHTML += `❌ Graph API call failed: ${(graphError as Error).message}<br>`;
              if ((graphError as Error).message.includes("Failed to fetch")) {
                output.innerHTML +=
                  "ℹ️ Network request blocked. CORS or network policy may be blocking the request.<br>";
              }
            }
          } else {
            output.innerHTML += "⚠️ Empty token received<br>";
          }
        } catch (tokenError) {
          const errorMsg = (tokenError as any).message || String(tokenError);
          const errorCode = (tokenError as any).code;

          output.innerHTML += `❌ Token acquisition failed<br>`;
          output.innerHTML += `&nbsp;&nbsp;Error: ${errorMsg}<br>`;
          if (errorCode) {
            output.innerHTML += `&nbsp;&nbsp;Code: ${errorCode}<br>`;
          }

          // Provide guidance based on error
          if (errorCode === 13001 || errorMsg.includes("13001")) {
            output.innerHTML +=
              "<br>ℹ️ Error 13001: SSO is not supported in this environment<br>";
            output.innerHTML +=
              "&nbsp;&nbsp;• Manifest may not have WebApplicationInfo configured<br>";
            output.innerHTML +=
              "&nbsp;&nbsp;• Azure AD app registration may be missing<br>";
          } else if (errorCode === 13002 || errorMsg.includes("13002")) {
            output.innerHTML +=
              "<br>ℹ️ Error 13002: User needs to consent to the app<br>";
            output.innerHTML +=
              "&nbsp;&nbsp;• Try with allowConsentPrompt: true<br>";
          } else if (errorCode === 13003 || errorMsg.includes("13003")) {
            output.innerHTML +=
              "<br>ℹ️ Error 13003: Azure AD app configuration issue<br>";
            output.innerHTML +=
              "&nbsp;&nbsp;• Check the application ID in manifest matches Azure AD<br>";
          } else if (errorCode === 13005 || errorMsg.includes("13005")) {
            output.innerHTML +=
              "<br>ℹ️ Error 13005: Resource issue - check the API permissions<br>";
          } else if (errorCode === 13006 || errorMsg.includes("13006")) {
            output.innerHTML +=
              "<br>ℹ️ Error 13006: SSO failed - user may need to sign in<br>";
          } else if (errorCode === 13007 || errorMsg.includes("13007")) {
            output.innerHTML +=
              "<br>ℹ️ Error 13007: Office is in offline mode<br>";
          } else if (errorCode === 13008 || errorMsg.includes("13008")) {
            output.innerHTML +=
              "<br>ℹ️ Error 13008: Previous SSO operation still in progress<br>";
          } else if (errorCode === 13010 || errorMsg.includes("13010")) {
            output.innerHTML +=
              "<br>ℹ️ Error 13010: Running in Office on the web in Edge Legacy<br>";
          } else if (errorCode === 13012 || errorMsg.includes("13012")) {
            output.innerHTML +=
              "<br>ℹ️ Error 13012: Office version doesn't support SSO<br>";
          }

          output.innerHTML +=
            "<br><strong>📋 To enable Graph API, you need:</strong><br>";
          output.innerHTML +=
            "&nbsp;&nbsp;1. Register an app in Azure AD (Microsoft Entra)<br>";
          output.innerHTML +=
            "&nbsp;&nbsp;2. Add WebApplicationInfo to manifest.xml<br>";
          output.innerHTML +=
            "&nbsp;&nbsp;3. Configure API permissions (Mail.Read, User.Read)<br>";
          output.innerHTML +=
            "&nbsp;&nbsp;4. Admin consent for the organization<br>";
        }
      } else {
        output.innerHTML +=
          "⚠️ Office.auth.getAccessToken is not available<br>";
        output.innerHTML +=
          "ℹ️ SSO requires Azure AD registration (admin access needed)<br>";
      }

      // Step 3 (Alternative): Try callback token approach - NO Azure AD needed!
      output.innerHTML +=
        "<br><strong>Step 3 (Alternative): Testing Callback Token (No Azure AD needed)...</strong><br>";
      output.innerHTML +=
        "ℹ️ This approach uses existing add-in permissions without Azure AD registration<br>";

      await this.testCallbackTokenApproach(output);

      output.innerHTML +=
        "<br><strong>✅ Graph API Test Complete</strong><br><br>";
    } catch (error) {
      output.innerHTML += `❌ Unexpected error: ${(error as Error).message}<br>`;
    } finally {
      // Re-enable button
      if (button) {
        (button as HTMLButtonElement).disabled = false;
        (button as HTMLButtonElement).textContent = originalButtonText;
      }
    }
  }

  /**
   * Test callback token approach for REST API access
   * This method doesn't require Azure AD registration!
   */
  private async testCallbackTokenApproach(
    output: HTMLElement,
  ): Promise<boolean> {
    try {
      if (
        !Office.context.mailbox ||
        !Office.context.mailbox.getCallbackTokenAsync
      ) {
        output.innerHTML +=
          "❌ getCallbackTokenAsync not available on this platform<br>";
        return false;
      }

      output.innerHTML += "⏳ Getting callback token (REST mode)...<br>";

      // Try to get a REST callback token first
      let tokenResult = await new Promise<{
        success: boolean;
        token?: string;
        error?: string;
        errorCode?: number;
        restUrl?: string;
        mode?: string;
      }>((resolve) => {
        const timeout = setTimeout(() => {
          resolve({ success: false, error: "Token request timed out" });
        }, 15000);

        Office.context.mailbox.getCallbackTokenAsync(
          { isRest: true },
          (result) => {
            clearTimeout(timeout);
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve({
                success: true,
                token: result.value,
                restUrl: (Office.context.mailbox as any).restUrl,
                mode: "REST",
              });
            } else {
              resolve({
                success: false,
                error: result.error?.message || "Failed to get token",
                errorCode: (result.error as any)?.code,
                mode: "REST",
              });
            }
          },
        );
      });

      // If REST mode failed, try EWS mode (without isRest flag)
      if (!tokenResult.success) {
        output.innerHTML += `❌ REST callback token failed: ${tokenResult.error}<br>`;
        if (tokenResult.errorCode) {
          output.innerHTML += `&nbsp;&nbsp;Error Code: ${tokenResult.errorCode}<br>`;
        }

        output.innerHTML +=
          "<br>⏳ Trying EWS callback token (legacy mode)...<br>";

        tokenResult = await new Promise<{
          success: boolean;
          token?: string;
          error?: string;
          errorCode?: number;
          restUrl?: string;
          mode?: string;
        }>((resolve) => {
          const timeout = setTimeout(() => {
            resolve({ success: false, error: "Token request timed out" });
          }, 15000);

          // Try without isRest - this gets an EWS token
          Office.context.mailbox.getCallbackTokenAsync((result) => {
            clearTimeout(timeout);
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve({
                success: true,
                token: result.value,
                mode: "EWS",
              });
            } else {
              resolve({
                success: false,
                error: result.error?.message || "Failed to get token",
                errorCode: (result.error as any)?.code,
                mode: "EWS",
              });
            }
          });
        });

        if (!tokenResult.success) {
          output.innerHTML += `❌ EWS callback token also failed: ${tokenResult.error}<br>`;
          if (tokenResult.errorCode) {
            output.innerHTML += `&nbsp;&nbsp;Error Code: ${tokenResult.errorCode}<br>`;
          }

          // Check diagnostics info
          output.innerHTML += "<br>📋 <strong>Environment Info:</strong><br>";
          try {
            const diag = Office.context.mailbox.diagnostics;
            output.innerHTML += `&nbsp;&nbsp;Host: ${diag?.hostName || "Unknown"}<br>`;
            output.innerHTML += `&nbsp;&nbsp;Version: ${diag?.hostVersion || "Unknown"}<br>`;
            output.innerHTML += `&nbsp;&nbsp;OWA View: ${diag?.OWAView || "N/A"}<br>`;
          } catch {
            output.innerHTML += "&nbsp;&nbsp;Could not get diagnostics<br>";
          }

          output.innerHTML +=
            "<br>⚠️ <strong>Both token methods failed.</strong><br>";
          output.innerHTML +=
            "ℹ️ This suggests Exchange/M365 policies are blocking add-in API access.<br>";
          output.innerHTML +=
            "ℹ️ Contact your IT administrator to check if add-in mail access is enabled.<br>";
          return false;
        }
      }

      output.innerHTML += `✅ Callback token obtained successfully (${tokenResult.mode} mode)<br>`;
      output.innerHTML += `&nbsp;&nbsp;Token length: ${tokenResult.token?.length} characters<br>`;

      // Check for REST URL - only use REST URL if we got a REST token
      let apiUrl: string;
      if (tokenResult.mode === "REST") {
        apiUrl = tokenResult.restUrl || "https://outlook.office.com/api/v2.0";
        output.innerHTML += `&nbsp;&nbsp;REST URL: ${apiUrl}<br>`;
      } else {
        // EWS token - we can't use REST API with it
        output.innerHTML +=
          "<br>⚠️ EWS token obtained, but this requires EWS API (not REST)<br>";
        output.innerHTML +=
          "ℹ️ EWS token can be used with makeEwsRequestAsync but that's already failing.<br>";
        output.innerHTML +=
          "ℹ️ The EWS token approach doesn't help if EWS is blocked.<br>";
        return false;
      }

      const restUrl = apiUrl;

      // Try to call the REST API
      output.innerHTML +=
        "<br>⏳ Testing REST API call to get sent items...<br>";

      try {
        // Build the REST URL for sent items
        const sentItemsUrl = `${restUrl}/me/mailfolders/sentitems/messages?$top=3&$select=Subject,DateTimeSent,ToRecipients`;

        const response = await fetch(sentItemsUrl, {
          headers: {
            Authorization: `Bearer ${tokenResult.token}`,
            Accept: "application/json",
            "Content-Type": "application/json",
          },
        });

        if (response.ok) {
          const data = await response.json();
          output.innerHTML += "✅ REST API call successful!<br>";

          if (data.value && data.value.length > 0) {
            output.innerHTML += "&nbsp;&nbsp;Recent sent emails:<br>";
            data.value.forEach((email: any, index: number) => {
              const sentDate = email.DateTimeSent
                ? new Date(email.DateTimeSent).toLocaleDateString()
                : "Unknown";
              output.innerHTML += `&nbsp;&nbsp;&nbsp;&nbsp;${index + 1}. ${this.escapeHtml(email.Subject || "(No subject)")} (${sentDate})<br>`;
            });

            output.innerHTML +=
              "<br>🎉 <strong style='color: green;'>SUCCESS! Callback Token + REST API works!</strong><br>";
            output.innerHTML +=
              "ℹ️ This can replace EWS without needing Azure AD registration.<br>";
            output.innerHTML +=
              "ℹ️ The add-in can be updated to use REST API with callback tokens.<br>";
            return true;
          } else {
            output.innerHTML += "⚠️ No sent emails found<br>";
            output.innerHTML +=
              "ℹ️ REST API access works, but no emails in timeframe<br>";
            return true;
          }
        } else {
          const errorText = await response.text();
          output.innerHTML += `❌ REST API call failed: ${response.status} ${response.statusText}<br>`;

          // Parse error for more details
          try {
            const errorJson = JSON.parse(errorText);
            if (errorJson.error) {
              output.innerHTML += `&nbsp;&nbsp;Error Code: ${this.escapeHtml(errorJson.error.code || "Unknown")}<br>`;
              output.innerHTML += `&nbsp;&nbsp;Message: ${this.escapeHtml(errorJson.error.message || errorText.substring(0, 200))}<br>`;
            }
          } catch {
            output.innerHTML += `&nbsp;&nbsp;Error: ${this.escapeHtml(errorText.substring(0, 200))}<br>`;
          }

          if (response.status === 401) {
            output.innerHTML +=
              "ℹ️ Token may have expired or lacks required permissions<br>";
          } else if (response.status === 403) {
            output.innerHTML +=
              "ℹ️ Access forbidden - organization may restrict REST API access<br>";
          } else if (response.status === 404) {
            output.innerHTML +=
              "ℹ️ Endpoint not found - try different REST URL format<br>";

            // Try alternative URL format
            output.innerHTML += "<br>⏳ Trying Microsoft Graph endpoint...<br>";
            return await this.tryGraphWithCallbackToken(
              output,
              tokenResult.token!,
            );
          }
          return false;
        }
      } catch (fetchError) {
        output.innerHTML += `❌ REST API request failed: ${(fetchError as Error).message}<br>`;

        if ((fetchError as Error).message.includes("Failed to fetch")) {
          output.innerHTML +=
            "ℹ️ Network request blocked - CORS or network policy issue<br>";
          output.innerHTML +=
            "ℹ️ Try running in Outlook on the web (OWA) which may have fewer restrictions<br>";
        }
        return false;
      }
    } catch (error) {
      output.innerHTML += `❌ Callback token test failed: ${(error as Error).message}<br>`;
      return false;
    }
  }

  /**
   * Try using callback token with Microsoft Graph endpoint
   */
  private async tryGraphWithCallbackToken(
    output: HTMLElement,
    token: string,
  ): Promise<boolean> {
    try {
      const graphUrl =
        "https://graph.microsoft.com/v1.0/me/mailFolders/sentItems/messages?$top=3&$select=subject,sentDateTime,toRecipients";

      const response = await fetch(graphUrl, {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/json",
        },
      });

      if (response.ok) {
        const data = await response.json();
        output.innerHTML += "✅ Microsoft Graph API call successful!<br>";

        if (data.value && data.value.length > 0) {
          output.innerHTML += "&nbsp;&nbsp;Recent sent emails:<br>";
          data.value.forEach((email: any, index: number) => {
            const sentDate = email.sentDateTime
              ? new Date(email.sentDateTime).toLocaleDateString()
              : "Unknown";
            output.innerHTML += `&nbsp;&nbsp;&nbsp;&nbsp;${index + 1}. ${email.subject || "(No subject)"} (${sentDate})<br>`;
          });

          output.innerHTML +=
            "<br>🎉 <strong style='color: green;'>SUCCESS! Graph API with callback token works!</strong><br>";
          return true;
        }
        return true;
      } else {
        const errorText = await response.text();
        output.innerHTML += `❌ Graph API call failed: ${response.status}<br>`;
        output.innerHTML += `&nbsp;&nbsp;${errorText.substring(0, 150)}<br>`;

        if (response.status === 401) {
          output.innerHTML +=
            "ℹ️ Callback tokens typically can't access Microsoft Graph directly<br>";
          output.innerHTML +=
            "ℹ️ They work with Outlook REST API (outlook.office.com) instead<br>";
        }
        return false;
      }
    } catch (error) {
      output.innerHTML += `❌ Graph API test failed: ${(error as Error).message}<br>`;
      return false;
    }
  }

  // AI Service Management Methods
  private updateAiStatus(): void {
    const aiDisabled = localStorage.getItem("aiDisabled") === "true";

    if (aiDisabled) {
      this.setAiStatus("warning", "⚠️ AI features manually disabled");
      this.enableAiFeaturesCheckbox.checked = false;
      return;
    }

    if (
      !this.llmService ||
      !this.llmEndpointInput.value.trim() ||
      !this.llmApiKeyInput.value.trim()
    ) {
      this.setAiStatus(
        "warning",
        "⚠️ AI features disabled - No API configuration",
      );
      return;
    }

    // Check if circuit breaker is open for llm-api
    const circuitStates = this.retryService.getCircuitBreakerStates();
    if (circuitStates["llm-api"] === "OPEN") {
      this.setAiStatus(
        "error",
        "❌ AI service temporarily unavailable - Too many failures detected",
      );
      return;
    }

    this.setAiStatus("success", "✅ AI service configured and ready");
  }

  private toggleAiFeatures(): void {
    const isEnabled = this.enableAiFeaturesCheckbox.checked;
    localStorage.setItem("aiDisabled", (!isEnabled).toString());

    if (!isEnabled) {
      this.setAiStatus("warning", "⚠️ AI features manually disabled");
      this.showStatus(
        "AI features disabled - errors will stop appearing",
        "success",
      );
    } else {
      this.updateAiStatus();
      this.showStatus("AI features re-enabled", "success");
    }
  }

  private disableAiFeatures(): void {
    localStorage.setItem("aiDisabled", "true");
    this.enableAiFeaturesCheckbox.checked = false;
    this.setAiStatus("warning", "⚠️ AI features manually disabled");
    this.showStatus(
      "AI features disabled - no more AI-related errors will appear",
      "success",
    );
  }

  private handleProviderChange(): void {
    const provider = this.llmProviderSelect.value;
    const isAzure =
      provider === "azure" ||
      (provider === "" &&
        this.llmEndpointInput.value.includes("openai.azure.com"));

    this.azureSpecificOptions.style.display = isAzure ? "block" : "none";

    // Set default values based on provider
    if (provider === "dial" && !this.llmEndpointInput.value) {
      this.llmEndpointInput.value =
        "https://ai-proxy.lab.epam.com/openai/chat/completions";
      this.llmModelInput.value = "gpt-35-turbo";
    } else if (provider === "azure" && !this.llmEndpointInput.value) {
      this.llmEndpointInput.placeholder =
        "https://your-resource.openai.azure.com";
      this.llmModelInput.value = "gpt-35-turbo";
      this.llmApiVersionInput.value = "2023-12-01-preview";
    } else if (provider === "openai" && !this.llmEndpointInput.value) {
      this.llmEndpointInput.value =
        "https://api.openai.com/v1/chat/completions";
      this.llmModelInput.value = "gpt-3.5-turbo";
    }
  }

  private setAiStatus(
    type: "success" | "warning" | "error",
    message: string,
  ): void {
    this.aiStatusDiv.style.display = "block";
    this.aiStatusDiv.className = `ai-status ${type}`;
    this.aiStatusText.textContent = message;
  }

  private async testAiConnection(): Promise<void> {
    if (
      !this.llmEndpointInput.value.trim() ||
      !this.llmApiKeyInput.value.trim()
    ) {
      this.setAiStatus(
        "warning",
        "⚠️ Please enter both API endpoint and API key",
      );
      return;
    }

    this.setAiStatus("warning", "🔄 Testing AI connection...");
    this.testAiConnectionButton.disabled = true;
    this.testAiConnectionButton.textContent = "Testing...";

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
        llmModel: this.llmModelInput.value.trim() || "gpt-35-turbo",
        llmProvider:
          (this.llmProviderSelect.value as "azure" | "dial" | "openai") ||
          undefined,
        llmDeploymentName: this.llmDeploymentNameInput.value.trim(),
        llmApiVersion:
          this.llmApiVersionInput.value.trim() || "2023-12-01-preview",
        showSnoozedEmails: false,
        showDismissedEmails: false,
        selectedAccounts: [],
        snoozeOptions: [],
      };

      // Create a temporary LLM service for testing
      const testLlmService = new LlmService(testConfig, this.retryService);

      // Test with a simple prompt
      const testPrompt =
        "Test connection. Please respond with 'Connection successful'.";
      const response =
        await testLlmService.generateFollowupSuggestions(testPrompt);

      if (response && response.length > 0) {
        this.setAiStatus("success", "✅ AI connection test successful!");
      } else {
        this.setAiStatus(
          "error",
          "❌ AI service responded but with empty result",
        );
      }
    } catch (error) {
      const errorMessage = (error as Error).message;
      if (errorMessage.includes("502")) {
        this.setAiStatus(
          "error",
          "❌ Connection failed: AI service not available (502 error)",
        );
      } else if (errorMessage.includes("403") || errorMessage.includes("401")) {
        this.setAiStatus(
          "error",
          "❌ Connection failed: Invalid API key or unauthorized",
        );
      } else if (errorMessage.includes("429")) {
        this.setAiStatus("error", "❌ Connection failed: Rate limit exceeded");
      } else {
        this.setAiStatus("error", `❌ Connection failed: ${errorMessage}`);
      }
    } finally {
      this.testAiConnectionButton.disabled = false;
      this.testAiConnectionButton.textContent = "Test AI Connection";
    }
  }
}

// Initialize the task pane when Office is ready (skip in test/non-DOM environments)
let initialized = false;

function initializeTaskpane() {
  if (initialized) return;

  if (
    typeof document !== "undefined" &&
    document.getElementById("analyzeButton")
  ) {
    try {
      initialized = true;
      console.log("[Followup Suggester] Initializing task pane...");
      const taskpaneManager = new TaskpaneManager();
      taskpaneManager.initialize().catch(console.error);
    } catch (e) {
      console.warn("Taskpane initialization failed:", e);
      initialized = false;
    }
  }
}

if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
  Office.onReady((info) => {
    try {
      // Initialize if running in Outlook host and expected root element exists
      if (info.host === Office.HostType.Outlook) {
        initializeTaskpane();
      } else {
        // Office.js loaded but not in Outlook - initialize anyway for development
        console.log(
          "[Followup Suggester] Office.js loaded but not in Outlook host, initializing anyway for development",
        );
        initializeTaskpane();
      }
    } catch (e) {
      // Swallow errors in headless/unit test environments
      console.warn("Taskpane auto-initialization skipped:", e);
      // Try fallback initialization
      initializeTaskpane();
    }
  });
} else {
  // Fallback: Initialize in browser environment for development/testing
  // This allows the UI to work even when Office.js is not fully available
  if (typeof document !== "undefined") {
    // Wait for DOM to be ready
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", initializeTaskpane);
    } else {
      // DOM already loaded
      initializeTaskpane();
    }
  }
}
