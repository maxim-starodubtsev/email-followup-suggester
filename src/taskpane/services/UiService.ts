import { FollowupEmail } from "../../models/FollowupEmail";
import { SnoozeOption } from "../../models/Configuration";

export class UiService {
  // Main controls
  public analyzeButton!: HTMLButtonElement;
  public refreshButton!: HTMLButtonElement;
  public settingsButton!: HTMLButtonElement;
  public emailCountSelect!: HTMLSelectElement;
  public daysBackSelect!: HTMLSelectElement;
  public accountFilterSelect!: HTMLSelectElement;
  public enableLlmSummaryCheckbox!: HTMLInputElement;
  public enableLlmSuggestionsCheckbox!: HTMLInputElement;

  // Display elements
  private statusDiv!: HTMLDivElement;
  private loadingDiv!: HTMLDivElement;
  private emptyStateDiv!: HTMLDivElement;
  private emailListDiv!: HTMLDivElement;

  // Modal elements
  public snoozeModal!: HTMLDivElement;
  public settingsModal!: HTMLDivElement;
  public snoozeOptionsSelect!: HTMLSelectElement;
  public customSnoozeGroup!: HTMLDivElement;
  public customSnoozeDate!: HTMLInputElement;
  public llmEndpointInput!: HTMLInputElement;
  public llmApiKeyInput!: HTMLInputElement;
  public showSnoozedEmailsCheckbox!: HTMLInputElement;
  public showDismissedEmailsCheckbox!: HTMLInputElement;
  public aiStatusDiv!: HTMLDivElement;
  public aiStatusText!: HTMLSpanElement;
  public testAiConnectionButton!: HTMLButtonElement;
  public disableAiFeaturesButton!: HTMLButtonElement;
  public enableAiFeaturesCheckbox!: HTMLInputElement;
  public llmProviderSelect!: HTMLSelectElement;
  public llmModelInput!: HTMLInputElement;
  public llmDeploymentNameInput!: HTMLInputElement;
  public llmApiVersionInput!: HTMLInputElement;
  public azureSpecificOptions!: HTMLDivElement;

  // Enhanced UI elements
  private statsDashboard!: HTMLDivElement;
  public showStatsButton!: HTMLButtonElement;
  public toggleStatsButton!: HTMLButtonElement;
  private advancedFilters!: HTMLDivElement;
  public toggleAdvancedFiltersButton!: HTMLButtonElement;
  private threadModal!: HTMLDivElement;
  private threadSubject!: HTMLHeadingElement;
  private threadBody!: HTMLDivElement;

  // Statistics elements
  private totalEmailsAnalyzedSpan!: HTMLSpanElement;
  private needingFollowupSpan!: HTMLSpanElement;
  private highPriorityCountSpan!: HTMLSpanElement;
  private avgResponseTimeSpan!: HTMLSpanElement;

  // Filter elements
  public priorityFilter!: HTMLSelectElement;
  public responseTimeFilter!: HTMLSelectElement;
  public subjectFilter!: HTMLInputElement;
  public senderFilter!: HTMLInputElement;
  public aiSuggestionFilter!: HTMLSelectElement;
  public clearFiltersButton!: HTMLButtonElement;

  // Progress elements
  private loadingStep!: HTMLSpanElement;
  private loadingDetail!: HTMLDivElement;
  private progressFill!: HTMLDivElement;

  // Action handlers
  private onActionCallback?: (action: string, emailId: string) => void;

  constructor() {
    this.initializeElements();
  }

  public setOnActionCallback(callback: (action: string, emailId: string) => void) {
    this.onActionCallback = callback;
  }

  private initializeElements(): void {
    const safeElement = (id: string): any => {
      const el = document.getElementById(id);
      if (el) return el;
      // Stub for testing/headless
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
        dataset: {},
        selectedOptions: [],
      } as any;
    };

    // Main controls
    this.analyzeButton = safeElement("analyzeButton");
    this.refreshButton = safeElement("refreshButton");
    this.settingsButton = safeElement("settingsButton");
    this.emailCountSelect = safeElement("emailCount");
    this.daysBackSelect = safeElement("daysBack");
    this.accountFilterSelect = safeElement("accountFilter");
    this.enableLlmSummaryCheckbox = safeElement("enableLlmSummary");
    this.enableLlmSuggestionsCheckbox = safeElement("enableLlmSuggestions");

    // Display elements
    this.statusDiv = safeElement("status");
    this.loadingDiv = safeElement("loadingMessage");
    this.emptyStateDiv = safeElement("emptyState");
    this.emailListDiv = safeElement("emailList");

    // Modal elements
    this.snoozeModal = safeElement("snoozeModal");
    this.settingsModal = safeElement("settingsModal");
    this.snoozeOptionsSelect = safeElement("snoozeOptions");
    this.customSnoozeGroup = safeElement("customSnoozeGroup");
    this.customSnoozeDate = safeElement("customSnoozeDate");
    this.llmEndpointInput = safeElement("llmEndpoint");
    this.llmApiKeyInput = safeElement("llmApiKey");
    this.showSnoozedEmailsCheckbox = safeElement("showSnoozedEmails");
    this.showDismissedEmailsCheckbox = safeElement("showDismissedEmails");
    this.aiStatusDiv = safeElement("aiStatus");
    this.aiStatusText = safeElement("aiStatusText");
    this.testAiConnectionButton = safeElement("testAiConnection");
    this.disableAiFeaturesButton = safeElement("disableAiFeatures");
    this.enableAiFeaturesCheckbox = safeElement("enableAiFeatures");
    this.llmProviderSelect = safeElement("llmProvider");
    this.llmModelInput = safeElement("llmModel");
    this.llmDeploymentNameInput = safeElement("llmDeploymentName");
    this.llmApiVersionInput = safeElement("llmApiVersion");
    this.azureSpecificOptions = safeElement("azureSpecificOptions");

    // New UI elements
    this.statsDashboard = safeElement("statsDashboard");
    this.showStatsButton = safeElement("showStatsButton");
    this.toggleStatsButton = safeElement("toggleStats");
    this.advancedFilters = safeElement("advancedFilters");
    this.toggleAdvancedFiltersButton = safeElement("toggleAdvancedFilters");
    this.threadModal = safeElement("threadModal");
    this.threadSubject = safeElement("threadSubject");
    this.threadBody = safeElement("threadBody");

    // Statistics elements
    this.totalEmailsAnalyzedSpan = safeElement("totalEmailsAnalyzed");
    this.needingFollowupSpan = safeElement("needingFollowup");
    this.highPriorityCountSpan = safeElement("highPriorityCount");
    this.avgResponseTimeSpan = safeElement("avgResponseTime");

    // Filter elements
    this.priorityFilter = safeElement("priorityFilter");
    this.responseTimeFilter = safeElement("responseTimeFilter");
    this.subjectFilter = safeElement("subjectFilter");
    this.senderFilter = safeElement("senderFilter");
    this.aiSuggestionFilter = safeElement("aiSuggestionFilter");
    this.clearFiltersButton = safeElement("clearFilters");

    // Progress elements
    this.loadingStep = safeElement("loadingStep");
    this.loadingDetail = safeElement("loadingDetail");
    this.progressFill = safeElement("progressFill");
  }

  public showStatus(message: string, type: "success" | "error" | "warning"): void {
    this.statusDiv.textContent = message;
    this.statusDiv.className = `status ${type}`;
    this.statusDiv.style.display = "block";

    if (type === "success") {
      setTimeout(() => this.hideStatus(), 5000);
    }
  }

  public hideStatus(): void {
    this.statusDiv.style.display = "none";
  }

  public setLoadingState(isLoading: boolean): void {
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

  public updateProgress(percentage: number, step: string, detail: string): void {
    this.progressFill.style.width = `${percentage}%`;
    this.loadingStep.textContent = step;
    this.loadingDetail.textContent = detail;
  }

  public getCurrentProgress(): number {
    return parseInt(this.progressFill.style.width) || 0;
  }

  public hideAllStates(): void {
    this.loadingDiv.style.display = "none";
    this.emptyStateDiv.style.display = "none";
    this.emailListDiv.style.display = "none";
  }

  public displayEmails(emails: FollowupEmail[]): void {
    this.hideAllStates();

    if (emails.length === 0) {
      this.emptyStateDiv.innerHTML = `
                <h3>No emails need follow-up</h3>
                <p>Great! You're all caught up with your email responses.</p>
            `;
      this.emptyStateDiv.style.display = "block";
      return;
    }

    // Always display most recent emails first
    const recentFirst = [...emails].sort(
      (a, b) => b.sentDate.getTime() - a.sentDate.getTime()
    );

    this.emailListDiv.innerHTML = "";
    recentFirst.forEach((email) => {
      const emailElement = this.createEmailElement(email);
      this.emailListDiv.appendChild(emailElement);
    });

    this.emailListDiv.style.display = "block";
  }

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

    // Calculate confidence score (mock)
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
          (email.llmSuggestion ? 20 : 0)
      )
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
            ${
              email.llmSuggestion
                ? `<div class="llm-suggestion"><strong>AI Suggestion:</strong> ${this.escapeHtml(
                    email.llmSuggestion
                  )}</div>`
                : ""
            }
            <div class="email-actions">
                <button class="action-button" data-email-id="${email.id}" data-action="reply">Reply</button>
                <button class="action-button" data-email-id="${email.id}" data-action="forward">Forward</button>
                <button class="action-button" data-email-id="${email.id}" data-action="snooze">Snooze</button>
                <button class="action-button" data-email-id="${email.id}" data-action="dismiss">Dismiss</button>
                ${
                  email.threadMessages.length > 1
                    ? `<button class="action-button" data-email-id="${email.id}" data-action="view-thread">View Thread (${email.threadMessages.length})</button>`
                    : ""
                }
            </div>
        `;

    // Attach event listeners to action buttons
    const actionButtons = emailDiv.querySelectorAll(".action-button");
    actionButtons.forEach((button) => {
      button.addEventListener("click", (e) => {
        const target = e.target as HTMLElement;
        const emailId = target.dataset.emailId;
        const action = target.dataset.action;
        if (emailId && action && this.onActionCallback) {
          this.onActionCallback(action, emailId);
        }
      });
    });

    return emailDiv;
  }

  public updateStatistics(
    total: number,
    needingFollowup: number,
    highPriority: number,
    avgDays: number
  ): void {
    this.totalEmailsAnalyzedSpan.textContent = total.toString();
    this.needingFollowupSpan.textContent = needingFollowup.toString();
    this.highPriorityCountSpan.textContent = highPriority.toString();
    this.avgResponseTimeSpan.textContent = avgDays.toString();
  }

  public toggleStatsDashboard(show: boolean): void {
    if (show) {
      this.statsDashboard.classList.add("show");
      this.showStatsButton.style.display = "none";
    } else {
      this.statsDashboard.classList.remove("show");
      this.showStatsButton.style.display = "inline-block";
    }
  }

  public toggleAdvancedFilters(): void {
    this.advancedFilters.classList.toggle("show");
    const isShown = this.advancedFilters.classList.contains("show");
    this.toggleAdvancedFiltersButton.textContent = isShown
      ? "Hide Advanced Filters"
      : "Advanced Filters";
  }

  public populateAccountFilter(
    selectedAccounts: string[],
    availableAccounts: string[]
  ): void {
    this.accountFilterSelect.innerHTML = "";

    if (availableAccounts.length === 1) {
      const account = availableAccounts[0];
      const option = document.createElement("option");
      option.value = account;
      option.textContent = `${account} (Current)`;
      option.selected = true;
      this.accountFilterSelect.appendChild(option);
      this.accountFilterSelect.disabled = true;
      return;
    }

    this.accountFilterSelect.disabled = false;
    availableAccounts.forEach((account) => {
      const option = document.createElement("option");
      option.value = account;
      option.textContent = account;
      option.selected = selectedAccounts.includes(account);
      this.accountFilterSelect.appendChild(option);
    });
  }

  public populateSnoozeOptions(options: SnoozeOption[]): void {
    this.snoozeOptionsSelect.innerHTML = "";

    if (!options || options.length === 0) {
      // Fallback
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
    });
  }

  public showThreadView(email: FollowupEmail): void {
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

  public hideThreadModal(): void {
    this.threadModal.style.display = "none";
  }

  public setAiStatus(type: "success" | "warning" | "error", message: string): void {
    this.aiStatusDiv.style.display = "block";
    this.aiStatusDiv.className = `ai-status ${type}`;
    this.aiStatusText.textContent = message;
  }

  // Helper methods
  private escapeHtml(text: string): string {
    const div = document.createElement("div");
    div.textContent = text;
    return div.innerHTML;
  }

  private decodeHtmlEntities(text: string): string {
    const parser = new DOMParser();
    const doc = parser.parseFromString(text, "text/html");
    return doc.documentElement.textContent || "";
  }

  private renderSummaryHtml(raw: string): string {
    if (!raw) return "";
    let normalized = raw.replace(/&nbsp;/gi, " ").replace(/\u00A0/g, " ");
    normalized = this.decodeHtmlEntities(normalized);
    const lines = normalized
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter((l) => l.length > 0);
    const safe = lines.map((l) => this.escapeHtml(l));
    return safe.join("<br>");
  }

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
}
