import { describe, it, expect, beforeEach, vi, afterEach } from "vitest";
import { AppController } from "../../src/taskpane/services/AppController";
import { UiService } from "../../src/taskpane/services/UiService";
import { EmailAnalysisService } from "../../src/services/EmailAnalysisService";
import { ConfigurationService } from "../../src/services/ConfigurationService";
import { FollowupEmail } from "../../src/models/FollowupEmail";

// Mock dependencies
vi.mock("../../src/services/EmailAnalysisService");
vi.mock("../../src/services/ConfigurationService");
vi.mock("../../src/services/LlmService");
vi.mock("../../src/services/RetryService");

describe("AppController", () => {
  let appController: AppController;
  let mockUiService: UiService;
  let mockEmailService: any;
  let mockConfigService: any;

  beforeEach(() => {
    // Reset DOM
    document.body.innerHTML = `
      <button id="analyzeButton"></button>
      <button id="refreshButton"></button>
      <button id="settingsButton"></button>
      <select id="emailCount"><option value="10">10</option></select>
      <select id="daysBack"><option value="7">7</option></select>
      <select id="accountFilter"></select>
      <input type="checkbox" id="enableLlmSummary" />
      <input type="checkbox" id="enableLlmSuggestions" />
      
      <div id="status"></div>
      <div id="loadingMessage"></div>
      <div id="emptyState"></div>
      <div id="emailList"></div>
      
      <div id="snoozeModal"></div>
      <div id="settingsModal"></div>
      <select id="snoozeOptions"></select>
      <div id="customSnoozeGroup"></div>
      <input id="customSnoozeDate" />
      <input id="llmEndpoint" />
      <input id="llmApiKey" />
      <input type="checkbox" id="showSnoozedEmails" />
      <input type="checkbox" id="showDismissedEmails" />
      <div id="aiStatus"></div>
      <span id="aiStatusText"></span>
      <button id="testAiConnection"></button>
      <button id="disableAiFeatures"></button>
      <input type="checkbox" id="enableAiFeatures" />
      <select id="llmProvider"></select>
      <input id="llmModel" />
      <input id="llmDeploymentName" />
      <input id="llmApiVersion" />
      <div id="azureSpecificOptions"></div>

      <div id="statsDashboard"></div>
      <button id="showStatsButton"></button>
      <button id="toggleStats"></button>
      <div id="advancedFilters"></div>
      <button id="toggleAdvancedFilters"></button>
      <div id="threadModal"></div>
      <h2 id="threadSubject"></h2>
      <div id="threadBody"></div>

      <span id="totalEmailsAnalyzed"></span>
      <span id="needingFollowup"></span>
      <span id="highPriorityCount"></span>
      <span id="avgResponseTime"></span>

      <select id="priorityFilter"></select>
      <select id="responseTimeFilter"></select>
      <input id="subjectFilter" />
      <input id="senderFilter" />
      <select id="aiSuggestionFilter"></select>
      <button id="clearFilters"></button>

      <span id="loadingStep"></span>
      <div id="loadingDetail"></div>
      <div id="progressFill"></div>
      
      <button id="confirmSnooze"></button>
      <button id="cancelSnooze"></button>
      <button id="saveSettings"></button>
      <button id="cancelSettings"></button>
    `;

    // Mock UiService methods
    mockUiService = new UiService();
    // We need to spy on UiService methods to verify calls
    vi.spyOn(mockUiService, "showStatus");
    vi.spyOn(mockUiService, "setLoadingState");
    vi.spyOn(mockUiService, "updateProgress");
    vi.spyOn(mockUiService, "displayEmails");
    vi.spyOn(mockUiService, "updateStatistics");
    vi.spyOn(mockUiService, "populateAccountFilter");
    vi.spyOn(mockUiService, "populateSnoozeOptions");

    // Initialize services mocks
    mockEmailService = EmailAnalysisService.prototype;
    mockEmailService.analyzeEmails = vi.fn().mockResolvedValue([]);
    mockEmailService.setConfiguration = vi.fn();
    mockEmailService.setLlmService = vi.fn();

    mockConfigService = ConfigurationService.prototype;
    mockConfigService.getConfiguration = vi.fn().mockResolvedValue({
      emailCount: 10,
      daysBack: 7,
      selectedAccounts: [],
      snoozeOptions: []
    });
    mockConfigService.getAvailableAccounts = vi.fn().mockResolvedValue(["test@example.com"]);
    mockConfigService.getCachedAnalysisResults = vi.fn().mockResolvedValue([]);

    appController = new AppController(mockUiService);
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("should initialize correctly", async () => {
    await appController.initialize();
    
    expect(mockConfigService.getConfiguration).toHaveBeenCalled();
    expect(mockUiService.populateAccountFilter).toHaveBeenCalled();
    expect(mockUiService.populateSnoozeOptions).toHaveBeenCalled();
  });

  it("should handle analyze emails flow", async () => {
    const mockEmails: FollowupEmail[] = [
      {
        id: "1",
        subject: "Test Email",
        priority: "high",
        sentDate: new Date(),
        recipients: ["recipient@example.com"],
        body: "body",
        summary: "summary",
        daysWithoutResponse: 5,
        hasAttachments: false,
        accountEmail: "me@example.com",
        threadMessages: [],
        isSnoozed: false,
        isDismissed: false
      }
    ];
    mockEmailService.analyzeEmails.mockResolvedValue(mockEmails);

    await appController.initialize();
    
    // Trigger analysis (simulate button click logic)
    await (appController as any).analyzeEmails();

    expect(mockUiService.setLoadingState).toHaveBeenCalledWith(true);
    expect(mockEmailService.analyzeEmails).toHaveBeenCalled();
    expect(mockUiService.displayEmails).toHaveBeenCalledWith(mockEmails);
    expect(mockUiService.updateStatistics).toHaveBeenCalled();
    expect(mockUiService.setLoadingState).toHaveBeenCalledWith(false);
  });

  it("should filter emails correctly", async () => {
    const mockEmails: FollowupEmail[] = [
      {
        id: "1",
        subject: "High Priority",
        priority: "high",
        sentDate: new Date(),
        recipients: [],
        body: "",
        summary: "",
        daysWithoutResponse: 5,
        hasAttachments: false,
        accountEmail: "",
        threadMessages: [],
        isSnoozed: false,
        isDismissed: false
      },
      {
        id: "2",
        subject: "Low Priority",
        priority: "low",
        sentDate: new Date(),
        recipients: [],
        body: "",
        summary: "",
        daysWithoutResponse: 1,
        hasAttachments: false,
        accountEmail: "",
        threadMessages: [],
        isSnoozed: false,
        isDismissed: false
      }
    ];
    
    // Inject emails directly for testing filter logic
    (appController as any).allEmails = mockEmails;
    
    // Set filter
    mockUiService.priorityFilter.value = "high";
    
    // Trigger filter
    (appController as any).applyFilters();

    expect(mockUiService.displayEmails).toHaveBeenCalledWith([mockEmails[0]]);
  });
});
