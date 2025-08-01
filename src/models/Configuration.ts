export interface Configuration {
    emailCount: number;
    daysBack: number;
    lastAnalysisDate: Date;
    autoRefreshInterval?: number; // minutes
    priorityThresholds?: {
        high: number; // days without response
        medium: number;
        low: number;
    };
    // New properties for enhanced functionality
    snoozeOptions: SnoozeOption[];
    enableLlmSummary: boolean;
    enableLlmSuggestions: boolean;
    llmApiEndpoint?: string;
    llmApiKey?: string;
    llmModel?: string;
    selectedAccounts: string[]; // email addresses of accounts to analyze
    showSnoozedEmails: boolean;
    showDismissedEmails: boolean;
}

export interface SnoozeOption {
    label: string;
    value: number; // minutes
    isCustom?: boolean;
}