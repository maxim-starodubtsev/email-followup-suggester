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
    // Enhanced functionality
    snoozeOptions: SnoozeOption[];
    // LLM Service Configuration
    enableLlmSummary: boolean;
    enableLlmSuggestions: boolean;
    llmProvider?: 'azure' | 'dial' | 'openai'; // API provider type (default: 'dial')
    llmApiEndpoint?: string; // Base URL for API calls (default: 'http://localhost:8080' for DIAL)
    llmApiKey?: string; // API key for authentication
    llmModel?: string; // Model name (default: 'gpt-4o-mini' for DIAL)
    llmApiVersion?: string; // For Azure OpenAI API version (default: '2024-02-01')
    llmDeploymentName?: string; // For Azure OpenAI deployment name
    selectedAccounts: string[]; // email addresses of accounts to analyze
    showSnoozedEmails: boolean;
    showDismissedEmails: boolean;
}

export interface SnoozeOption {
    label: string;
    value: number; // minutes
    isCustom?: boolean;
}