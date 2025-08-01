import { Configuration } from '../models/Configuration';
import { FollowupEmail } from '../models/FollowupEmail';
import { CacheService } from './CacheService';

export class ConfigurationService {
    private readonly STORAGE_KEY = 'followup-suggester-config';
    private cacheService: CacheService;
    
    constructor() {
        this.cacheService = new CacheService({
            defaultTtl: 30 * 60 * 1000, // 30 minutes
            maxMemoryUsage: 10 * 1024 * 1024, // 10MB
            maxEntries: 1000
        });
    }

    private readonly DEFAULT_CONFIG: Configuration = {
        emailCount: 25,
        daysBack: 7,
        lastAnalysisDate: new Date(),
        autoRefreshInterval: 30,
        priorityThresholds: {
            high: 7,
            medium: 3,
            low: 1
        },
        // New enhanced properties
        snoozeOptions: [
            { label: '15 minutes', value: 15 },
            { label: '1 hour', value: 60 },
            { label: '4 hours', value: 240 },
            { label: '1 day', value: 1440 },
            { label: '3 days', value: 4320 },
            { label: '1 week', value: 10080 },
            { label: 'Custom...', value: 0, isCustom: true }
        ],
        enableLlmSummary: false,
        enableLlmSuggestions: false,
        selectedAccounts: [],
        showSnoozedEmails: false,
        showDismissedEmails: false
    };

    public async getConfiguration(): Promise<Configuration> {
        try {
            // Try to get configuration from Office settings
            return new Promise((resolve) => {
                const settings = Office.context.roamingSettings;
                const stored = settings.get(this.STORAGE_KEY);
                
                if (stored && typeof stored === 'object') {
                    // Merge with defaults to ensure all properties exist
                    const config = { ...this.DEFAULT_CONFIG, ...stored };
                    
                    // Ensure date is properly converted
                    if (typeof config.lastAnalysisDate === 'string') {
                        config.lastAnalysisDate = new Date(config.lastAnalysisDate);
                    }
                    
                    resolve(config);
                } else {
                    resolve(this.DEFAULT_CONFIG);
                }
            });
        } catch (error) {
            console.error('Error loading configuration from Office settings:', error);
            
            // Fallback to localStorage
            try {
                const stored = localStorage.getItem(this.STORAGE_KEY);
                if (stored) {
                    const parsed = JSON.parse(stored);
                    const config = { ...this.DEFAULT_CONFIG, ...parsed };
                    
                    if (typeof config.lastAnalysisDate === 'string') {
                        config.lastAnalysisDate = new Date(config.lastAnalysisDate);
                    }
                    
                    return config;
                }
            } catch (localStorageError) {
                console.error('Error loading from localStorage:', localStorageError);
            }
            
            return this.DEFAULT_CONFIG;
        }
    }

    public async saveConfiguration(config: Configuration): Promise<void> {
        try {
            // Save to Office settings (roams across devices)
            return new Promise((resolve, reject) => {
                const configToSave = {
                    ...config,
                    lastAnalysisDate: config.lastAnalysisDate.toISOString()
                };

                Office.context.roamingSettings.set(this.STORAGE_KEY, configToSave);
                Office.context.roamingSettings.saveAsync((result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resolve();
                    } else {
                        // Fallback to localStorage instead of rejecting
                        try {
                            localStorage.setItem(this.STORAGE_KEY, JSON.stringify(configToSave));
                            resolve();
                        } catch (localStorageError) {
                            console.error('Error saving to localStorage:', localStorageError);
                            reject(new Error('Failed to save configuration'));
                        }
                    }
                });
            });
        } catch (error) {
            console.error('Error saving to Office settings:', error);
            
            // Fallback to localStorage
            try {
                const configToSave = {
                    ...config,
                    lastAnalysisDate: config.lastAnalysisDate.toISOString()
                };
                localStorage.setItem(this.STORAGE_KEY, JSON.stringify(configToSave));
            } catch (localStorageError) {
                console.error('Error saving to localStorage:', localStorageError);
                throw new Error('Failed to save configuration');
            }
        }
    }

    public async resetConfiguration(): Promise<void> {
        await this.saveConfiguration(this.DEFAULT_CONFIG);
    }

    public getDefaultConfiguration(): Configuration {
        return { ...this.DEFAULT_CONFIG };
    }

    public async updateLastAnalysisDate(): Promise<void> {
        const config = await this.getConfiguration();
        config.lastAnalysisDate = new Date();
        await this.saveConfiguration(config);
    }

    public async isAnalysisStale(maxAgeMinutes: number = 30): Promise<boolean> {
        const config = await this.getConfiguration();
        const now = new Date();
        const diffMinutes = (now.getTime() - config.lastAnalysisDate.getTime()) / (1000 * 60);
        return diffMinutes > maxAgeMinutes;
    }

    public async getAvailableAccounts(): Promise<string[]> {
        // Get all configured accounts in Outlook
        try {
            return new Promise((resolve) => {
                Office.context.mailbox.getUserIdentityTokenAsync((result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        // For now, return the current user's email
                        // This can be enhanced to get multiple accounts if available
                        const currentEmail = Office.context.mailbox.userProfile.emailAddress;
                        resolve([currentEmail]);
                    } else {
                        resolve([]);
                    }
                });
            });
        } catch (error) {
            console.error('Error getting available accounts:', error);
            return [];
        }
    }

    public async updateSelectedAccounts(accounts: string[]): Promise<void> {
        const config = await this.getConfiguration();
        config.selectedAccounts = accounts;
        await this.saveConfiguration(config);
    }

    public async updateLlmSettings(endpoint: string, apiKey: string, enableSummary: boolean, enableSuggestions: boolean): Promise<void> {
        const config = await this.getConfiguration();
        config.llmApiEndpoint = endpoint;
        config.llmApiKey = apiKey;
        config.enableLlmSummary = enableSummary;
        config.enableLlmSuggestions = enableSuggestions;
        await this.saveConfiguration(config);
    }

    /**
     * Cache analysis results for task pane display
     */
    public async cacheAnalysisResults(followupEmails: FollowupEmail[]): Promise<void> {
        try {
            const cacheKey = `analysis_results_${Date.now()}`;
            this.cacheService.set(cacheKey, followupEmails, 30 * 60 * 1000); // 30 minutes TTL
            
            // Store the cache key using Office settings
            return new Promise((resolve, reject) => {
                Office.context.roamingSettings.set('last_analysis_cache_key', cacheKey);
                Office.context.roamingSettings.saveAsync((result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resolve();
                    } else {
                        // Fallback to localStorage
                        try {
                            localStorage.setItem('last_analysis_cache_key', cacheKey);
                            resolve();
                        } catch (error) {
                            reject(error);
                        }
                    }
                });
            });
        } catch (error) {
            console.error('Error caching analysis results:', error);
        }
    }

    /**
     * Get cached analysis results
     */
    public async getCachedAnalysisResults(): Promise<FollowupEmail[]> {
        try {
            // Get cache key from Office settings
            const cacheKey = Office.context.roamingSettings.get('last_analysis_cache_key') as string;
            if (cacheKey) {
                const results = this.cacheService.get(cacheKey) as FollowupEmail[];
                return results || [];
            }
            
            // Fallback to localStorage
            const fallbackKey = localStorage.getItem('last_analysis_cache_key');
            if (fallbackKey) {
                const results = this.cacheService.get(fallbackKey) as FollowupEmail[];
                return results || [];
            }
            
            return [];
        } catch (error) {
            console.error('Error getting cached analysis results:', error);
            return [];
        }
    }

    /**
     * Cache AI followup suggestion for specific email
     */
    public async cacheFollowupSuggestion(emailId: string, suggestion: string): Promise<void> {
        try {
            const cacheKey = `followup_suggestion_${emailId}`;
            this.cacheService.set(cacheKey, suggestion, 24 * 60 * 60 * 1000); // 24 hours TTL
        } catch (error) {
            console.error('Error caching followup suggestion:', error);
        }
    }

    /**
     * Get cached followup suggestion for specific email
     */
    public async getCachedFollowupSuggestion(emailId: string): Promise<string | null> {
        try {
            const cacheKey = `followup_suggestion_${emailId}`;
            return this.cacheService.get(cacheKey) as string | null;
        } catch (error) {
            console.error('Error getting cached followup suggestion:', error);
            return null;
        }
    }
}