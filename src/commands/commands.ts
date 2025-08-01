// Commands for ribbon buttons and other Office UI interactions
import { EmailAnalysisService } from '../services/EmailAnalysisService';
import { ConfigurationService } from '../services/ConfigurationService';
import { ThreadMessage } from '../models/FollowupEmail';

class CommandHandler {
    private emailAnalysisService: EmailAnalysisService;
    private configurationService: ConfigurationService;

    constructor() {
        this.emailAnalysisService = new EmailAnalysisService();
        this.configurationService = new ConfigurationService();
    }

    // Show notification helper
    private showNotification(message: string, type: 'success' | 'error' | 'info' = 'info'): void {
        const notification = document.getElementById('notification');
        if (notification) {
            notification.textContent = message;
            notification.className = `notification ${type} show`;
            
            setTimeout(() => {
                notification.classList.remove('show');
            }, 3000);
        }
    }

    // Quick analyze command - analyzes last 10 emails
    public async quickAnalyze(event: Office.AddinCommands.Event): Promise<void> {
        try {
            const followupEmails = await this.emailAnalysisService.analyzeEmails(10, 7, []);
            const count = followupEmails.length;
            
            if (count > 0) {
                this.showNotification(`Found ${count} email(s) needing follow-up`, 'success');
                // Store results for task pane to display
                await this.configurationService.cacheAnalysisResults(followupEmails);
            } else {
                this.showNotification('No emails need follow-up', 'success');
            }
        } catch (error) {
            console.error('Quick analyze error:', error);
            this.showNotification('Error analyzing emails', 'error');
        } finally {
            event.completed();
        }
    }

    // Mark current email as followed up
    public async markAsFollowedUp(event: Office.AddinCommands.Event): Promise<void> {
        try {
            const item = Office.context.mailbox.item;
            if (item && item.itemId) {
                this.emailAnalysisService.dismissEmail(item.itemId);
                this.showNotification('Email marked as followed up', 'success');
            } else {
                this.showNotification('No email selected', 'error');
            }
        } catch (error) {
            console.error('Mark as followed up error:', error);
            this.showNotification('Error marking email', 'error');
        } finally {
            event.completed();
        }
    }

    // Snooze current email for later follow-up
    public async snoozeCurrentEmail(event: Office.AddinCommands.Event): Promise<void> {
        try {
            const item = Office.context.mailbox.item;
            if (item && item.itemId) {
                // Snooze for 24 hours by default
                this.emailAnalysisService.snoozeEmail(item.itemId, 24 * 60);
                this.showNotification('Email snoozed for 24 hours', 'success');
            } else {
                this.showNotification('No email selected', 'error');
            }
        } catch (error) {
            console.error('Snooze email error:', error);
            this.showNotification('Error snoozing email', 'error');
        } finally {
            event.completed();
        }
    }

    // Generate AI follow-up suggestion for current email
    public async generateFollowupSuggestion(event: Office.AddinCommands.Event): Promise<void> {
        try {
            const item = Office.context.mailbox.item;
            if (item && item.itemId) {
                // Create a mock ThreadMessage from the current email
                const currentMessage: ThreadMessage = {
                    id: item.itemId,
                    subject: item.subject || 'No Subject',
                    from: Office.context.mailbox.userProfile.emailAddress,
                    to: [], // Would need to be populated from item recipients
                    sentDate: new Date(), // Would need to be populated from item date
                    body: '', // Would need to be populated from item body
                    isFromCurrentUser: true
                };

                const followupEmail = await (this.emailAnalysisService as any).createFollowupEmailEnhanced(
                    currentMessage,
                    [currentMessage],
                    Office.context.mailbox.userProfile.emailAddress
                );

                if (followupEmail.llmSuggestion) {
                    this.showNotification(`Suggestion: ${followupEmail.llmSuggestion}`, 'info');
                } else {
                    this.showNotification('No AI suggestion available', 'info');
                }
            } else {
                this.showNotification('No email selected', 'error');
            }
        } catch (error) {
            console.error('Generate followup suggestion error:', error);
            this.showNotification('Error generating suggestion', 'error');
        } finally {
            event.completed();
        }
    }
}

const commandHandler = new CommandHandler();

Office.onReady(() => {
    // Commands are ready
});

// Function to show the task pane (called from ribbon button)
function showTaskpane(event: Office.AddinCommands.Event) {
    // The task pane will be shown automatically when the button is clicked
    // due to the ShowTaskpane action in the manifest
    event.completed();
}

// Quick analyze function for ribbon command
function quickAnalyze(event: Office.AddinCommands.Event) {
    commandHandler.quickAnalyze(event);
}

// Mark as followed up function for ribbon command
function markAsFollowedUp(event: Office.AddinCommands.Event) {
    commandHandler.markAsFollowedUp(event);
}

// Snooze current email function for ribbon command
function snoozeCurrentEmail(event: Office.AddinCommands.Event) {
    commandHandler.snoozeCurrentEmail(event);
}

// Generate AI suggestion function for ribbon command
function generateFollowupSuggestion(event: Office.AddinCommands.Event) {
    commandHandler.generateFollowupSuggestion(event);
}

// Register the functions so they can be called from the manifest
(Office as any).actions = {
    showTaskpane: showTaskpane,
    quickAnalyze: quickAnalyze,
    markAsFollowedUp: markAsFollowedUp,
    snoozeCurrentEmail: snoozeCurrentEmail,
    generateFollowupSuggestion: generateFollowupSuggestion
};