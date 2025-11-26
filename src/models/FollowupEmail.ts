export interface FollowupEmail {
  id: string;
  subject: string;
  recipients: string[];
  sentDate: Date;
  body: string;
  summary: string;
  priority: "high" | "medium" | "low";
  daysWithoutResponse: number;
  conversationId?: string;
  hasAttachments: boolean;
  // New properties
  accountEmail: string;
  threadMessages: ThreadMessage[];
  isSnoozed: boolean;
  snoozeUntil?: Date;
  isDismissed: boolean;
  llmSuggestion?: string;
  llmSummary?: string;
  sentiment?: "positive" | "neutral" | "negative" | "urgent"; // Add sentiment field
}

export interface ThreadMessage {
  id: string;
  // EWS ChangeKey for the item, used by EWS operations that require a reference
  changeKey?: string;
  subject: string;
  from: string;
  to: string[];
  sentDate: Date;
  receivedDate?: Date; // New: DateTimeReceived for accurate last-message determination
  body: string;
  isFromCurrentUser: boolean;
}

export interface SnoozeOption {
  label: string;
  value: number; // minutes
  isCustom?: boolean;
}
