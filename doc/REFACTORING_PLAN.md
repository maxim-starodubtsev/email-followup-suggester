# Refactoring & Fixing Plan

Based on the [Updated Requirements](REQUIREMENTS-UPDATED.md), the following tasks are planned to align the implementation with the user's specific requests.

## Task 1: Refine "Reply" Action Behavior
**Goal**: Ensure the "Reply" button opens the *existing* email item using Outlook's native inspector, allowing the user to click "Reply All" naturally.
- [ ] **Analyze**: Check `src/taskpane/taskpane.ts` `replyToEmail` method.
- [ ] **Fix**: 
  - Prioritize `Office.context.mailbox.displayMessageForm(itemId)`.
  - If `displayMessageForm` is unavailable, show a clear user notification instead of silently creating a new Draft (which breaks the "same entity" requirement).
  - Only use draft creation (`displayNewMessageForm`) if explicitly requested or as a last resort with a warning.

## Task 2: Robust Smart Threading
**Goal**: Ensure the "Smart Algorithm" (Subject Matching) is the primary driver for threading, minimizing reliance on EWS Conversation IDs which might be unavailable.
- [ ] **Analyze**: `src/services/EmailAnalysisService.ts`.
- [ ] **Refactor**: 
  - Verify `processConversationWithCaching` prefers `groupEmailsBySubject` groups.
  - Ensure `buildArtificialThreadFromRecentEmails` is utilized when conversation structure is missing.
  - *Current Status*: The code already uses `groupEmailsBySubject` as the main entry point. We will add a test case to confirm `normalizeSubject` handles the user's specific examples (e.g. "Re:", "FW:", "Update").

## Task 3: Clarify Account Scope
**Goal**: The UI implies "Filter by accounts", but the backend only supports the current account.
- [ ] **Update UI**: In `src/taskpane/taskpane.ts`, if only one account is available, hide or simplify the "Account Filter" dropdown to avoid confusion about "connected mailboxes".
- [ ] **Update Service**: In `src/services/ConfigurationService.ts`, add a comment or log indicating that `getAvailableAccounts` is scoped to the current profile due to platform limitations.

## Task 4: LLM Prompt Tuning
**Goal**: Ensure the LLM prompt matches the "Closing vs Pending" criteria.
- [ ] **Update**: `src/services/LlmService.ts` -> `analyzeFollowupNeed`.
- [ ] **Refine Prompt**: Add specific examples from the requirements:
  - Closing: "Thank you", "Final update", "To sum up".
  - Pending: "Did you look?", "Appreciate response", "Need help".

## Task 5: Verify "Autonomous" Work
**Goal**: Ensure the "Bulk Analysis" runs smoothly without selecting a specific email.
- [ ] **Verify**: The `analyzeEmails` flow triggers from the main "Analyze" button and uses `FindItem` on `["sentitems", "inbox", "drafts", "archive"]`. (Already implemented).

## Execution Order
1. Task 1 (Reply Action) - Critical User Experience requirement.
2. Task 4 (LLM Prompt) - Critical Core Logic requirement.
3. Task 3 (Account UI) - Polish.
4. Task 2 (Threading) - Verification.
