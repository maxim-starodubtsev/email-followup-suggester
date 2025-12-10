<<<<<<< Current (Your changes)
=======
# Updated Requirements for Followup Suggester

## 1. Autonomous Analysis & Scope
- **Scope**: The add-in must work across **all emails** in the **current user's mailbox**.
  - *Clarification*: "Connected mailboxes" interprets as **all folders** (Inbox, Sent, Drafts, Archive, sub-folders) within the active Outlook profile/account. Due to Office.js limitations, accessing *other* distinct accounts (delegated/shared) is not supported in this version without specific EWS/Graph permissions, but the add-in handles the current mailbox globally.
- **Autonomy**: It must work **without the user selecting a specific email** to start.
  - The "Bulk Analysis" mode is the primary function.
  - It uses `FindItem` (EWS) to scan the mailbox independent of the reading pane.
- **Constraint**: **NO EWS / GraphQL** for complex operations (specifically Conversation grouping).
  - *Implication*: We rely on **Client-Side Threading** using `makeEwsRequestAsync` (FindItem) only for basic item retrieval.

## 2. Smart Threading Algorithm (Client-Side)
- **Logic**: Build email threads by finding emails with the **same subject** across folders (Sent, Inbox, Drafts, Archive).
  - *Requirement*: Use a "smart algorithm" that normalizes subjects (stripping RE:, FW:, etc.) and matches them.
  - *Fallback*: Use body containment checks ("artificial threading") if subject matching is ambiguous or to handle thread fragmentation.
- **Sorting**: Order the thread by email date (oldest to newest).
- **Analysis**:
  - Identify the **last sender**.
  - If last sender is **Current User**, it is a candidate for follow-up.
  - Check if there are **no subsequent responses** from others.

## 3. Intelligent Content Analysis (LLM / Heuristic)
- **Goal**: Determine if the last email *needs* a follow-up.
- **Criteria**:
  - **Closing Email**: (e.g., "Thank you", "Final update", "To sum up") -> **NO Follow-up**.
  - **Pending Question/Action**: (e.g., "Did you look?", "Appreciate response", "Need help") -> **YES Follow-up**.
- **Mechanism**:
  - Use **LLM Service** (EPAM DIAL / OpenAI) if available.
  - Fallback to **Smart In-built Logic** (keywords/heuristics) if LLM is offline.

## 4. User Interaction & Follow-up Action
- **Action**: When user selects an item to follow up:
  - **Open the exact email item**.
  - **Trigger/Enable "Reply All"**: Use Outlook's in-built functionality to reply to the *same* entity.
    - *Implementation*: Use `Office.context.mailbox.displayMessageForm(itemId)` to open the email inspector. This allows the user to click the native "Reply All" button, preserving thread context, recipients, and formatting perfectly.
    - *Avoid*: Creating a new "Draft" programmatically unless opening the item fails, as creating a draft might lose thread references or formatting nuances.

## 5. Refactoring Plan
- **Task 1**: **Verify "Reply" Action**. Ensure `taskpane.ts` uses `displayMessageForm` as the primary action for the "Reply" button. Remove/deprioritize `displayNewMessageForm` (draft creation) unless as a strict error fallback.
- **Task 2**: **Enforce Smart Threading**. In `EmailAnalysisService`, ensure `buildThreadBySubject` / `groupEmailsBySubject` is the primary logic, and EWS `GetConversationItems` is only secondary or removed if it violates the "No EWS" spirit (though `FindItem` is EWS, "Conversation" operations are often flakier). *Decision*: Stick to the current hybrid but ensure Subject Matching is robust.
- **Task 3**: **Clarify Account Selection**. In `ConfigurationService`, clarify that "Available Accounts" currently only returns the active user, and update UI to reflect this single-account scope to avoid user confusion.
- **Task 4**: **LLM Prompt Tuning**. Verify `LlmService.analyzeFollowupNeed` prompt specifically targets the "Closing vs Pending" distinction as requested.
>>>>>>> Incoming (Background Agent changes)
