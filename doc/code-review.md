# Role
Act as a Senior Lead Software Architect and QA Engineer specializing in TypeScript, Outlook Office Add-ins, and LLM integrations.

# Context
I have a TypeScript-based Outlook Add-in project (`email-followup-suggester`) that uses LLMs to analyze emails and suggest replies. The project includes:
- **Frontend:** Office.js Taskpane (HTML/CSS/TS).
- **Core Logic:** Services for API calls, XML parsing, Caching, and Retry logic.
- **Testing:** Jest/Vitest with mocks.
- **Infrastructure:** Docker and Webpack.

# Objective
Perform a deep-scan code review of the entire repository to ensure the code is production-ready, bug-free, fully tested, and ready for commit.

# Review Instructions
Please analyze the `src/`, `tests/`, `manifest.xml`, `webpack.config.js`, and `package.json` files. Focus on the following five pillars:

## 1. Code Quality & Architecture
- **Type Safety:** specific checks for `any` types that should be defined interfaces.
- **Modularity:** Ensure separation of concerns (e.g., `Taskpane.ts` should handle UI, `LlmService.ts` should handle AI logic).
- **Async/Await:** Verify correct handling of Promises, especially regarding Office.js asynchronous operations.
- **Error Handling:** Check if `try-catch` blocks are implemented in all external service calls and if errors are surfaced meaningfully to the user.

## 2. Logic & Functionality Verification
- **LLM Integration:** Review `LlmService.ts` and `RetryService.ts`. Is the retry logic robust? Are API keys handled securely (not hardcoded)?
- **Parsing:** Review `XmlParsingService.ts`. Does it handle malformed XML gracefully?
- **Caching:** Review `CacheService.ts`. Is the TTL logic correct?

## 3. Test Coverage & Integrity
- **Mapping:** Verify that every major service in `src/services` has a corresponding test file in `tests/services`.
- **Mocking:** Check `OfficeMockFactory.ts`. Are we mocking the Office API correctly, or are we testing implementation details?
- **Gaps:** Identify any critical logic paths (like network failures) that are NOT currently tested.

## 4. Office Add-in Specifics
- **Manifest:** Validate `manifest.xml` for correct permissions (ReadWriteMailbox), icon paths, and URL schemes (HTTPS).
- **Webpack:** Ensure `webpack.config.js` correctly bundles assets and handles dev vs. prod modes.

## 5. Deployment Readiness (The "Fixed & Committed" Check)
- Check for any leftover `console.log`, `TODO`, or `FIXME` comments.
- detailed check for potential memory leaks (e.g., event listeners not removed).

# Output Format
Please provide a "Code Review Report" in Markdown with the following sections:

1.  **Executive Summary:** A 1-paragraph overview of the codebase health.
2.  **Critical Issues (Blockers):** Bugs or security risks that MUST be fixed before commit.
3.  **Test Coverage Gap Analysis:** A table showing Service Name | Has Test File | Missing Scenarios.
4.  **Refactoring Suggestions:** Improvements for readability or performance (non-blocking).
5.  **Manifest & Build Check:** Verification of XML and Webpack configs.
6.  **Conclusion:** A "Go/No-Go" decision for deployment.