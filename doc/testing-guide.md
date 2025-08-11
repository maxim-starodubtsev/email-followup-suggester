# Testing Guide

## Overview
This project uses Jest for unit testing with comprehensive test coverage for all service classes.

## Test Structure
```
tests/
├── setup.ts              # Test configuration
└── services/
    ├── ConfigurationService.test.ts
    ├── EmailAnalysisService.test.ts
    └── LlmService.test.ts
```

## Running Tests

### All Tests
```powershell
npm test
```

### Watch Mode
```powershell
npm run test:watch
```

### Coverage Report
```powershell
npm run test:coverage
```

## Test Categories

### EmailAnalysisService Tests
- **Priority Calculation**: Tests email aging logic (high/medium/low priority)
- **Response Detection**: Validates thread analysis for responses
- **LLM Integration**: Tests AI-powered email analysis
- **Email Filtering**: Verifies sent email processing

Key test scenarios:
- Emails >7 days = high priority
- Emails 3-6 days = medium priority
- Emails <3 days = low priority
- Response detection in email threads
- Artificial threading fallback with strict oldest→newest containment chain
    - Multi-hop chain A→B→C where each quotes the previous and C is from external sender (suppresses follow-up)
    - Broken chain where C does not contain B (chain ends at B for decision)
    - Short/noisy bodies are skipped but valid chain still formed

Notes:
- Normalized-body minimum length threshold is 20 characters to reduce false matches.
- Telemetry events for artificial threading can be asserted indirectly via behavior and logs in tests.

### LlmService Tests
- **Configuration Management**: API endpoint and model settings
- **Thread Analysis**: AI processing of email conversations
- **Model Validation**: Available models and limits
- **Error Handling**: Network failures and API errors

Key features tested:
- Custom API configuration
- Thread summarization
- Model switching (GPT-4, Claude, etc.)
- Rate limiting and error recovery

### ConfigurationService Tests
- **Settings Persistence**: Save/load user preferences
- **Validation**: Input validation and defaults
- **Migration**: Configuration updates between versions

## Mock Objects
Tests use Jest mocks for:
- LLM API calls
- Outlook Office.js APIs
- Configuration storage
- Network requests

## Test Data
Standard test email threads and configurations are defined in each test file for consistent testing scenarios.

## Best Practices
- Each service has isolated unit tests
- Mocked dependencies prevent external API calls
- Test both success and error scenarios
- Verify edge cases (empty threads, invalid configs)