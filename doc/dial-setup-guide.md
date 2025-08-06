# DIAL API Setup Guide - Followup Suggester

## Quick Setup for DIAL API

### Default Configuration

The Followup Suggester comes with DIAL API as the default provider with these settings:

```typescript
// Default settings in ConfigurationService
llmProvider: 'dial'
llmApiEndpoint: 'http://localhost:8080'
llmModel: 'gpt-4o-mini'
llmApiKey: ''                    // You need to set this
llmApiVersion: '2024-02-01'
```

## Setup Methods

### 1. Using the Settings UI (Recommended)

1. **Open Outlook** and load the Followup Suggester add-in
2. **Click Settings** (gear icon) in the task pane
3. **Configure LLM Provider**:
   - **Provider**: Select "DIAL API" (default)
   - **Base URL**: Enter your DIAL server URL (e.g., `http://localhost:8080`)
   - **API Key**: Enter your DIAL API key
   - **Model**: Enter model name (e.g., `gpt-4o-mini`)
4. **Enable AI Features**:
   - ☑️ Use AI for email summaries
   - ☑️ Use AI for follow-up suggestions
5. **Click Save Settings**
6. **Test Connection** to verify it works

### 2. Programmatic Setup

```typescript
import { ConfigurationService } from './services/ConfigurationService';

const configService = new ConfigurationService();

// Basic setup with defaults
await configService.setupDialApi(
    'http://localhost:8080',     // Your DIAL server endpoint
    'your-dial-api-key',         // Your API key
    'gpt-4o-mini'               // Model name
);

// Enable AI features
const config = await configService.getConfiguration();
config.enableLlmSummary = true;
config.enableLlmSuggestions = true;
await configService.saveConfiguration(config);
```

### 3. Check Configuration

```typescript
// Verify current settings
const llmConfig = await configService.getLlmConfiguration();
console.log('Provider:', llmConfig.provider);        // 'dial'
console.log('Endpoint:', llmConfig.endpoint);        // Your endpoint
console.log('Model:', llmConfig.model);             // Your model
console.log('Configured:', llmConfig.isConfigured); // true if API key is set
```

## DIAL API Details

### Endpoints
- **Base URL**: Your DIAL server (e.g., `http://localhost:8080`)
- **Chat Completions**: `/v1/chat/completions`
- **Models**: `/v1/models`

### Authentication
- **Method**: Bearer token in Authorization header
- **Header**: `Authorization: Bearer <your-api-key>`

### Request Format
```json
{
  "model": "gpt-4o-mini",
  "messages": [
    {
      "role": "system",
      "content": "You are an email analysis assistant..."
    },
    {
      "role": "user",
      "content": "Analyze this email..."
    }
  ],
  "temperature": 0.3,
  "max_tokens": 1000
}
```

## Common DIAL Server URLs

- **Local Development**: `http://localhost:8080`
- **Docker**: `http://dial-server:8080`
- **Corporate**: `https://ai-proxy.company.com`

## Troubleshooting

### Connection Issues
1. **Verify URL**: Make sure the DIAL server is running and accessible
2. **Check API Key**: Ensure your API key is valid and has permissions
3. **Test Connection**: Use the "Test AI Connection" button in Settings
4. **Check Console**: Look for error messages in browser dev tools

### Model Issues
1. **List Available Models**: Check `/v1/models` endpoint on your DIAL server
2. **Model Name**: Ensure the model name exactly matches what's available
3. **Permissions**: Verify your API key has access to the specified model

### Feature Not Working
1. **Enable Features**: Make sure AI summaries and suggestions are enabled
2. **Check Status**: Look for AI status indicators in the UI
3. **Disable/Re-enable**: Try turning AI features off and on again

## Example Configurations

### Development Setup
```typescript
await configService.setupDialApi(
    'http://localhost:8080',
    'dev-api-key',
    'gpt-4o-mini'
);
```

### Production Setup
```typescript
await configService.setupDialApi(
    'https://ai-proxy.company.com',
    'prod-api-key-xxxxx',
    'gpt-4'
);
```

### Docker Setup
```typescript
await configService.setupDialApi(
    'http://dial-server:8080',
    'docker-api-key',
    'gpt-4o-mini'
);
```

## Next Steps

Once configured:
1. **Analyze Emails**: Run email analysis to see AI summaries
2. **Check Suggestions**: Look for AI-generated follow-up suggestions
3. **Monitor Performance**: Use the diagnostics to check AI response times
4. **Adjust Settings**: Fine-tune model and endpoint as needed

The DIAL API integration provides intelligent email analysis while keeping your data on your infrastructure!
