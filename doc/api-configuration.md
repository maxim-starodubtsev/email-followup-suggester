# API Configuration Guide - Followup Suggester

This guide explains how to configure and use the AI API integration in the Followup Suggester Outlook Add-in.

## 🤖 DIAL API Integration

The Followup Suggester uses the DIAL API (Distributed Inference for AI Language models) to provide intelligent email analysis, summaries, and follow-up suggestions.

### Default Configuration

The add-in comes pre-configured with the following DIAL API settings:

```javascript
// Pre-configured DIAL API settings
const DEFAULT_CONFIG = {
  apiEndpoint: "https://ai-proxy.lab.epam.com",
  apiKey: "dial-qjboupii21tb26eakd3ytcsb9po",
  modelName: "gpt-4o-mini-2024-07-18"
};
```

### API Endpoints

#### Primary Endpoint
- **Base URL**: `https://ai-proxy.lab.epam.com`
- **Chat Completions**: `/openai/chat/completions`
- **Models List**: `/openai/models`
- **Model Limits**: `/v1/deployments/{model}/limits`

#### Authentication
- **Method**: API Key in header
- **Header**: `Api-Key: dial-qjboupii21tb26eakd3ytcsb9po`
- **Content-Type**: `application/json`

## 🔧 Configuration Methods

### Method 1: Using the Settings UI

1. **Open the Add-in** in Outlook
2. **Click the Settings button** (gear icon)
3. **Configure AI Integration section**:
   - **LLM API Endpoint**: Enter your DIAL API endpoint URL
   - **API Key**: Enter your authentication key
4. **Enable AI Features**:
   - ☑️ Use AI for email summaries
   - ☑️ Use AI for follow-up suggestions
5. **Click Save Settings**

### Method 2: Direct Code Configuration

Update the configuration in `src/services/LlmService.ts`:

```typescript
constructor(apiEndpoint?: string, apiKey?: string, modelName?: string) {
    this.apiEndpoint = apiEndpoint || "https://your-dial-endpoint.com";
    this.apiKey = apiKey || "your-api-key-here";
    this.modelName = modelName || "gpt-4o-mini-2024-07-18";
}
```

### Method 3: Environment Variables

For development, you can use environment variables (requires build system modification):

```bash
# Add to your .env file
DIAL_API_ENDPOINT=https://ai-proxy.lab.epam.com
DIAL_API_KEY=dial-qjboupii21tb26eakd3ytcsb9po
DIAL_MODEL_NAME=gpt-4o-mini-2024-07-18
```

## 📋 API Request Format

### Chat Completion Request

```javascript
const requestBody = {
  model: "gpt-4o-mini-2024-07-18",
  temperature: 0.3,
  max_tokens: 1000,
  messages: [
    {
      role: "system",
      content: "You are an email analysis assistant..."
    },
    {
      role: "user", 
      content: "Analyze this email thread..."
    }
  ]
};

const response = await fetch("https://ai-proxy.lab.epam.com/openai/chat/completions", {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
    'Api-Key': 'dial-qjboupii21tb26eakd3ytcsb9po'
  },
  body: JSON.stringify(requestBody)
});
```

### Expected Response Format

```javascript
{
  "choices": [
    {
      "message": {
        "content": "{\"summary\": \"Brief email summary\", \"suggestion\": \"Follow-up action\", \"confidence\": 0.85}"
      }
    }
  ]
}
```

## 🎯 AI Analysis Features

### Email Summary Generation

The AI analyzes email content and generates concise summaries:

```javascript
// Input: Email thread messages
const threadContext = formatThreadForAnalysis(messages, currentUserEmail);

// Output: Intelligent summary
{
  "summary": "Discussion about project timeline with client concerns about delivery date",
  "suggestion": "Follow up with revised timeline and mitigation plan",
  "confidence": 0.92
}
```

### Follow-up Suggestions

Context-aware suggestions based on email content and conversation history:

- **Gentle reminders** for routine follow-ups
- **Urgent follow-ups** for time-sensitive matters
- **Informational updates** for status changes
- **Action requests** for pending decisions

### Priority Classification Enhancement

AI analysis enhances priority calculation by considering:
- **Content urgency** indicators
- **Recipient importance** based on domain/title
- **Response patterns** in conversation history
- **Time sensitivity** keywords and phrases

## 🔒 Security and Privacy

### API Key Security
- ✅ **Encrypted transmission**: All API calls use HTTPS
- ✅ **No local storage**: API keys stored in Office.js roaming settings
- ✅ **Environment isolation**: Development vs production keys
- ⚠️ **Key rotation**: Regularly update API keys

### Data Privacy
- **Email content**: Sent to DIAL API for analysis only
- **No persistent storage**: Content not stored on external servers
- **User control**: AI features can be disabled
- **Minimal data**: Only necessary content sent for analysis

### Compliance Considerations
- **Data residency**: Check DIAL API data location policies
- **GDPR compliance**: Ensure proper consent for AI processing
- **Corporate policies**: Verify AI usage approval
- **Audit trails**: API calls logged for monitoring

## 🚨 Error Handling

### Common Error Scenarios

#### 1. API Authentication Errors
```javascript
// Error: 401 Unauthorized
{
  "error": "Invalid API key"
}
```
**Solution**: Verify API key in settings

#### 2. Rate Limiting
```javascript
// Error: 429 Too Many Requests
{
  "error": "Rate limit exceeded"
}
```
**Solution**: Implement exponential backoff, reduce request frequency

#### 3. Model Unavailable
```javascript
// Error: 404 Not Found
{
  "error": "Model not found"
}
```
**Solution**: Update model name to available version

#### 4. Network Connectivity
```javascript
// Error: Network timeout
{
  "error": "Request timeout"
}
```
**Solution**: Automatic retry with fallback to non-AI analysis

### Fallback Strategies

The application implements graceful degradation:

1. **Primary**: DIAL API analysis
2. **Fallback**: Local text processing
3. **Ultimate**: Basic email listing without AI features

## 📊 Performance Optimization

### Request Optimization
- **Content truncation**: Limit email body to 2000 characters
- **Selective analysis**: Only analyze relevant email parts
- **Batch processing**: Future enhancement for multiple emails
- **Caching**: Store results to avoid duplicate API calls

### Response Time Targets
- **API Response**: < 3 seconds typical
- **UI Feedback**: Loading states for > 1 second operations
- **Timeout**: 10 second maximum wait time
- **Retry Logic**: 3 attempts with exponential backoff

## 🔄 Model Management

### Available Models
To get current available models:

```javascript
const models = await llmService.getAvailableModels();
console.log(models); // ["gpt-4o-mini-2024-07-18", "gpt-4", ...]
```

### Model Selection Criteria
- **Speed vs Quality**: Mini models for fast responses
- **Token limits**: Consider input/output token constraints
- **Cost optimization**: Balance performance vs API costs
- **Feature support**: Ensure JSON output capability

### Model Configuration
```javascript
// Switch models dynamically
llmService.setModel("gpt-4o-mini-2024-07-18");

// Check model limits
const limits = await llmService.checkModelLimits("gpt-4o-mini-2024-07-18");
```

## 🧪 Testing API Integration

### Unit Testing
```javascript
// Test API connectivity
const testRequest = {
  threadMessages: [/* sample data */],
  currentUserEmail: "test@example.com"
};

const result = await llmService.analyzeThread(testRequest);
console.log(result); // Verify response format
```

### Integration Testing
1. **Enable AI features** in settings
2. **Analyze test emails** with known content
3. **Verify summaries** are generated correctly
4. **Test error scenarios** with invalid API keys
5. **Monitor performance** under different loads

### Production Monitoring
- **Response time tracking**: Monitor API performance
- **Error rate monitoring**: Track failed requests
- **Token usage**: Monitor API consumption
- **User feedback**: Collect AI accuracy feedback

## 📈 Best Practices

### Development
1. **API key management**: Use environment-specific keys
2. **Error handling**: Comprehensive try-catch blocks
3. **User feedback**: Clear status messages
4. **Graceful degradation**: Fallback to non-AI features

### Production
1. **Key rotation**: Regular API key updates
2. **Monitoring**: Track API usage and errors
3. **User control**: Allow AI feature toggle
4. **Performance**: Monitor response times

### Security
1. **HTTPS only**: Never use HTTP for API calls
2. **Key protection**: Secure storage of credentials
3. **Input validation**: Sanitize data before API calls
4. **Audit logging**: Track API usage for compliance

This configuration guide provides everything needed to successfully integrate and manage the DIAL API in the Followup Suggester add-in.