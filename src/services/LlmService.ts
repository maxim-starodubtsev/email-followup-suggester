import { Configuration } from '../models/Configuration';
import { RetryService, RetryOptions } from './RetryService';

export interface LlmOptions {
    temperature?: number;
    maxTokens?: number;
    topP?: number;
    frequencyPenalty?: number;
    presencePenalty?: number;
}

export interface LlmResponse {
    content: string;
    usage?: {
        promptTokens: number;
        completionTokens: number;
        totalTokens: number;
    };
    model?: string;
    finishReason?: string;
}

export class LlmService {
    private retryService: RetryService;
    private configuration: Configuration;

    constructor(configuration: Configuration, retryService: RetryService) {
        this.configuration = configuration;
        this.retryService = retryService;
    }

    public async generateFollowupSuggestions(
        emailContent: string, 
        context?: string,
        options: LlmOptions = {}
    ): Promise<string[]> {
        const prompt = this.buildFollowupPrompt(emailContent, context);
        
        try {
            const response = await this.callLlmApi(prompt, options);
            return this.parseFollowupSuggestions(response.content);
        } catch (error: any) {
            console.error('Error generating followup suggestions:', error);
            throw new Error(`Failed to generate followup suggestions: ${error.message}`);
        }
    }

    public async analyzeTone(emailContent: string, options: LlmOptions = {}): Promise<string> {
        const prompt = `Analyze the tone of the following email and provide a brief description:

${emailContent}

Provide a concise tone analysis in one sentence.`;

        try {
            const response = await this.callLlmApi(prompt, options);
            return response.content.trim();
        } catch (error: any) {
            console.error('Error analyzing tone:', error);
            throw new Error(`Failed to analyze tone: ${error.message}`);
        }
    }

    public async summarizeEmail(emailContent: string, options: LlmOptions = {}): Promise<string> {
        const prompt = `Summarize the following email concisely:

${emailContent}

Provide a brief summary in 2-3 sentences.`;

        try {
            const response = await this.callLlmApi(prompt, options);
            return response.content.trim();
        } catch (error: any) {
            console.error('Error summarizing email:', error);
            throw new Error(`Failed to summarize email: ${error.message}`);
        }
    }

    public async analyzeThread(params: { emails: any[]; context?: string }, options: LlmOptions = {}): Promise<string> {
        const emailsText = params.emails.map((email, index) => 
            `Email ${index + 1}:\n${email.subject}\n${email.content}\n---`
        ).join('\n\n');
        
        const prompt = `Analyze the following email thread and provide insights:

${emailsText}

${params.context ? `Additional context: ${params.context}` : ''}

Provide a brief analysis of the thread including key points, sentiment, and suggested actions.`;

        try {
            const response = await this.callLlmApi(prompt, options);
            return response.content.trim();
        } catch (error: any) {
            console.error('Error analyzing thread:', error);
            throw new Error(`Failed to analyze thread: ${error.message}`);
        }
    }

    public async analyzeSentiment(emailContent: string, options: LlmOptions = {}): Promise<{ sentiment: 'positive' | 'neutral' | 'negative' | 'urgent' }> {
        const prompt = `Analyze the sentiment and urgency of the following email content:

${emailContent}

Return one of these sentiment classifications: positive, neutral, negative, urgent

Just return the classification word, nothing else.`;

        try {
            const response = await this.callLlmApi(prompt, options);
            const content = response.content.trim().toLowerCase();
            
            if (content.includes('urgent')) {
                return { sentiment: 'urgent' };
            } else if (content.includes('positive')) {
                return { sentiment: 'positive' };
            } else if (content.includes('negative')) {
                return { sentiment: 'negative' };
            } else {
                return { sentiment: 'neutral' };
            }
        } catch (error: any) {
            console.error('Error analyzing sentiment:', error);
            return { sentiment: 'neutral' };
        }
    }

    private async callLlmApi(prompt: string, options: LlmOptions = {}): Promise<LlmResponse> {
        const retryOptions: RetryOptions = {
            maxAttempts: 3,
            baseDelayMs: 1000,
            maxDelayMs: 10000,
            backoffFactor: 2
        };

        return this.retryService.executeWithRetry(
            () => this.makeApiCall(prompt, options),
            retryOptions,
            'llm-api'
        );
    }

    private async callDialAPI(prompt: string, options: LlmOptions = {}): Promise<LlmResponse> {
        const requestBody = {
            model: this.configuration.llmModel || 'gpt-35-turbo',
            messages: [
                { role: 'user', content: prompt }
            ],
            temperature: options.temperature ?? 0.7,
            max_tokens: options.maxTokens ?? 500,
            top_p: options.topP ?? 1.0,
            frequency_penalty: options.frequencyPenalty ?? 0,
            presence_penalty: options.presencePenalty ?? 0
        };

        const headers: Record<string, string> = {
            'Content-Type': 'application/json'
        };

        if (this.configuration.llmApiKey) {
            headers['Api-Key'] = this.configuration.llmApiKey;
        }

        // Use the endpoint URL directly if it already contains the completions path
        const apiUrl = this.configuration.llmApiEndpoint!.includes('/openai/chat/completions') 
            ? this.configuration.llmApiEndpoint!
            : `${this.configuration.llmApiEndpoint}/openai/chat/completions`;

        const response = await fetch(apiUrl, {
            method: 'POST',
            headers,
            body: JSON.stringify(requestBody)
        });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`DIAL API request failed: ${response.status} ${response.statusText} - ${errorText}`);
        }

        const data = await response.json();
        
        if (!data.choices || data.choices.length === 0) {
            throw new Error('No response choices returned from DIAL API');
        }

        return {
            content: data.choices[0].message.content,
            usage: data.usage ? {
                promptTokens: data.usage.prompt_tokens,
                completionTokens: data.usage.completion_tokens,
                totalTokens: data.usage.total_tokens
            } : undefined,
            model: data.model,
            finishReason: data.choices[0].finish_reason
        };
    }

    private async makeApiCall(prompt: string, options: LlmOptions = {}): Promise<LlmResponse> {
        // Check if API endpoint is configured
        if (!this.configuration.llmApiEndpoint) {
            throw new Error('LLM API endpoint is not configured');
        }

        // Use DIAL API if the endpoint contains the DIAL pattern
        if (this.configuration.llmApiEndpoint.includes('ai-proxy.lab.epam.com') || 
            this.configuration.llmApiEndpoint.includes('/openai/chat/completions')) {
            return this.callDialAPI(prompt, options);
        }

        // Fallback to Azure OpenAI format
        const requestBody = {
            messages: [
                { role: 'user', content: prompt }
            ],
            temperature: options.temperature ?? 0.7,
            max_tokens: options.maxTokens ?? 500,
            top_p: options.topP ?? 1.0,
            frequency_penalty: options.frequencyPenalty ?? 0,
            presence_penalty: options.presencePenalty ?? 0
        };

        const model = this.configuration.llmModel || 'gpt-35-turbo';
        const response = await fetch(`${this.configuration.llmApiEndpoint}/openai/deployments/${model}/chat/completions?api-version=2023-12-01-preview`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${this.configuration.llmApiKey}`
            },
            body: JSON.stringify(requestBody)
        });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Azure OpenAI API request failed: ${response.status} ${response.statusText} - ${errorText}`);
        }

        const data = await response.json();
        
        if (!data.choices || data.choices.length === 0) {
            throw new Error('No response choices returned from Azure OpenAI API');
        }

        return {
            content: data.choices[0].message.content,
            usage: data.usage ? {
                promptTokens: data.usage.prompt_tokens,
                completionTokens: data.usage.completion_tokens,
                totalTokens: data.usage.total_tokens
            } : undefined,
            model: data.model,
            finishReason: data.choices[0].finish_reason
        };
    }

    private buildFollowupPrompt(emailContent: string, context?: string): string {
        let prompt = `Based on the following email, suggest 3-5 professional followup responses that would be appropriate:

Email content:
${emailContent}`;

        if (context) {
            prompt += `\n\nAdditional context:
${context}`;
        }

        prompt += `\n\nPlease provide followup suggestions in the following format:
1. [First suggestion]
2. [Second suggestion]
3. [Third suggestion]
[etc.]

Focus on being professional, helpful, and contextually appropriate.`;

        return prompt;
    }

    private parseFollowupSuggestions(response: string): string[] {
        const lines = response.split('\n').filter(line => line.trim());
        const suggestions: string[] = [];

        for (const line of lines) {
            const match = line.match(/^\d+\.\s*(.+)$/);
            if (match) {
                suggestions.push(match[1].trim());
            }
        }

        return suggestions.length > 0 ? suggestions : [response.trim()];
    }
}