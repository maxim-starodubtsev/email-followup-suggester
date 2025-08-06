/**
 * API Setup Examples for Followup Suggester
 * 
 * This file demonstrates how to configure different LLM providers
 * using the ConfigurationService helper methods.
 */

import { ConfigurationService } from '../src/services/ConfigurationService';

// Initialize the configuration service
const configService = new ConfigurationService();

/**
 * Example 1: Basic DIAL API Setup
 * Uses default settings with custom endpoint and API key
 */
export async function setupBasicDialApi() {
    console.log('Setting up DIAL API with defaults...');
    
    await configService.setupDialApi(
        'http://localhost:8080',  // Default DIAL endpoint
        'your-dial-api-key',      // Your API key
        'gpt-4o-mini'            // Default model
    );
    
    console.log('DIAL API configured successfully!');
}

/**
 * Example 2: Custom DIAL API Setup
 * Configure DIAL API with custom settings
 */
export async function setupCustomDialApi() {
    console.log('Setting up custom DIAL API...');
    
    await configService.setupDialApi(
        'http://my-dial-server.company.com:8080',  // Custom endpoint
        'my-company-api-key',                      // Company API key
        'gpt-4'                                    // Different model
    );
    
    console.log('Custom DIAL API configured successfully!');
}

/**
 * Example 3: Azure OpenAI Setup
 * Configure Azure OpenAI with deployment details
 */
export async function setupAzureOpenAI() {
    console.log('Setting up Azure OpenAI...');
    
    await configService.setupAzureOpenAi(
        'https://my-resource.openai.azure.com',    // Azure resource endpoint
        'azure-api-key-here',                      // Azure API key
        'gpt-4-deployment',                        // Deployment name
        '2024-02-01',                              // API version
        'gpt-4'                                    // Model name
    );
    
    console.log('Azure OpenAI configured successfully!');
}

/**
 * Example 4: Direct OpenAI Setup
 * Configure direct OpenAI API access
 */
export async function setupDirectOpenAI() {
    console.log('Setting up direct OpenAI...');
    
    await configService.setupOpenAi(
        'sk-your-openai-api-key-here',  // OpenAI API key
        'gpt-4',                        // Model name
        'https://api.openai.com'        // OpenAI endpoint (optional)
    );
    
    console.log('Direct OpenAI configured successfully!');
}

/**
 * Example 5: Check Current Configuration
 * Get and display current LLM configuration
 */
export async function checkCurrentConfig() {
    console.log('Checking current LLM configuration...');
    
    const config = await configService.getLlmConfiguration();
    
    console.log('Current Configuration:');
    console.log(`  Provider: ${config.provider}`);
    console.log(`  Endpoint: ${config.endpoint}`);
    console.log(`  Model: ${config.model}`);
    console.log(`  API Key: ${config.apiKey ? '***configured***' : 'not set'}`);
    console.log(`  Is Configured: ${config.isConfigured}`);
    
    if (config.provider === 'azure') {
        console.log(`  Deployment: ${config.deploymentName}`);
        console.log(`  API Version: ${config.apiVersion}`);
    }
}

/**
 * Example 6: Quick Setup Function
 * Easy setup based on provider type
 */
export async function quickSetup(provider: 'dial' | 'azure' | 'openai', options: any) {
    console.log(`Quick setup for ${provider}...`);
    
    switch (provider) {
        case 'dial':
            await configService.setupDialApi(
                options.endpoint || 'http://localhost:8080',
                options.apiKey,
                options.model || 'gpt-4o-mini'
            );
            break;
            
        case 'azure':
            await configService.setupAzureOpenAi(
                options.endpoint,
                options.apiKey,
                options.deploymentName,
                options.apiVersion || '2024-02-01',
                options.model
            );
            break;
            
        case 'openai':
            await configService.setupOpenAi(
                options.apiKey,
                options.model || 'gpt-4',
                options.endpoint
            );
            break;
    }
    
    console.log(`${provider.toUpperCase()} configured successfully!`);
}

/**
 * Example Usage in Add-in
 */
export async function exampleAddInUsage() {
    // Quick DIAL API setup for development
    await quickSetup('dial', {
        endpoint: 'http://localhost:8080',
        apiKey: 'dev-api-key',
        model: 'gpt-4o-mini'
    });
    
    // Enable AI features
    const config = await configService.getConfiguration();
    config.enableLlmSummary = true;
    config.enableLlmSuggestions = true;
    await configService.saveConfiguration(config);
    
    console.log('Add-in is now ready with AI features enabled!');
}

/**
 * Configuration Templates
 */
export const configTemplates = {
    // Development DIAL setup
    dialDev: {
        provider: 'dial',
        endpoint: 'http://localhost:8080',
        model: 'gpt-4o-mini',
        // apiKey: set this when calling
    },
    
    // Production DIAL setup
    dialProd: {
        provider: 'dial',
        endpoint: 'https://ai-proxy.company.com',
        model: 'gpt-4',
        // apiKey: set this when calling
    },
    
    // Azure OpenAI setup
    azure: {
        provider: 'azure',
        endpoint: 'https://company-ai.openai.azure.com',
        deploymentName: 'gpt-4-deployment',
        apiVersion: '2024-02-01',
        model: 'gpt-4',
        // apiKey: set this when calling
    }
};
