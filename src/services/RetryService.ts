// Retry service with comprehensive error recovery, exponential backoff, and circuit breaker pattern
export interface RetryOptions {
    maxAttempts?: number;
    baseDelayMs?: number;
    maxDelayMs?: number;
    backoffFactor?: number;
    jitterMs?: number;
    retryableErrors?: (error: Error) => boolean;
    onRetry?: (attempt: number, error: Error, nextDelayMs: number) => void;
}

export class RetryableError extends Error {
    public readonly isRetryable: boolean = true;
    public readonly shouldCircuitBreak: boolean = false;

    constructor(message: string, shouldCircuitBreak = false) {
        super(message);
        this.name = 'RetryableError';
        this.shouldCircuitBreak = shouldCircuitBreak;
    }
}

export class NonRetryableError extends Error {
    public readonly isRetryable: boolean = false;

    constructor(message: string) {
        super(message);
        this.name = 'NonRetryableError';
    }
}

export enum CircuitBreakerState {
    CLOSED = 'CLOSED',
    OPEN = 'OPEN',
    HALF_OPEN = 'HALF_OPEN'
}

export interface CircuitBreakerOptions {
    failureThreshold?: number;
    recoveryTimeoutMs?: number;
    monitoringPeriodMs?: number;
    halfOpenMaxAttempts?: number;
}

export class CircuitBreaker {
    private state: CircuitBreakerState = CircuitBreakerState.CLOSED;
    private failureCount = 0;
    private nextAttemptTime = 0;
    private halfOpenAttempts = 0;
    private readonly options: Required<CircuitBreakerOptions>;

    constructor(options: CircuitBreakerOptions = {}) {
        this.options = {
            failureThreshold: options.failureThreshold ?? 5,
            recoveryTimeoutMs: options.recoveryTimeoutMs ?? 60000, // 1 minute
            monitoringPeriodMs: options.monitoringPeriodMs ?? 10000, // 10 seconds
            halfOpenMaxAttempts: options.halfOpenMaxAttempts ?? 3
        };
    }

    public async execute<T>(operation: () => Promise<T>): Promise<T> {
        if (!this.canExecute()) {
            throw new Error(`Circuit breaker is ${this.state}. Next attempt allowed at ${new Date(this.nextAttemptTime).toISOString()}`);
        }

        try {
            const result = await operation();
            this.onSuccess();
            return result;
        } catch (error) {
            this.onFailure(error as Error);
            throw error;
        }
    }

    public getState(): CircuitBreakerState {
        return this.state;
    }

    public getFailureCount(): number {
        return this.failureCount;
    }

    public reset(): void {
        this.state = CircuitBreakerState.CLOSED;
        this.failureCount = 0;
        this.nextAttemptTime = 0;
        this.halfOpenAttempts = 0;
    }

    private canExecute(): boolean {
        const now = Date.now();

        switch (this.state) {
            case CircuitBreakerState.CLOSED:
                return true;

            case CircuitBreakerState.OPEN:
                if (now >= this.nextAttemptTime) {
                    this.state = CircuitBreakerState.HALF_OPEN;
                    this.halfOpenAttempts = 0;
                    return true;
                }
                return false;

            case CircuitBreakerState.HALF_OPEN:
                return this.halfOpenAttempts < this.options.halfOpenMaxAttempts;

            default:
                return false;
        }
    }

    private onSuccess(): void {
        this.failureCount = 0;
        this.halfOpenAttempts = 0;
        
        if (this.state === CircuitBreakerState.HALF_OPEN) {
            this.state = CircuitBreakerState.CLOSED;
        }
    }

    private onFailure(error: Error): void {
        this.failureCount++;

        const isRetryableError = error instanceof RetryableError;
        const shouldCircuitBreak = isRetryableError && error.shouldCircuitBreak;

        switch (this.state) {
            case CircuitBreakerState.CLOSED:
                if (shouldCircuitBreak || this.failureCount >= this.options.failureThreshold) {
                    this.openCircuit();
                }
                break;

            case CircuitBreakerState.HALF_OPEN:
                this.halfOpenAttempts++;
                this.openCircuit();
                break;
        }
    }

    private openCircuit(): void {
        this.state = CircuitBreakerState.OPEN;
        this.nextAttemptTime = Date.now() + this.options.recoveryTimeoutMs;
        this.halfOpenAttempts = 0;
    }
}

export interface RetryStats {
    totalAttempts: number;
    totalRetries: number;
    successfulRetries: number;
    failedRetries: number;
    averageRetryDelay: number;
    circuitBreakerTrips: number;
}

export class RetryService {
    private readonly defaultOptions: Required<RetryOptions>;
    private readonly circuitBreakers = new Map<string, CircuitBreaker>();
    private readonly stats: RetryStats = {
        totalAttempts: 0,
        totalRetries: 0,
        successfulRetries: 0,
        failedRetries: 0,
        averageRetryDelay: 0,
        circuitBreakerTrips: 0
    };
    private totalRetryDelay = 0;

    constructor(defaultOptions: RetryOptions = {}) {
        this.defaultOptions = {
            maxAttempts: defaultOptions.maxAttempts ?? 3,
            baseDelayMs: defaultOptions.baseDelayMs ?? 1000,
            maxDelayMs: defaultOptions.maxDelayMs ?? 30000,
            backoffFactor: defaultOptions.backoffFactor ?? 2,
            jitterMs: defaultOptions.jitterMs ?? 100,
            retryableErrors: defaultOptions.retryableErrors ?? this.defaultRetryableErrorCheck,
            onRetry: defaultOptions.onRetry ?? (() => {})
        };
    }

    public async executeWithRetry<T>(
        operation: () => Promise<T>,
        options: RetryOptions = {},
        circuitBreakerKey?: string
    ): Promise<T> {
        const mergedOptions = this.mergeOptions(options);
        const circuitBreaker = circuitBreakerKey ? this.getOrCreateCircuitBreaker(circuitBreakerKey) : null;

        this.stats.totalAttempts++;

        let lastError: Error;
        let attempt = 0;

        while (attempt < mergedOptions.maxAttempts) {
            try {
                if (circuitBreaker) {
                    return await circuitBreaker.execute(operation);
                } else {
                    return await operation();
                }
            } catch (error) {
                lastError = error as Error;
                attempt++;

                // Check if error is retryable
                if (!mergedOptions.retryableErrors(lastError)) {
                    throw lastError;
                }

                // Don't retry on last attempt
                if (attempt >= mergedOptions.maxAttempts) {
                    this.stats.failedRetries++;
                    break;
                }

                // Calculate delay with exponential backoff and jitter
                const baseDelay = Math.min(
                    mergedOptions.baseDelayMs * Math.pow(mergedOptions.backoffFactor, attempt - 1),
                    mergedOptions.maxDelayMs
                );
                const jitter = Math.random() * mergedOptions.jitterMs;
                const delayMs = Math.floor(baseDelay + jitter);

                this.stats.totalRetries++;
                this.totalRetryDelay += delayMs;
                this.stats.averageRetryDelay = this.totalRetryDelay / this.stats.totalRetries;

                // Call retry callback
                mergedOptions.onRetry(attempt, lastError, delayMs);

                // Wait before retry
                await this.delay(delayMs);
            }
        }

        throw lastError!;
    }

    public async executeWithExponentialBackoff<T>(
        operation: () => Promise<T>,
        maxAttempts = 3,
        baseDelayMs = 1000,
        circuitBreakerKey?: string
    ): Promise<T> {
        return this.executeWithRetry(
            operation,
            {
                maxAttempts,
                baseDelayMs,
                backoffFactor: 2,
                maxDelayMs: 30000,
                jitterMs: 100
            },
            circuitBreakerKey
        );
    }

    public createCircuitBreaker(key: string, options: CircuitBreakerOptions = {}): CircuitBreaker {
        const circuitBreaker = new CircuitBreaker(options);
        this.circuitBreakers.set(key, circuitBreaker);
        return circuitBreaker;
    }

    public getCircuitBreaker(key: string): CircuitBreaker | undefined {
        return this.circuitBreakers.get(key);
    }

    public resetCircuitBreaker(key: string): boolean {
        const circuitBreaker = this.circuitBreakers.get(key);
        if (circuitBreaker) {
            circuitBreaker.reset();
            return true;
        }
        return false;
    }

    public resetAllCircuitBreakers(): void {
        this.circuitBreakers.forEach(cb => cb.reset());
        this.stats.circuitBreakerTrips = 0;
    }

    public getStats(): RetryStats {
        return { ...this.stats };
    }

    public resetStats(): void {
        this.stats.totalAttempts = 0;
        this.stats.totalRetries = 0;
        this.stats.successfulRetries = 0;
        this.stats.failedRetries = 0;
        this.stats.averageRetryDelay = 0;
        this.stats.circuitBreakerTrips = 0;
        this.totalRetryDelay = 0;
    }

    public getCircuitBreakerStates(): Record<string, CircuitBreakerState> {
        const states: Record<string, CircuitBreakerState> = {};
        this.circuitBreakers.forEach((cb, key) => {
            states[key] = cb.getState();
        });
        return states;
    }

    // Static utility methods for common retry scenarios
    public static async retryOnRateLimit<T>(
        operation: () => Promise<T>,
        maxAttempts = 3,
        baseDelayMs = 5000
    ): Promise<T> {
        const retryService = new RetryService({
            maxAttempts,
            baseDelayMs,
            backoffFactor: 2,
            maxDelayMs: 60000,
            retryableErrors: (error: Error) => {
                return error.message.includes('rate limit') || 
                       error.message.includes('429') ||
                       error.message.includes('quota exceeded');
            }
        });

        return retryService.executeWithRetry(operation);
    }

    public static async retryOnNetworkError<T>(
        operation: () => Promise<T>,
        maxAttempts = 3,
        baseDelayMs = 1000
    ): Promise<T> {
        const retryService = new RetryService({
            maxAttempts,
            baseDelayMs,
            backoffFactor: 2,
            maxDelayMs: 30000,
            retryableErrors: (error: Error) => {
                const message = error.message.toLowerCase();
                return message.includes('network') ||
                       message.includes('timeout') ||
                       message.includes('connection') ||
                       message.includes('socket') ||
                       message.includes('econnreset') ||
                       message.includes('enotfound');
            }
        });

        return retryService.executeWithRetry(operation);
    }

    private mergeOptions(options: RetryOptions): Required<RetryOptions> {
        return {
            maxAttempts: options.maxAttempts ?? this.defaultOptions.maxAttempts,
            baseDelayMs: options.baseDelayMs ?? this.defaultOptions.baseDelayMs,
            maxDelayMs: options.maxDelayMs ?? this.defaultOptions.maxDelayMs,
            backoffFactor: options.backoffFactor ?? this.defaultOptions.backoffFactor,
            jitterMs: options.jitterMs ?? this.defaultOptions.jitterMs,
            retryableErrors: options.retryableErrors ?? this.defaultOptions.retryableErrors,
            onRetry: options.onRetry ?? this.defaultOptions.onRetry
        };
    }

    private getOrCreateCircuitBreaker(key: string): CircuitBreaker {
        let circuitBreaker = this.circuitBreakers.get(key);
        if (!circuitBreaker) {
            circuitBreaker = new CircuitBreaker();
            this.circuitBreakers.set(key, circuitBreaker);
        }
        return circuitBreaker;
    }

    private defaultRetryableErrorCheck(error: Error): boolean {
        if (error instanceof NonRetryableError) {
            return false;
        }

        if (error instanceof RetryableError) {
            return true;
        }

        // Default logic for determining retryable errors
        const message = error.message.toLowerCase();
        
        // Network-related errors are generally retryable
        if (message.includes('network') || 
            message.includes('timeout') || 
            message.includes('connection') ||
            message.includes('socket') ||
            message.includes('econnreset') ||
            message.includes('enotfound')) {
            return true;
        }

        // Rate limiting is retryable
        if (message.includes('rate limit') || 
            message.includes('429') ||
            message.includes('quota exceeded')) {
            return true;
        }

        // Server errors (5xx) are generally retryable
        if (message.includes('500') ||
            message.includes('502') ||
            message.includes('503') ||
            message.includes('504') ||
            message.includes('internal server error') ||
            message.includes('bad gateway') ||
            message.includes('service unavailable') ||
            message.includes('gateway timeout')) {
            return true;
        }

        // Client errors (4xx except 429) are generally not retryable
        if (message.includes('400') ||
            message.includes('401') ||
            message.includes('403') ||
            message.includes('404') ||
            message.includes('bad request') ||
            message.includes('unauthorized') ||
            message.includes('forbidden') ||
            message.includes('not found')) {
            return false;
        }

        // Default to retryable for unknown errors
        return true;
    }

    private async delay(ms: number): Promise<void> {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
}