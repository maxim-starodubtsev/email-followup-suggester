export interface BatchProcessingOptions {
    batchSize: number;
    maxConcurrentBatches: number;
    retryOptions: {
        maxRetries: number;
        baseDelay: number;
        maxDelay: number;
    };
    onProgress?: (current: number, total: number, currentBatch: number, totalBatches: number) => void;
    onBatchComplete?: (batchIndex: number, results: any[], errors: Error[]) => void;
    onBatchError?: (batchIndex: number, error: Error) => void;
}

export interface BatchResult<T> {
    success: boolean;
    results: T[];
    errors: Error[];
    totalProcessed: number;
    totalBatches: number;
    processingTime: number;
    cancelled: boolean;
}

export class BatchProcessor {
    private readonly DEFAULT_OPTIONS: BatchProcessingOptions = {
        batchSize: 10,
        maxConcurrentBatches: 3,
        retryOptions: {
            maxRetries: 3,
            baseDelay: 1000,
            maxDelay: 10000
        }
    };

    private cancellationTokens: Map<string, boolean> = new Map();
    private processingTasks: Map<string, Promise<any>> = new Map();

    public async processBatch<TInput, TOutput>(
        items: TInput[],
        processor: (item: TInput) => Promise<TOutput>,
        options: Partial<BatchProcessingOptions> = {}
    ): Promise<BatchResult<TOutput>> {
        const processingId = this.generateProcessingId();
        const mergedOptions: BatchProcessingOptions = {
            ...this.DEFAULT_OPTIONS,
            ...options,
            retryOptions: {
                ...this.DEFAULT_OPTIONS.retryOptions,
                ...(options.retryOptions ?? {})
            }
        };
        const startTime = Date.now();

        // Initialize cancellation token
        this.cancellationTokens.set(processingId, false);

        try {
            const batches = this.createBatches(items, mergedOptions.batchSize);
            const results: TOutput[] = [];
            const errors: Error[] = [];
            let processedCount = 0;

            // Track progress
            this.reportProgress(0, items.length, 0, batches.length, mergedOptions.onProgress);

            // Create array to maintain order of results
            const batchResults: Array<{ index: number; results: TOutput[]; errors: Error[] }> = [];

            // Process batches with concurrency control
            const batchPromises: Promise<void>[] = [];
            const semaphore = new Semaphore(mergedOptions.maxConcurrentBatches);

            for (let batchIndex = 0; batchIndex < batches.length; batchIndex++) {
                const batchPromise = semaphore.acquire().then(async (release) => {
                    try {
                        // Check for cancellation
                        if (this.cancellationTokens.get(processingId)) {
                            return;
                        }

                        const batch = batches[batchIndex];
                        const currentBatchResults: TOutput[] = [];
                        const currentBatchErrors: Error[] = [];

                        // Process items in current batch sequentially for error isolation
                        for (const item of batch) {
                            if (this.cancellationTokens.get(processingId)) {
                                break;
                            }

                            try {
                                const result = await this.withRetry(
                                    () => processor(item),
                                    mergedOptions.retryOptions
                                );
                                currentBatchResults.push(result);
                                processedCount++;
                            } catch (error) {
                                const batchError = new Error(`Batch ${batchIndex} item error: ${(error as Error).message}`);
                                currentBatchErrors.push(batchError);
                                
                                if (mergedOptions.onBatchError) {
                                    mergedOptions.onBatchError(batchIndex, batchError);
                                }
                            }

                            // Report progress after each item
                            this.reportProgress(processedCount, items.length, batchIndex + 1, batches.length, mergedOptions.onProgress);
                        }

                        // Store results with batch index to maintain order
                        batchResults[batchIndex] = {
                            index: batchIndex,
                            results: currentBatchResults,
                            errors: currentBatchErrors
                        };

                        // Notify batch completion
                        if (mergedOptions.onBatchComplete) {
                            mergedOptions.onBatchComplete(batchIndex, currentBatchResults, currentBatchErrors);
                        }

                    } finally {
                        release();
                    }
                });

                batchPromises.push(batchPromise);
            }

            // Wait for all batches to complete
            await Promise.all(batchPromises);

            // Combine results in order
            for (let i = 0; i < batchResults.length; i++) {
                if (batchResults[i]) {
                    results.push(...batchResults[i].results);
                    errors.push(...batchResults[i].errors);
                }
            }

            const processingTime = Date.now() - startTime;
            const cancelled = this.cancellationTokens.get(processingId) || false;

            return {
                success: errors.length === 0,
                results,
                errors,
                totalProcessed: processedCount,
                totalBatches: batches.length,
                processingTime,
                cancelled
            };

        } finally {
            // Cleanup
            this.cancellationTokens.delete(processingId);
            this.processingTasks.delete(processingId);
        }
    }

    public cancelProcessing(processingId?: string): void {
        if (processingId) {
            this.cancellationTokens.set(processingId, true);
        } else {
            // Cancel all processing tasks
            for (const [id] of this.cancellationTokens) {
                this.cancellationTokens.set(id, true);
            }
        }
    }

    public getActiveProcessingTasks(): string[] {
        return Array.from(this.cancellationTokens.keys()).filter(id => !this.cancellationTokens.get(id));
    }

    private createBatches<T>(items: T[], batchSize: number): T[][] {
        const batches: T[][] = [];
        for (let i = 0; i < items.length; i += batchSize) {
            batches.push(items.slice(i, i + batchSize));
        }
        return batches;
    }

    private async withRetry<T>(
        operation: () => Promise<T>,
        retryOptions: BatchProcessingOptions['retryOptions']
    ): Promise<T> {
        let lastError: Error | null = null;
        
        for (let attempt = 0; attempt <= retryOptions.maxRetries; attempt++) {
            try {
                return await operation();
            } catch (error) {
                lastError = error as Error;
                
                if (attempt === retryOptions.maxRetries) {
                    break;
                }
                
                const delay = Math.min(
                    retryOptions.baseDelay * Math.pow(2, attempt),
                    retryOptions.maxDelay
                );
                
                await new Promise(resolve => setTimeout(resolve, delay));
            }
        }
        
        throw lastError || new Error('Operation failed after all retries');
    }

    private reportProgress(
        current: number,
        total: number,
        currentBatch: number,
        totalBatches: number,
        onProgress?: (current: number, total: number, currentBatch: number, totalBatches: number) => void
    ): void {
        if (onProgress) {
            onProgress(current, total, currentBatch, totalBatches);
        }
    }

    private generateProcessingId(): string {
        return `batch_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    }
}

// Semaphore class for concurrency control
class Semaphore {
    private permits: number;
    private waiting: Array<() => void> = [];

    constructor(permits: number) {
        this.permits = permits;
    }

    async acquire(): Promise<() => void> {
        return new Promise((resolve) => {
            if (this.permits > 0) {
                this.permits--;
                resolve(() => this.release());
            } else {
                this.waiting.push(() => {
                    this.permits--;
                    resolve(() => this.release());
                });
            }
        });
    }

    private release(): void {
        this.permits++;
        if (this.waiting.length > 0) {
            const next = this.waiting.shift();
            if (next) {
                next();
            }
        }
    }
}