import { defineConfig } from 'vitest/config';
import path from 'path';

export default defineConfig({
  test: {
    globals: true,
    environment: 'happy-dom',
    setupFiles: ['./tests/setup.ts'],
    include: ['tests/**/*.test.ts'],
    exclude: ['node_modules/', 'dist/', 'debug/'],
    pool: 'vmForks',
    poolOptions: {
      vmForks: {
        singleFork: true,
        execArgv: ['--max-old-space-size=4096']
      }
    },
    coverage: {
      provider: 'v8',
      reporter: ['text', 'json', 'html', 'lcov'],
      exclude: [
        'node_modules/',
        'dist/',
        'tests/',
        'debug/',
        '**/*.d.ts',
        '**/*.config.*',
        'verify-build.js',
        'webpack.config.js'
      ],
      all: true,
      lines: 80,
      functions: 80,
      branches: 80,
      statements: 80
    },
    testTimeout: 10000,
    hookTimeout: 10000
  },
  resolve: {
    alias: {
      '@': path.resolve(__dirname, './src')
    }
  }
});

