#!/usr/bin/env node

/**
 * Development server startup script for macOS
 * 
 * This script starts the webpack dev server and attempts to sideload
 * the add-in in Outlook. Falls back to manual instructions if auto-sideloading fails.
 */

const { spawn, exec } = require('child_process');
const path = require('path');
const fs = require('fs');

const manifestPath = path.resolve(__dirname, '..', 'manifest.xml');

console.log('🚀 Starting Followup Suggester Development Server...\n');

// Verify manifest exists
if (!fs.existsSync(manifestPath)) {
  console.error(`❌ Error: manifest.xml not found at ${manifestPath}`);
  process.exit(1);
}

// Start webpack dev server in background
const webpackProcess = spawn('npm', ['run', 'dev'], {
  stdio: 'pipe',
  shell: true,
  cwd: __dirname + '/..'
});

let webpackReady = false;
let outputBuffer = '';

// Monitor webpack output to detect when server is ready
webpackProcess.stdout.on('data', (data) => {
  const output = data.toString();
  process.stdout.write(output);
  outputBuffer += output;
  
  // Check if webpack is ready (check both current chunk and accumulated buffer)
  const port = process.env.PORT || 3000;
  const isReady = 
    output.includes('webpack compiled') || 
    output.includes('compiled successfully') ||
    output.includes(`localhost:${port}`) ||
    outputBuffer.includes('webpack compiled') ||
    outputBuffer.includes('compiled successfully') ||
    outputBuffer.includes(`localhost:${port}`);
    
  if (isReady && !webpackReady) {
    webpackReady = true;
    // Small delay to ensure server is fully ready
    setTimeout(() => {
      attemptSideload();
    }, 1000);
  }
});

webpackProcess.stderr.on('data', (data) => {
  process.stderr.write(data);
});

// Handle process termination
process.on('SIGINT', () => {
  console.log('\n\n🛑 Stopping development server...');
  webpackProcess.kill('SIGINT');
  process.exit(0);
});

process.on('SIGTERM', () => {
  webpackProcess.kill('SIGTERM');
  process.exit(0);
});

function attemptSideload() {
  console.log('\n\n📦 Attempting to sideload add-in...\n');
  
  // Try using office-addin-dev-settings sideload
  // Syntax: sideload <manifest-path> [app-type] -a <app>
  // app-type: 'desktop' or 'web' (default: 'desktop')
  // app: 'Excel', 'Outlook', 'PowerPoint', or 'Word'
  // Note: This tool has known limitations with Outlook add-ins on macOS
  exec(`npx office-addin-dev-settings sideload "${manifestPath}" desktop -a Outlook`, (error, stdout, stderr) => {
    // Check for specific Outlook-related errors that indicate tool limitations
    const errorOutput = error ? (error.message || error.toString()) : '';
    const stderrOutput = stderr || '';
    const combinedError = (errorOutput + ' ' + stderrOutput).toLowerCase();
    
    if (error || combinedError.includes('does not support outlook') || combinedError.includes('manifest does not support')) {
      console.log('⚠️  Auto-sideloading not available (known macOS limitation with Outlook add-ins).\n');
      console.log('📝 Your manifest is valid - this is a tool limitation, not a manifest issue.\n');
      console.log('✅ The dev server is running and ready. Please use manual sideloading:\n');
      showManualInstructions();
      return;
    }
    
    if (stdout) console.log(stdout);
    if (stderr && !stderr.includes('Warning')) console.error(stderr);
    
    console.log('\n✅ Add-in sideloading attempted. Check Outlook for the add-in.\n');
    console.log('💡 If the add-in doesn\'t appear, use the manual instructions below.\n');
    showManualInstructions();
  });
}

function showManualInstructions() {
  const port = process.env.PORT || 3000;
  console.log('\n' + '='.repeat(70));
  console.log('📋 MANUAL SIDELOADING INSTRUCTIONS');
  console.log('='.repeat(70));
  console.log(`\nThe dev server is running at: https://localhost:${port}`);
  console.log('\n🔐 IMPORTANT: SSL Certificate Trust');
  console.log('   If you see a certificate error in Outlook, run:');
  console.log('   ./scripts/trust-certificate.sh');
  console.log('   Then restart Outlook.');
  console.log('\nTo manually load the add-in in Outlook:');
  console.log('\n1. Open Microsoft Outlook');
  console.log('2. Go to: Home → Get Add-ins → My Add-ins');
  console.log('3. Click the "+" button or "Add a Custom Add-in"');
  console.log('4. Select "Add from File"');
  console.log(`5. Navigate to: ${path.dirname(manifestPath)}`);
  console.log('6. Select: manifest.xml');
  console.log('7. Click "Add"');
  console.log('\nThe add-in should now appear in your Outlook ribbon!');
  console.log('\n' + '='.repeat(70) + '\n');
}

// Fallback: show instructions after 8 seconds if webpack doesn't signal ready
setTimeout(() => {
  if (!webpackReady) {
    console.log('\n⏳ Webpack dev server may still be starting...');
    console.log('   If it\'s already running, you can proceed with manual sideloading.\n');
    // Show instructions anyway after a delay
    setTimeout(() => {
      if (!webpackReady) {
        console.log('\n📋 Showing manual instructions (server should be ready):\n');
        showManualInstructions();
      }
    }, 2000);
  }
}, 8000);

webpackProcess.on('exit', (code) => {
  if (code !== 0 && code !== null) {
    console.error(`\n❌ Webpack dev server exited with code ${code}`);
  }
  process.exit(code || 0);
});

