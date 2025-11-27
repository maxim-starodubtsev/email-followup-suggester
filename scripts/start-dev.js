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

// Monitor webpack output to detect when server is ready
webpackProcess.stdout.on('data', (data) => {
  const output = data.toString();
  process.stdout.write(output);
  
  // Check if webpack is ready
  if (output.includes('webpack compiled') || output.includes('localhost:3000')) {
    if (!webpackReady) {
      webpackReady = true;
      attemptSideload();
    }
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
  exec(`npx office-addin-dev-settings sideload "${manifestPath}" outlook`, (error, stdout, stderr) => {
    if (error) {
      console.log('⚠️  Auto-sideloading failed. Use manual instructions below.\n');
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
  console.log('\n' + '='.repeat(70));
  console.log('📋 MANUAL SIDELOADING INSTRUCTIONS');
  console.log('='.repeat(70));
  console.log('\nThe dev server is running at: https://localhost:3000');
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

// Fallback: show instructions after 5 seconds if webpack doesn't signal ready
setTimeout(() => {
  if (!webpackReady) {
    console.log('\n⏳ Waiting for webpack dev server to start...');
    console.log('   (This may take a few more seconds)\n');
  }
}, 5000);

webpackProcess.on('exit', (code) => {
  if (code !== 0 && code !== null) {
    console.error(`\n❌ Webpack dev server exited with code ${code}`);
  }
  process.exit(code || 0);
});

