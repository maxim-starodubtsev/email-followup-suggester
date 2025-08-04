#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

// List of critical files that should be generated
const criticalFiles = [
    'dist/commands.js',
    'dist/taskpane.js',
    'dist/commands.html',
    'dist/taskpane.html',
    'dist/models/CacheEntry.d.ts',
    'dist/models/Configuration.d.ts',
    'dist/models/FollowupEmail.d.ts'
];

console.log('🔍 Verifying build output...');

let allFilesExist = true;
const missingFiles = [];

criticalFiles.forEach(file => {
    const fullPath = path.join(__dirname, file);
    if (fs.existsSync(fullPath)) {
        console.log(`✅ ${file}`);
    } else {
        console.log(`❌ ${file} - MISSING`);
        allFilesExist = false;
        missingFiles.push(file);
    }
});

console.log('\n📊 Build Verification Summary:');
if (allFilesExist) {
    console.log('🎉 All critical files are present. Build successful!');
    process.exit(0);
} else {
    console.log(`💥 Build verification failed. Missing ${missingFiles.length} file(s):`);
    missingFiles.forEach(file => console.log(`   - ${file}`));
    console.log('\n💡 Troubleshooting tips:');
    console.log('   1. Run: npm install');
    console.log('   2. Run: npm run build');
    console.log('   3. Check for TypeScript errors');
    process.exit(1);
}
