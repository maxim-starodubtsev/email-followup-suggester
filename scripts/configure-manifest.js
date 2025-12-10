const fs = require('fs');
const path = require('path');

const manifestPath = path.join(__dirname, '../manifest.xml');
const prodUrl = process.argv[2];

if (!prodUrl) {
  console.error('Usage: node configure-manifest.js <production-url>');
  process.exit(1);
}

// Remove trailing slash if present
const cleanUrl = prodUrl.replace(/\/$/, '');

try {
  let content = fs.readFileSync(manifestPath, 'utf8');
  
  // Replace localhost:3000 with production URL
  // We use a regex to capture the protocol and host to be safe, 
  // but targeting https://localhost:3000 specifically is safer for this project's default.
  const regex = /https:\/\/localhost:3000/g;
  
  if (!regex.test(content)) {
    console.warn('Warning: "https://localhost:3000" not found in manifest.xml');
  }
  
  const newContent = content.replace(regex, cleanUrl);
  
  fs.writeFileSync(manifestPath, newContent);
  console.log(`✅ Successfully updated manifest.xml to use URL: ${cleanUrl}`);
  
} catch (error) {
  console.error('Error updating manifest:', error);
  process.exit(1);
}
