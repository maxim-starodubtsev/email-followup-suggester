# 🚀 Local Development & Debugging Guide for macOS

This guide will help you run and debug the Followup Suggester Outlook Add-in locally on macOS.

## 📋 Prerequisites

Before you begin, ensure you have:

1. **Node.js** (v18 or higher)
   ```bash
   node --version  # Should show v18.x.x or higher
   ```

2. **npm** (comes with Node.js)
   ```bash
   npm --version
   ```

3. **Microsoft Outlook** (Desktop app or Outlook Web)
   - Desktop: Outlook for Mac (Office 365 or Microsoft 365)
   - Web: Access to Outlook on the web (outlook.office.com)

4. **Code Editor** (VS Code recommended)
   - Install VS Code extensions:
     - TypeScript and JavaScript Language Features
     - ESLint
     - Prettier (optional)

## 🔧 Initial Setup

### 1. Install Dependencies

```bash
cd /Users/Maksym_Starodubtsev/Desktop/email-followup-suggester
npm install
```

### 2. Build the Project

```bash
npm run build
```

This will:
- Compile TypeScript to JavaScript
- Generate type declaration files (.d.ts)
- Bundle assets and HTML files
- Output everything to the `dist/` folder

## 🏃 Running the Development Server

### Option 1: Using npm start (Recommended)

```bash
npm start
```

This command:
- Starts the Office Add-in debugging tool
- Launches the webpack dev server on HTTPS (port 3000)
- Automatically opens Outlook with the add-in loaded

### Option 2: Using npm run dev (Manual)

```bash
npm run dev
```

This starts only the webpack dev server. You'll need to manually sideload the add-in in Outlook.

**Expected Output:**
```
webpack compiled successfully
webpack dev server running on https://localhost:3000
```

### 3. Access the Application

Once the dev server is running:

- **Task Pane**: `https://localhost:3000/taskpane.html`
- **Commands**: `https://localhost:3000/commands.html`
- **Assets**: `https://localhost:3000/assets/icon-*.png`

## 📱 Sideloading the Add-in in Outlook

### For Outlook Desktop (macOS)

1. **Open Microsoft Outlook**

2. **Navigate to Add-ins**:
   - Click **Home** tab
   - Click **Get Add-ins** button
   - Or go to **Tools** → **Get Add-ins**

3. **Upload Custom Add-in**:
   - Click **"My Add-ins"** tab
   - Click **"+"** or **"Add a Custom Add-in"**
   - Select **"Add from File"**
   - Navigate to your project folder
   - Select `manifest.xml`
   - Click **"Add"**

4. **Verify Installation**:
   - The add-in should appear in your ribbon
   - Look for "Followup Suggester" button
   - Click it to open the task pane

### For Outlook on the Web

1. **Open Outlook Web** (outlook.office.com)

2. **Navigate to Settings**:
   - Click the **Settings** gear icon (⚙️)
   - Go to **View all Outlook settings**
   - Click **Mail** → **Integrate** → **Add-ins**

3. **Upload Custom Add-in**:
   - Click **"My add-ins"** tab
   - Click **"+"** button
   - Select **"Add a custom add-in"** → **"Add from file"**
   - Upload `manifest.xml` from your project folder

4. **Access the Add-in**:
   - Open any email
   - Look for the add-in in the message ribbon
   - Click to open the task pane

## 🐛 Debugging

### 1. Browser DevTools (Outlook Web)

**For Outlook on the Web:**

1. **Open DevTools**:
   - Right-click anywhere in the task pane
   - Select **"Inspect"** or **"Inspect Element"**
   - Or press `Cmd + Option + I` (macOS)

2. **Debug Console**:
   - Check the **Console** tab for errors and logs
   - Use `console.log()`, `console.error()`, etc. in your code
   - Look for Office.js API errors

3. **Network Tab**:
   - Monitor API calls to DIAL API
   - Check for failed requests
   - Verify request/response payloads

4. **Sources Tab**:
   - Set breakpoints in your TypeScript files
   - Step through code execution
   - Inspect variables and call stack

### 2. Safari Web Inspector (Outlook Desktop macOS)

**For Outlook Desktop on macOS:**

1. **Enable Web Inspector**:
   ```bash
   defaults write com.microsoft.Outlook WebKitDeveloperExtrasEnabledPreferenceKey -bool true
   ```
   Restart Outlook after running this command.

2. **Open Web Inspector**:
   - Right-click in the task pane
   - Select **"Inspect Element"**
   - Safari Web Inspector will open

3. **Debug**:
   - Use Console, Network, and Sources tabs
   - Set breakpoints and step through code
   - Monitor network requests

### 3. VS Code Debugging

Create `.vscode/launch.json`:

```json
{
  "version": "0.2.0",
  "configurations": [
    {
      "type": "node",
      "request": "launch",
      "name": "Debug Tests",
      "runtimeExecutable": "npm",
      "runtimeArgs": ["run", "test"],
      "console": "integratedTerminal",
      "internalConsoleOptions": "neverOpen"
    },
    {
      "type": "chrome",
      "request": "launch",
      "name": "Debug Task Pane",
      "url": "https://localhost:3000/taskpane.html",
      "webRoot": "${workspaceFolder}",
      "sourceMaps": true
    }
  ]
}
```

### 4. Console Logging

Add debug statements in your code:

```typescript
// In any service file
console.log('[DEBUG] Current state:', state);
console.error('[ERROR] Failed to fetch:', error);

// Conditional logging
if (process.env.NODE_ENV === 'development') {
  console.log('[DEV] Debug info:', data);
}
```

### 5. TypeScript Source Maps

Source maps are enabled in development mode. This allows you to:
- Debug original TypeScript files (not compiled JavaScript)
- See original line numbers in stack traces
- Set breakpoints in `.ts` files

## 🔍 Common Debugging Scenarios

### Issue: Add-in Not Loading

**Symptoms**: Task pane shows blank or error message

**Debug Steps**:
1. Check browser console for errors
2. Verify dev server is running: `https://localhost:3000`
3. Check manifest.xml for correct URLs
4. Verify SSL certificate (may need to accept self-signed cert)
5. Check Office.js is loading: Look for `Office.onReady()` errors

**Solution**:
```bash
# Restart dev server
npm run dev

# Check manifest.xml SourceLocation URLs match dev server
# Should be: https://localhost:3000/taskpane.html
```

### Issue: Office.js API Errors

**Symptoms**: `Office.context.mailbox` is undefined or methods fail

**Debug Steps**:
1. Check Office.js version in manifest.xml
2. Verify `Office.onReady()` callback executed
3. Check API permissions in manifest.xml
4. Verify you're testing in correct context (MessageRead, etc.)

**Solution**:
```typescript
// Add this at the start of your code
Office.onReady((info) => {
  console.log('[DEBUG] Office ready:', info);
  console.log('[DEBUG] Mailbox available:', !!Office.context.mailbox);
});
```

### Issue: API Calls Failing

**Symptoms**: DIAL API requests return errors

**Debug Steps**:
1. Check Network tab for failed requests
2. Verify API endpoint URL in ConfigurationService
3. Check API key is set correctly
4. Look for CORS errors
5. Check request headers and payload

**Solution**:
```typescript
// Add logging in LlmService.ts
console.log('[DEBUG] API Request:', {
  url: endpoint,
  headers: headers,
  body: requestBody
});
```

### Issue: TypeScript Compilation Errors

**Symptoms**: Build fails with TypeScript errors

**Debug Steps**:
1. Run TypeScript compiler directly:
   ```bash
   npx tsc --noEmit
   ```
2. Check specific error messages
3. Verify all imports are correct
4. Check tsconfig.json settings

**Solution**:
```bash
# Clean and rebuild
npm run clean:build

# Check for type errors
npm run build:types
```

## 🧪 Testing

### Run Unit Tests

```bash
# Run all tests
npm test

# Watch mode (re-runs on file changes)
npm run test:watch

# With UI
npm run test:ui

# Coverage report
npm run test:coverage
```

### Manual Testing Checklist

1. **Email Analysis**:
   - [ ] Send test emails to yourself
   - [ ] Run "Analyze Emails"
   - [ ] Verify emails appear in results
   - [ ] Check priority classification

2. **AI Features**:
   - [ ] Enable AI in settings
   - [ ] Verify summaries appear
   - [ ] Check follow-up suggestions
   - [ ] Test with different email types

3. **UI Interactions**:
   - [ ] Test snooze functionality
   - [ ] Test dismiss functionality
   - [ ] Verify settings modal opens
   - [ ] Check all buttons work

## 🛠️ Development Workflow

### Typical Development Session

1. **Start Dev Server**:
   ```bash
   npm run dev
   ```

2. **Make Code Changes**:
   - Edit TypeScript files in `src/`
   - Webpack will auto-recompile
   - Browser will auto-reload (if using Outlook Web)

3. **Test Changes**:
   - Refresh Outlook add-in task pane
   - Check browser console for errors
   - Test functionality

4. **Debug Issues**:
   - Use browser DevTools
   - Add console.log statements
   - Check Network tab for API calls

5. **Run Tests**:
   ```bash
   npm test
   ```

### Hot Reload

The webpack dev server supports hot module replacement:
- Changes to TypeScript files trigger recompilation
- Changes to HTML/CSS trigger page reload
- For Outlook Desktop, you may need to manually refresh

### Stopping the Dev Server

Press `Ctrl + C` in the terminal where the dev server is running.

Or use:
```bash
npm run stop  # Stops Office debugging session
```

## 🔐 SSL Certificate Issues

The dev server uses HTTPS with a self-signed certificate. You may need to:

1. **Accept the certificate** in your browser:
   - Navigate to `https://localhost:3000`
   - Click "Advanced" → "Proceed to localhost"

2. **For Outlook Desktop**, you may need to trust the certificate:
   ```bash
   # Find the certificate in Keychain Access
   # Mark it as "Always Trust"
   ```

## 📝 Useful Commands Reference

```bash
# Development
npm run dev              # Start dev server
npm start                # Start with Office debugging
npm run stop             # Stop Office debugging

# Building
npm run build            # Production build
npm run clean            # Clean dist folder
npm run clean:build      # Clean and build

# Testing
npm test                 # Run tests
npm run test:watch       # Watch mode
npm run test:ui          # Test UI
npm run test:coverage    # Coverage report

# Validation
npm run validate         # Validate manifest.xml
```

## 🎯 Tips for macOS Development

1. **Use Terminal.app or iTerm2** for better terminal experience

2. **Enable File Watching** (if you hit file limit):
   ```bash
   # Increase file watcher limit
   echo fs.inotify.max_user_watches=524288 | sudo tee -a /etc/sysctl.conf
   sudo sysctl -p
   ```

3. **Use VS Code Integrated Terminal**:
   - Open terminal in VS Code: `` Ctrl + ` ``
   - Run commands directly from editor

4. **Monitor System Resources**:
   - Check Activity Monitor for Node processes
   - Restart dev server if it becomes unresponsive

5. **Clear Browser Cache**:
   - If changes don't appear, hard refresh: `Cmd + Shift + R`
   - Or clear browser cache completely

## 🆘 Troubleshooting

### Port 3000 Already in Use

```bash
# Find process using port 3000
lsof -ti:3000

# Kill the process
kill -9 $(lsof -ti:3000)

# Or use a different port (edit webpack.config.js)
```

### Module Not Found Errors

```bash
# Reinstall dependencies
rm -rf node_modules package-lock.json
npm install
```

### Office.js Not Loading

- Verify you're accessing via HTTPS (not HTTP)
- Check manifest.xml SourceLocation URLs
- Ensure Office.js CDN is accessible
- Check browser console for CORS errors

### Build Errors

```bash
# Clean everything and rebuild
npm run clean
rm -rf node_modules
npm install
npm run build
```

## 📚 Additional Resources

- [Office Add-ins Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/)
- [Office.js API Reference](https://docs.microsoft.com/en-us/javascript/api/office)
- [Webpack Documentation](https://webpack.js.org/)
- [TypeScript Handbook](https://www.typescriptlang.org/docs/)

---

**Happy Debugging! 🐛✨**

