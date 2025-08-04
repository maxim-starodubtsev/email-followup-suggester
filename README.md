# 📧 Followup Suggester - Outlook Add-in

An intelligent Outlook add-in that helps manage email follow-ups by analyzing sent emails and suggesting which ones need responses using AI-powered analysis.

## 🎯 Project Overview

The Followup Suggester is an Office Add-in designed for Microsoft Outlook that automatically identifies sent emails that may require follow-up. It leverages AI analysis through the DIAL API to provide intelligent summaries and actionable follow-up suggestions.

### Key Features

- **📊 Smart Email Analysis**: Automatically analyzes sent emails to identify those awaiting responses
- **🤖 AI-Powered Insights**: Uses DIAL API (GPT-4o-mini) for intelligent email summaries and follow-up suggestions
- **⏰ Priority Classification**: Categorizes emails as High/Medium/Low priority based on response time
- **💤 Snooze & Dismiss**: Full email management with customizable snooze options
- **🔄 Thread Analysis**: Analyzes entire conversation threads to understand context
- **👥 Multi-Account Support**: Filter analysis by different email accounts
- **📱 Modern UI**: Beautiful, responsive interface with modal dialogs

## 🏗️ Architecture

### Project Structure
```
Followup Suggester/
├── manifest.xml                    # Office Add-in manifest
├── package.json                    # Node.js dependencies
├── tsconfig.json                   # TypeScript configuration
├── webpack.config.js               # Build configuration
├── assets/                         # Icon assets
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-64.png
│   └── icon-80.png
├── doc/                           # Documentation
│   ├── project-plan.md
│   ├── development-log.md
│   ├── api-configuration.md
│   └── testing-guide.md
└── src/                           # Source code
    ├── commands/                  # Office commands
    │   ├── commands.html
    │   └── commands.ts
    ├── models/                    # TypeScript models
    │   ├── Configuration.ts
    │   └── FollowupEmail.ts
    ├── services/                  # Business logic
    │   ├── ConfigurationService.ts
    │   ├── EmailAnalysisService.ts
    │   └── LlmService.ts
    └── taskpane/                  # UI components
        ├── taskpane.html
        └── taskpane.ts
```

### Technology Stack

- **TypeScript**: Strongly-typed JavaScript for better development experience
- **Webpack**: Module bundler and development server
- **Office.js**: Microsoft Office JavaScript API
- **DIAL API**: AI integration for email analysis
- **HTML/CSS**: Modern responsive UI design

## 🚀 Getting Started

### Prerequisites

- **Node.js** (version 14 or higher)
- **npm** (comes with Node.js)
- **Microsoft Outlook** (Desktop or Web)
- **DIAL API Access** (for AI features)

### Installation

1. **Clone or download the project**
   ```bash
   cd "path/to/followup-suggester"
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Build the project**
   ```bash
   npm run build
   ```

### Development

1. **Start development server**
   ```bash
   npm start
   ```
   Or manually:
   ```bash
   npx webpack serve --mode development --port 3000
   ```

2. **Access the application**
   - Development server: `https://localhost:3000/`
   - Task pane: `https://localhost:3000/taskpane.html`

### Available Scripts

- `npm start` - Start development server with Office debugging
- `npm run build` - Build for production
- `npm run dev` - Development mode with webpack dev server

## 📋 Installation in Outlook

### Method 1: Sideloading (Recommended for Development)

1. **Open Microsoft Outlook**
2. **Navigate to Add-ins**:
   - **Desktop Outlook**: Home → Get Add-ins → Manage Add-ins
   - **Outlook Web**: Settings (⚙️) → View all Outlook settings → Mail → Integrate → Add-ins
3. **Upload Custom Add-in**:
   - Click "Upload My Add-in"
   - Select the `manifest.xml` file from your project directory
   - Click "Upload"
4. **Access the Add-in**:
   - The "Followup Suggester" button will appear in your Outlook ribbon
   - Click it to open the task pane

### Method 2: Office 365 Admin Center (For Organizations)

1. Deploy through your organization's Office 365 Admin Center
2. Add the manifest.xml to your organization's add-in catalog
3. Users can install from the organizational store

## 🔧 Configuration

### AI API Setup

The add-in comes pre-configured with DIAL API credentials:

- **Endpoint**: `https://ai-proxy.lab.epam.com`
- **API Key**: `dial-qjboupii21tb26eakd3ytcsb9po`
- **Model**: `gpt-4o-mini-2024-07-18`

### Customizing Settings

1. **Open the add-in** in Outlook
2. **Click Settings** button
3. **Configure**:
   - LLM API Endpoint
   - API Key
   - Enable/disable AI features
   - Display options

### Email Analysis Settings

- **Email Count**: Number of emails to analyze (10-100)
- **Days Back**: How far back to look (7-60 days)
- **Account Filter**: Filter by specific email accounts
- **AI Features**: Toggle AI summaries and suggestions

## 🧪 Testing

### Manual Testing

1. **Start the development server**:
   ```bash
   npm start
   ```

2. **Test UI directly**:
   - Visit `https://localhost:3000/taskpane.html`
   - Verify all controls and modals work

3. **Test in Outlook**:
   - Sideload the add-in using `manifest.xml`
   - Click "Analyze Emails" to test functionality
   - Verify email analysis and AI features

### Testing Scenarios

1. **Email Analysis**:
   - Send test emails to yourself
   - Wait a few minutes
   - Run analysis to see if they appear

2. **AI Integration**:
   - Enable AI features in settings
   - Verify summaries and suggestions appear
   - Test with different email types

3. **Snooze/Dismiss**:
   - Test snooze functionality with different time periods
   - Verify dismissed emails don't reappear
   - Test custom snooze dates

### Debugging

1. **Browser DevTools**:
   - Right-click in task pane → Inspect
   - Check Console for errors
   - Monitor Network tab for API calls

2. **Office.js Debugging**:
   - Use `console.log()` statements
   - Check Office.js API responses
   - Verify manifest configuration

## 🔍 Development Process

### Phase 1: Project Setup
- ✅ Created TypeScript-based Office Add-in project
- ✅ Configured Webpack build system
- ✅ Set up proper project structure

### Phase 2: Core Development
- ✅ Implemented email analysis service using Office.js
- ✅ Created configuration management system
- ✅ Built responsive UI with modern styling

### Phase 3: AI Integration
- ✅ Integrated DIAL API for intelligent analysis
- ✅ Implemented LLM service for summaries and suggestions
- ✅ Added error handling and fallback mechanisms

### Phase 4: Testing & Debugging
- ✅ Fixed 43 TypeScript compilation errors
- ✅ Resolved Office.js API compatibility issues
- ✅ Implemented proper error handling

### Phase 5: Documentation & Deployment
- ✅ Created comprehensive documentation
- ✅ Set up development server
- ✅ Prepared for production deployment

## 🎯 Achievements

### Technical Accomplishments
- **Zero Build Errors**: Successfully resolved all TypeScript compilation issues
- **Modern Architecture**: Clean separation of services, models, and UI components
- **AI Integration**: Seamless integration with DIAL API for intelligent features
- **Office.js Compatibility**: Full compatibility with modern Office.js APIs
- **Type Safety**: Comprehensive TypeScript implementation

### Feature Completeness
- **Email Analysis**: Full conversation thread analysis
- **Priority Classification**: Intelligent priority assignment
- **AI-Powered Insights**: Smart summaries and follow-up suggestions
- **User Management**: Snooze, dismiss, and organization features
- **Multi-Account Support**: Enterprise-ready account filtering

### User Experience
- **Modern UI**: Beautiful, responsive interface design
- **Intuitive Navigation**: Easy-to-use controls and modals
- **Real-time Feedback**: Status messages and loading states
- **Accessibility**: Proper HTML structure and ARIA labels

## 🚀 Future Enhancements

- **Calendar Integration**: Sync with calendar for meeting follow-ups
- **Template System**: Pre-built follow-up email templates
- **Analytics Dashboard**: Email response metrics and insights
- **Mobile Support**: Optimized mobile experience
- **Integration APIs**: Connect with CRM systems

## 📝 License

This project is developed by Max Starodubtsev.

## 🤝 Support

For technical support or questions about this add-in:
1. Check the documentation in the `doc/` folder
2. Review the development log for troubleshooting
3. Contact the development team

---

**Built with ❤️ using TypeScript, Office.js, and AI-powered analysis**