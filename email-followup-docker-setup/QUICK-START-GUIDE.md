# 🚀 Quick Start Guide - Running Tests with Docker

## ⚡ Two Solutions Available

### **RECOMMENDED: Option 1 - Move Repository (2 minutes)**

**Fastest and simplest solution. No Docker needed!**

```bash
# 1. Move repository out of OneDrive to your Desktop
mv ~/Library/CloudStorage/OneDrive-EPAM/Repos/email-followup-suggester ~/Desktop/email-followup-suggester

# 2. Navigate to new location
cd ~/Desktop/email-followup-suggester

# 3. Install dependencies (if needed)
npm install

# 4. Run tests - THEY WILL WORK!
npm test

# 5. Generate coverage
npm run test:coverage
```

**Why this works:**
- ✅ No OneDrive sync conflicts
- ✅ No macOS permissions issues
- ✅ Full IDE integration
- ✅ Instant results
- ⚡ **This is the easiest solution**

---

### **Option 2 - Docker Setup (if you need OneDrive sync)**

If you must keep the repo in OneDrive, use Docker:

#### **Prerequisites**

1. **Install Docker Desktop**
   - Download: https://www.docker.com/products/docker-desktop
   - Install and start Docker Desktop
   - Verify installation:
     ```bash
     docker --version
     docker-compose --version
     ```

#### **Setup Steps**

```bash
# 1. Copy Docker files to your project
cp ~/Desktop/email-followup-docker-setup/* ~/Library/CloudStorage/OneDrive-EPAM/Repos/email-followup-suggester/

# 2. Navigate to your project
cd ~/Library/CloudStorage/OneDrive-EPAM/Repos/email-followup-suggester

# 3. Make script executable
chmod +x docker-test.sh

# 4. Build Docker image (first time only, takes 3-5 minutes)
./docker-test.sh build

# 5. Run tests
./docker-test.sh test

# 6. Or run in watch mode
./docker-test.sh watch

# 7. Or generate coverage
./docker-test.sh coverage
```

#### **Docker Commands**

```bash
./docker-test.sh build      # Build Docker image
./docker-test.sh test       # Run all tests once
./docker-test.sh watch      # Run tests in watch mode (auto-rerun)
./docker-test.sh coverage   # Generate coverage report
./docker-test.sh shell      # Open shell in container for debugging
./docker-test.sh clean      # Clean up Docker resources
./docker-test.sh rebuild    # Rebuild everything from scratch
```

#### **What to Expect**

When tests run successfully:

```
=== Followup Suggester Docker Test Runner ===

Running tests...

 ✓ tests/services/EmailAnalysisService.test.ts (35)
 ✓ tests/services/LlmService.test.ts (28)
 ✓ tests/services/ConfigurationService.test.ts (12)
 ✓ tests/services/BatchProcessor.test.ts (18)
 ✓ tests/services/CacheService.test.ts (24)
 ✓ tests/services/XmlParsingService.test.ts (8)
 ✓ tests/services/DialApiUrl.test.ts (6)
 ✓ tests/services/LlmAndUiIntegration.test.ts (4)

Test Files  8 passed (8)
     Tests  135 passed (135)

✓ Tests complete
```

---

## 🎯 Which Option Should I Choose?

### Choose **Option 1 (Move Repository)** if:
- ✅ You want the quickest solution
- ✅ You don't need OneDrive sync for development
- ✅ You want full IDE support without Docker complexity
- ✅ **RECOMMENDED FOR MOST USERS**

### Choose **Option 2 (Docker)** if:
- ✅ You must keep the repo in OneDrive for backup/sync
- ✅ You're comfortable with Docker
- ✅ You want isolated testing environment
- ✅ You need consistent CI/CD setup

---

## 🐛 Troubleshooting

### "Docker not found"
Install Docker Desktop: https://www.docker.com/products/docker-desktop

### "Permission denied" on docker-test.sh
```bash
chmod +x docker-test.sh
```

### Docker build fails
```bash
./docker-test.sh clean
./docker-test.sh rebuild
```

### Tests still fail with EPERM
This means you're running tests directly (npm test) instead of through Docker.
Use `./docker-test.sh test` instead.

---

## 📊 Expected Test Results

After successful setup, you should see:
- ✅ All 8 test suites passing
- ✅ 135+ tests passing
- ✅ Coverage report generated in `./coverage/`
- ✅ No EPERM errors

---

## 🆘 Need Help?

If you're stuck, try **Option 1 (Move Repository)** first - it's the simplest and fastest solution!

```bash
mv ~/Library/CloudStorage/OneDrive-EPAM/Repos/email-followup-suggester ~/Desktop/email-followup-suggester
cd ~/Desktop/email-followup-suggester
npm install
npm test
```

This will get your tests running in under 2 minutes! 🎉

