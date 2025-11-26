# Docker Setup for Email Followup Suggester

## 🎯 Purpose

These files provide a Docker-based testing solution to bypass macOS OneDrive permissions issues.

## 📦 Contents

- `Dockerfile` - Docker image definition
- `docker-compose.yml` - Docker services configuration
- `.dockerignore` - Files to exclude from Docker image
- `docker-test.sh` - Convenient test runner script

## 🚀 How to Use

### Step 1: Copy Files

Copy all files from this folder to your project root:

```bash
cp -r ~/Desktop/email-followup-docker-setup/* ~/Library/CloudStorage/OneDrive-EPAM/Repos/email-followup-suggester/
```

### Step 2: Make Script Executable

```bash
cd ~/Library/CloudStorage/OneDrive-EPAM/Repos/email-followup-suggester
chmod +x docker-test.sh
```

### Step 3: Run Tests

```bash
# Run all tests
./docker-test.sh test

# Run tests in watch mode
./docker-test.sh watch

# Generate coverage report
./docker-test.sh coverage
```

## 📋 Prerequisites

1. **Docker Desktop** must be installed and running
   - Download: https://www.docker.com/products/docker-desktop
   - Verify: `docker --version`

## 🔧 Commands

```bash
./docker-test.sh build      # Build Docker image
./docker-test.sh test       # Run tests once
./docker-test.sh watch      # Run tests in watch mode
./docker-test.sh coverage   # Generate coverage report
./docker-test.sh shell      # Open shell in container
./docker-test.sh clean      # Clean Docker resources
./docker-test.sh rebuild    # Rebuild from scratch
```

## ✅ Expected Result

When tests run successfully, you should see:

```
=== Followup Suggester Docker Test Runner ===

Running tests...
✓ Tests complete

Test Suites: 8 passed, 8 total
Tests:       150 passed, 150 total
```

## 🐛 Troubleshooting

### Docker Not Found
Install Docker Desktop from: https://www.docker.com/products/docker-desktop

### Permission Denied
```bash
chmod +x docker-test.sh
```

### Build Fails
```bash
./docker-test.sh rebuild
```

## 💡 Why Docker?

Docker solves the macOS OneDrive permissions issue by:
- Running tests in an isolated Linux container
- Bypassing macOS file system restrictions
- Providing a consistent environment across machines

---

**Note**: If Docker setup is too complex, consider **moving the repository** out of OneDrive instead:

```bash
mv ~/Library/CloudStorage/OneDrive-EPAM/Repos/email-followup-suggester ~/Desktop/email-followup-suggester
```

