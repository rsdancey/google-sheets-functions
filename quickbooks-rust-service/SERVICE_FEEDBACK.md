# Getting Feedback from the QuickBooks Sync Service

## The Problem
You mentioned the application runs but doesn't provide feedback. Here's how to see what it's doing:

## Command Line Options

### Show Help
```cmd
cargo run --release -- --help
```
Shows all available options.

### Show Current Configuration
```cmd
cargo run --release -- --show-config
```
Displays your current config (without showing sensitive API keys).

### Test Connections Only
```cmd
cargo run --release -- --test-connection
```
Tests QuickBooks and Google Sheets connections, then exits.

### Verbose Logging
```cmd
cargo run --release -- --verbose
```
Shows detailed debug information about what the service is doing.

### Skip Google Sheets Test
```cmd
cargo run --release -- --skip-sheets-test
```
Useful when you know Google Sheets isn't configured yet.

## What the Service Shows Now

### During Startup:
- ✅ Configuration loaded
- 📊 QuickBooks connection test results
- 📝 Google Sheets connection test results
- 🕐 Schedule information
- 🔄 Initial sync execution

### During Normal Operation:
- 🔄 Each sync start/completion
- 📊 Account data retrieved from QuickBooks
- 📝 Google Sheets update results  
- ✅ Success confirmations with timing
- 🕐 Periodic status updates (every 5 minutes)

### Error Reporting:
- ❌ Clear error messages
- 💡 Suggestions for fixes
- 🔧 Troubleshooting hints

## Example Output

```
[2025-07-06T01:30:00Z INFO qb_sync] Starting QuickBooks to Google Sheets sync service
[2025-07-06T01:30:00Z INFO qb_sync] Testing QuickBooks connection...
[2025-07-06T01:30:01Z INFO qb_sync] Mock mode: QuickBooks connection test skipped
[2025-07-06T01:30:01Z INFO qb_sync] Testing Google Sheets connection...
[2025-07-06T01:30:02Z INFO qb_sync] ✅ Initial sync completed successfully!
[2025-07-06T01:30:02Z INFO qb_sync] 🕐 Service now running on schedule: 0 0 * * * *
[2025-07-06T01:30:02Z INFO qb_sync] 📊 Will sync account 9445 to Google Sheets every time the schedule triggers
[2025-07-06T01:30:02Z INFO qb_sync] 🛑 Press Ctrl+C to stop the service
[2025-07-06T01:35:02Z INFO qb_sync] 🔄 Service running normally (5 minutes elapsed)
```

## Common Scenarios

### If Nothing Appears to Happen:
1. **Check if it's actually running**: Look for the process in Task Manager
2. **Use verbose mode**: `--verbose` flag shows much more detail
3. **Check configuration**: `--show-config` reveals what settings are loaded

### If You Get Errors:
1. **API Key issues**: "Invalid API key" means Google Apps Script setup needed
2. **QuickBooks issues**: "Failed to create session manager" means SDK installation needed
3. **Permission issues**: Run as Administrator

### If It Runs But Doesn't Sync:
1. **Check schedule**: The service waits for the next scheduled time
2. **Force immediate sync**: The service runs one sync immediately on startup
3. **Check Google Sheets**: Verify the spreadsheet and cell are being updated

## Quick Troubleshooting Commands

```cmd
# Show what configuration is loaded
cargo run --release -- --show-config

# Test everything with maximum feedback
cargo run --release -- --test-connection --verbose

# Run normally with detailed logging
cargo run --release -- --verbose

# Test just QuickBooks (if Google Sheets isn't set up yet)
cargo run --release -- --test-connection --skip-sheets-test --verbose
```

# Service Feedback and Troubleshooting

## ✅ Logger Initialization Issue Fixed

The service now properly handles logger initialization using `try_init()` method, which gracefully handles cases where the logger may already be initialized. This resolves the panic that could occur with:

```
Builder::init should not be called after logger initialized: SetLoggerError(())
```

## 🔧 Command-Line Options

The service provides comprehensive command-line options for testing and feedback:

### Show Help
```cmd
cargo run --release -- --help
```
Shows all available options.

### Show Current Configuration
```cmd
cargo run --release -- --show-config
```
Displays your current config (without showing sensitive API keys).

### Test Connections Only
```cmd
cargo run --release -- --test-connection
```
Tests QuickBooks and Google Sheets connections, then exits.

### Verbose Logging
```cmd
cargo run --release -- --verbose
```
Shows detailed debug information about what the service is doing.

### Skip Google Sheets Test
```cmd
cargo run --release -- --skip-sheets-test
```
Useful when you know Google Sheets isn't configured yet.

## What the Service Shows Now

### During Startup:
- ✅ Configuration loaded
- 📊 QuickBooks connection test results
- 📝 Google Sheets connection test results
- 🕐 Schedule information
- 🔄 Initial sync execution

### During Normal Operation:
- 🔄 Each sync start/completion
- 📊 Account data retrieved from QuickBooks
- 📝 Google Sheets update results  
- ✅ Success confirmations with timing
- 🕐 Periodic status updates (every 5 minutes)

### Error Reporting:
- ❌ Clear error messages
- 💡 Suggestions for fixes
- 🔧 Troubleshooting hints

## Example Output

```
[2025-07-06T01:30:00Z INFO qb_sync] Starting QuickBooks to Google Sheets sync service
[2025-07-06T01:30:00Z INFO qb_sync] Testing QuickBooks connection...
[2025-07-06T01:30:01Z INFO qb_sync] Mock mode: QuickBooks connection test skipped
[2025-07-06T01:30:01Z INFO qb_sync] Testing Google Sheets connection...
[2025-07-06T01:30:02Z INFO qb_sync] ✅ Initial sync completed successfully!
[2025-07-06T01:30:02Z INFO qb_sync] 🕐 Service now running on schedule: 0 0 * * * *
[2025-07-06T01:30:02Z INFO qb_sync] 📊 Will sync account 9445 to Google Sheets every time the schedule triggers
[2025-07-06T01:30:02Z INFO qb_sync] 🛑 Press Ctrl+C to stop the service
[2025-07-06T01:35:02Z INFO qb_sync] 🔄 Service running normally (5 minutes elapsed)
```

## Common Scenarios

### If Nothing Appears to Happen:
1. **Check if it's actually running**: Look for the process in Task Manager
2. **Use verbose mode**: `--verbose` flag shows much more detail
3. **Check configuration**: `--show-config` reveals what settings are loaded

### If You Get Errors:
1. **API Key issues**: "Invalid API key" means Google Apps Script setup needed
2. **QuickBooks issues**: "Failed to create session manager" means SDK installation needed
3. **Permission issues**: Run as Administrator

### If It Runs But Doesn't Sync:
1. **Check schedule**: The service waits for the next scheduled time
2. **Force immediate sync**: The service runs one sync immediately on startup
3. **Check Google Sheets**: Verify the spreadsheet and cell are being updated

## Quick Troubleshooting Commands

```cmd
# Show what configuration is loaded
cargo run --release -- --show-config

# Test everything with maximum feedback
cargo run --release -- --test-connection --verbose

# Run normally with detailed logging
cargo run --release -- --verbose

# Test just QuickBooks (if Google Sheets isn't set up yet)
cargo run --release -- --test-connection --skip-sheets-test --verbose
```

The service now provides much more feedback about what it's doing at each step!
