# QuickBooks Desktop Integration - User Testing Guide

## Quick Start

### Prerequisites
1. **QuickBooks Desktop** must be installed (Pro, Premier, or Enterprise)
2. **QuickBooks Desktop must be RUNNING**
3. **A company file must be OPEN** in QuickBooks
4. QuickBooks should be in a **ready state** (not processing anything)
5. You may need to run as **Administrator**

### Testing Steps

#### Step 1: Prepare QuickBooks
1. Open QuickBooks Desktop
2. Open a company file (any company file will work)
3. Wait for QuickBooks to finish loading completely
4. Make sure QuickBooks is not showing any modal dialogs
5. Leave QuickBooks running in the background

#### Step 2: Run the Test
```powershell
# From the quickbooks-rust-service directory
cargo run
```

Or run the executable directly:
```powershell
.\target\debug\qb_sync.exe
```

#### Step 3: Expected Behavior

**If Successful:**
- ✅ Session manager creation successful
- ✅ OpenConnection successful
- ✅ BeginSession successful
- ✅ Version information displayed
- ✅ Account data retrieval test
- ✅ XML response received

**If Failed:**
- ❌ Session manager functionality test failed
- ❌ COM exceptions (0x80020009, 0x80020006)
- ❌ "Object is not functional" messages

## Troubleshooting

### Common Issues

#### 1. "Session manager functionality test failed"
**Cause**: QuickBooks Desktop is not running or no company file is open.
**Solution**: 
- Make sure QuickBooks Desktop is running
- Open a company file
- Wait for QuickBooks to finish loading completely

#### 2. "COM exception 0x80020009"
**Cause**: QuickBooks Desktop is not ready or in a bad state.
**Solution**:
- Close and restart QuickBooks Desktop
- Open a company file
- Wait for it to finish loading
- Try running the test again

#### 3. "Method not found - wrong object type"
**Cause**: QBFC version mismatch or corrupted installation.
**Solution**:
- Check if QuickBooks Desktop is properly installed
- Try running QuickBooks Desktop as Administrator
- Reinstall QuickBooks Desktop if necessary

#### 4. "All OpenConnection methods failed"
**Cause**: Application needs to be authorized by QuickBooks.
**Solution**:
- When running the test, QuickBooks may show an authorization dialog
- Click "Yes" to authorize the application
- The application ID is "QuickBooks-Sheets-Sync-v1"

### Advanced Troubleshooting

#### Check Windows Registry
The service tests multiple QBFC versions:
- QBFC16.QBSessionManager (QB v24+)
- QBFC15.QBSessionManager (QB v23)
- QBFC14.QBSessionManager (QB v22)
- QBFC13.QBSessionManager (QB v21)

Check Windows Registry at:
```
HKEY_CLASSES_ROOT\QBFC16.QBSessionManager
HKEY_CLASSES_ROOT\QBFC15.QBSessionManager
etc.
```

#### Check QuickBooks Version
In QuickBooks Desktop:
1. Go to Help > About QuickBooks
2. Note the version number
3. Make sure it's a supported version (2021+)

#### Run as Administrator
Sometimes COM registration requires administrator privileges:
```powershell
# Run PowerShell as Administrator
cargo run
```

## What the Service Does

### Phase 1: SDK Availability Test
- Creates QuickBooks session manager objects
- Tests multiple QBFC versions
- Verifies basic functionality

### Phase 2: Registration Test
- Opens connection to QuickBooks
- Begins a session
- Tests version information
- Ends session and closes connection

### Phase 3: XML Processing Test
- Creates an account query XML request
- Processes the request through QuickBooks
- Returns XML response with account data

## Expected Output

### Successful Run
```
🔧 QuickBooks Desktop Integration Service
==========================================
Starting QuickBooks integration test...
📋 Step 1: Testing QuickBooks SDK availability...
✅ Successfully created session manager: QBFC16.QBSessionManager
✅ Session manager functionality test passed!
✅ QuickBooks SDK is available

📋 Step 2: Attempting to register with QuickBooks...
✅ OpenConnection2 successful (basic method)
✅ BeginSession successful
✅ Current company file: C:\Users\...\sample_company.qbw
✅ Version information displayed
✅ EndSession successful
✅ CloseConnection successful

📋 Step 3: Testing XML processing...
✅ Account data retrieval successful!
📄 XML Response received (2847 characters)
📋 Response preview: <?xml version="1.0" ?>...

🎉 SUCCESS! Application has been registered with QuickBooks.
```

### Failed Run
```
❌ Session manager functionality test failed: Session manager failed all tests
❌ QuickBooks SDK not available: Failed to create QuickBooks session manager
Error: QuickBooks SDK test failed
```

## Next Steps After Success

Once the service works successfully:
1. The application is registered with QuickBooks
2. You can proceed to implement full account data synchronization
3. You can test with different XML requests
4. You can integrate with Google Sheets API

## C++ SDK Comparison

This Rust implementation follows the same patterns as the official Intuit C++ SDK sample (`sdktest.cpp`):
- Same session management flow
- Same error handling approach
- Same XML processing method
- Same version testing

## Support

If you continue to have issues:
1. Check the QuickBooks Desktop logs
2. Verify your QuickBooks version is supported
3. Try with a different company file
4. Contact Intuit support for QuickBooks Desktop issues
