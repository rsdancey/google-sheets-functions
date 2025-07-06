# QuickBooks Integration Troubleshooting Guide

## Current Issue
BeginSession fails with error: `Exception occurred. (0x80020009)`

## Root Cause Analysis
The error `0x80020009` (DISP_E_EXCEPTION) indicates that QuickBooks Desktop is throwing an exception during the BeginSession call. This is typically caused by:

1. **Application Authorization Required** - QuickBooks needs to authorize the application
2. **Integrated Applications Settings** - QuickBooks Desktop settings may block external applications
3. **Company File Access** - The company file may be locked or inaccessible

## Steps to Fix

### 1. Check QuickBooks Desktop Integration Settings

1. **Open QuickBooks Desktop**
2. **Go to Edit > Preferences**
3. **Select "Integrated Applications" from the left panel**
4. **Click "Company Preferences" tab**
5. **Ensure the following settings:**
   - ✅ "Don't allow any applications to access this company file" should be **UNCHECKED**
   - ✅ "Allow applications to access this company file" should be **CHECKED**
   - ✅ "Prompt the user to allow each application to access this company file" should be **CHECKED**

### 2. Clear Application Authorization List

If the application was previously denied access:

1. **In the same "Integrated Applications" preferences**
2. **Click "Remove" next to any "QuickBooks-Sheets-Sync" or similar entries**
3. **This will allow the application to request permission again**

### 3. Test Application Permission Request

1. **Close QuickBooks Desktop completely**
2. **Reopen QuickBooks Desktop with your company file**
3. **Run the sync service again**
4. **Look for an authorization dialog from QuickBooks**

### 4. Check Company File Status

1. **Ensure the company file is open in QuickBooks Desktop**
2. **Ensure no other applications are accessing the file**
3. **Try closing and reopening the company file**

### 5. Try Single-User Mode

1. **In QuickBooks Desktop, go to File > Switch to Single-user Mode**
2. **Run the sync service again**

## Technical Details

- **Application ID**: "QuickBooks-Sheets-Sync"
- **Application Name**: "QuickBooks to Google Sheets Sync Service"
- **Connection Type**: Local QuickBooks Desktop (localQBD)
- **File Access Modes Tested**: qbFileOpenDoNotCare (0), qbFileOpenSingleUser (1), qbFileOpenMultiUser (2)

## Current Status

✅ **OpenConnection**: Working successfully
❌ **BeginSession**: Failing with exception
✅ **COM Variant Creation**: Fixed and working
✅ **Retry Logic**: Implemented with multiple modes

## Next Steps

1. **Check QuickBooks Desktop settings** (most likely fix)
2. **Clear any existing application permissions**
3. **Test with authorization dialog**
4. **If still failing, try specific file path instead of AUTO**

## Success Indicators

When working correctly, you should see:
- QuickBooks authorization dialog on first run
- "BeginSession successful" in the logs
- Real account data retrieved (not mock data)
- Google Sheets updated with actual values
