# Fix Google Sheets API Key Error

## Problem: "Invalid API key"

The service is working correctly, but the Google Apps Script doesn't recognize your API key.

## Step-by-Step Solution

### Step 1: Set Up Google Apps Script API Key

1. **Open Google Apps Script**:
   - Go to [script.google.com](https://script.google.com)
   - Find your project or create a new one

2. **Run the setup function**:
   - In the script editor, click the function dropdown
   - Select `setupQuickBooksIntegration`
   - Click the "Run" button (▶️)
   - **Copy the API key** from the execution log

### Step 2: Deploy as Web App

1. **Click "Deploy" → "New deployment"**
2. **Settings**:
   - Type: **Web app**
   - Description: "QuickBooks Sync"
   - Execute as: **Me**
   - Who has access: **Anyone**
3. **Click "Deploy"**
4. **Copy the Web App URL** (ends with `/exec`)

### Step 3: Update Your Configuration

Edit `config/config.toml`:

```toml
[google_sheets]
# Use the EXACT URL from Step 2
webapp_url = "https://script.google.com/macros/s/AKfycbx.../exec"
# Use the API key from Step 1
api_key = "12345678-1234-1234-1234-123456789abc"
# Your spreadsheet ID
spreadsheet_id = "YOUR_SPREADSHEET_ID"
# Your sheet tab name
sheet_name = "The Simple Buckets"
# Cell to update
cell_address = "A1"
```

### Step 4: Test Again

```cmd
cargo run --release -- --test-connection
```

## Quick Test Script

Here's the exact Google Apps Script function you need:

```javascript
function setupQuickBooksIntegration() {
  const apiKey = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty('QB_API_KEY', apiKey);
  console.log('API Key generated:', apiKey);
  console.log('Copy this API key to your Rust service config.toml file');
  return apiKey;
}
```

## Verification

To verify your setup works, test with curl:

```bash
curl -X POST "YOUR_WEBAPP_URL" \
  -H "Content-Type: application/json" \
  -d '{
    "accountNumber": "TEST",
    "accountValue": 123.45,
    "spreadsheetId": "YOUR_SPREADSHEET_ID",
    "cellAddress": "A1",
    "sheetName": "The Simple Buckets",
    "apiKey": "YOUR_API_KEY"
  }'
```

Expected response:
```json
{"success": true, "message": "Account TEST updated: 123.45 at ..."}
```

## If You Still Get "Invalid API key"

1. **Check the API key matches exactly** (no extra spaces)
2. **Re-run setupQuickBooksIntegration()** to generate a new key
3. **Make sure you deployed the latest code** with the setup function
4. **Check the Web App permissions** are set to "Anyone"

The good news is your QuickBooks service is working perfectly with mock data! Now you just need to connect it to Google Sheets properly.
