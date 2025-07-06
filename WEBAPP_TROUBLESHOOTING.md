# Google Apps Script Web App Troubleshooting

## Problem: "Page Not Found" Error

The error you're seeing indicates that the Web App URL is not working. Here's how to fix it:

## Solution Steps

### 1. Re-authenticate with Google Apps Script

```bash
cd /path/to/your/google-sheets-functions
npx clasp login
```

### 2. Push Your Code to Google Apps Script

```bash
npm run build
npx clasp push
```

### 3. Deploy as Web App

1. Open [script.google.com](https://script.google.com)
2. Find your project: "Google Sheets Custom Functions"
3. Click "Deploy" → "New deployment"
4. Configuration:
   - **Type**: Web app
   - **Description**: "QuickBooks Sync API"
   - **Execute as**: Me (your email)
   - **Who has access**: Anyone
5. Click "Deploy"
6. **IMPORTANT**: Copy the Web App URL that ends with `/exec`

### 4. Set Up API Key

1. In Google Apps Script editor, run this function:
   ```javascript
   function setupQuickBooksIntegration() {
     const apiKey = Utilities.getUuid();
     PropertiesService.getScriptProperties().setProperty('QB_API_KEY', apiKey);
     console.log('API Key:', apiKey);
     return apiKey;
   }
   ```

2. Copy the API key from the execution log

### 5. Update Your Rust Configuration

Edit `config/config.toml`:

```toml
[google_sheets]
# Use the EXACT URL from step 3
webapp_url = "https://script.google.com/macros/s/AKfycbx.../exec"
# Use the API key from step 4
api_key = "12345678-1234-1234-1234-123456789abc"
# Your spreadsheet ID (from the URL)
spreadsheet_id = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"
# Your sheet tab name
sheet_name = "The Simple Buckets"
# Cell to update
cell_address = "A1"
```

### 6. Test the Connection

```bash
cargo run --release -- --test-connection
```

## Common Issues

### Issue: "Authorization required" error
- **Solution**: Make sure "Who has access" is set to "Anyone" in the Web App deployment

### Issue: "Invalid API key" error
- **Solution**: Run `setupQuickBooksIntegration()` in Google Apps Script and copy the new API key

### Issue: "Sheet not found" error
- **Solution**: Check that `sheet_name` matches the exact tab name in your spreadsheet

### Issue: "Spreadsheet not found" error
- **Solution**: Verify the `spreadsheet_id` is correct and the Google Apps Script has access to it

## Web App URL Format

The correct URL format is:
```
https://script.google.com/macros/s/[SCRIPT_ID]/exec
```

NOT:
```
https://script.google.com/d/[SCRIPT_ID]/edit  # This is the editor URL
```

## Testing Your Web App

You can test the Web App directly with curl:

```bash
curl -X POST "https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec" \
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
