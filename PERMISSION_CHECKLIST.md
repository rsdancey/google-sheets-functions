# 🚀 Quick Permission Setup Checklist

## **Pre-Setup**
- [ ] You have edit access to your Google Sheets document
- [ ] You have access to the Google Apps Script project
- [ ] You know your spreadsheet ID and sheet name

## **Google Apps Script Setup**
- [ ] **Run `setupPermissions()`** - Authorize basic permissions
- [ ] **Run `checkPermissions()`** - Verify access levels
- [ ] **Run `setupQuickBooksIntegration()`** - Generate API key
- [ ] **Copy the API key** - Save for Rust configuration

## **Web App Deployment**
- [ ] **Click "Deploy" → "New Deployment"**
- [ ] **Choose "Web app"**
- [ ] **Set "Execute as": "Me"**
- [ ] **Set "Who has access": "Anyone"**
- [ ] **Click "Deploy"**
- [ ] **Copy the Web App URL** - Save for Rust configuration

## **Testing**
- [ ] **Run `testWebAppEndpoint()`** - Test internal functionality
- [ ] **Check your spreadsheet** - Should see test data
- [ ] **Test Rust service** - Use `--test-connection` flag

## **Configuration Update**
- [ ] **Update `config/config.toml`** with real values:
  - `webapp_url` = Your Web App URL
  - `api_key` = Your generated API key
  - `spreadsheet_id` = Your spreadsheet ID
  - `sheet_name` = Your sheet tab name
  - `cell_address` = Where to put data

## **Final Verification**
- [ ] **Rust service connects** without errors
- [ ] **Google Apps Script receives data** (check logs)
- [ ] **Spreadsheet updates** with QuickBooks data
- [ ] **Timestamps show** when data was last updated

## **If Something Goes Wrong**
1. **Check Google Apps Script logs** for error messages
2. **Verify spreadsheet ID** matches your actual spreadsheet
3. **Confirm sheet name** exists in your spreadsheet
4. **Test API key** is copied correctly
5. **Ensure Web App URL** is the deployed version

## **Key Values You Need**
From your setup, you'll need:
- **API Key**: `QB_API_KEY` from script properties
- **Web App URL**: From deployment dialog
- **Spreadsheet ID**: From your Google Sheets URL
- **Sheet Name**: The tab name in your spreadsheet

✅ **Ready to sync QuickBooks data to Google Sheets!**
