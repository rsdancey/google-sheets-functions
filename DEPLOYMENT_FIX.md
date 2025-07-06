# Step-by-Step Web App Deployment Fix

## The Problem
You're getting "Page Not Found" because the Web App URL in your `config.toml` is either:
1. **Not deployed as a Web App** (still just a script)
2. **Wrong URL format** (editor URL instead of execution URL)
3. **Old deployment** (needs to be re-deployed)

## Step-by-Step Solution

### Step 1: Check Your Current Config
What's in your `config/config.toml` file for the webapp_url?

It should look like:
```toml
[google_sheets]
webapp_url = "https://script.google.com/macros/s/SOME_LONG_ID/exec"
```

**NOT** like:
```toml
webapp_url = "https://script.google.com/d/SOME_ID/edit"
```

### Step 2: Deploy the Web App Correctly

1. **Go to [script.google.com](https://script.google.com)**
2. **Find your project** (should be called something like "Google Sheets Custom Functions")
3. **Click "Deploy" button** (top right)
4. **Click "New deployment"**
5. **Settings**:
   - Type: **Web app**
   - Description: "QuickBooks Sync"
   - Execute as: **Me**
   - Who has access: **Anyone**
6. **Click "Deploy"**
7. **Copy the Web App URL** (it will end with `/exec`)

### Step 3: Test the Web App URL

Before updating your config, test the URL directly:

```bash
curl -X POST "YOUR_WEB_APP_URL_HERE" \
  -H "Content-Type: application/json" \
  -d '{"test": "data"}'
```

**Expected response**: Should NOT be HTML with "Page Not Found"

### Step 4: Generate API Key

1. **In Google Apps Script editor**, run this function:
   ```javascript
   function setupQuickBooksIntegration() {
     const apiKey = Utilities.getUuid();
     PropertiesService.getScriptProperties().setProperty('QB_API_KEY', apiKey);
     console.log('API Key:', apiKey);
     return apiKey;
   }
   ```

2. **Copy the API key** from the execution log

### Step 5: Update Your Config

Edit `config/config.toml`:

```toml
[google_sheets]
webapp_url = "PASTE_THE_EXACT_URL_FROM_STEP_2"
api_key = "PASTE_THE_API_KEY_FROM_STEP_4"
spreadsheet_id = "YOUR_SPREADSHEET_ID"
sheet_name = "The Simple Buckets"
cell_address = "A1"
```

### Step 6: Test Again

```bash
cargo run --release -- --test-connection
```

## Common URL Format Mistakes

❌ **Wrong** (editor URL):
```
https://script.google.com/d/1Xzoi9iquQrT_nJJxb_7y100v5LPzI4w4u2ICLpPoBO7/edit
```

✅ **Correct** (execution URL):
```
https://script.google.com/macros/s/AKfycbxXXXXXXXXXXXXXXXXXXXXXXXXXXX/exec
```

## If You're Still Getting "Page Not Found"

1. **Check the deployment is active**:
   - Go to script.google.com
   - Click "Deploy" → "Manage deployments"
   - Make sure there's an active Web App deployment

2. **Try creating a new deployment**:
   - Sometimes old deployments get corrupted
   - Create a fresh "New deployment"

3. **Check permissions**:
   - Make sure "Who has access" is set to "Anyone"
   - If it's "Anyone with Google account", external services can't access it

## Next Steps

Please check your current `config.toml` file and let me know:
1. What's your current `webapp_url`?
2. Have you deployed it as a Web App (not just saved the script)?
3. What happens when you test the URL with curl?

This will help me identify exactly what's wrong with your deployment.
