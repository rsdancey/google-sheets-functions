# Google Sheets Custom Functions

This project provides custom functions for Google Sheets using Google Apps Script.

## Setup

1. Install dependencies:
   ```bash
   npm install
   ```

2. Install clasp globally (if not already installed):
   ```bash
   npm install -g @google/clasp
   ```

3. Login to Google Apps Script:
   ```bash
   clasp login
   ```

4. Create a new Google Apps Script project:
   ```bash
   clasp create --type sheets
   ```

5. Build the TypeScript files:
   ```bash
   npm run build
   ```

6. Push to Google Apps Script:
   ```bash
   npm run push
   ```

## Development

- Edit TypeScript files in the `src/` directory
- Run `npm run watch` to automatically compile changes
- Use `npm run push` to upload changes to Google Apps Script

## Custom Functions

The project includes several custom functions that can be used in Google Sheets:

- `HELLO(name)` - Returns a greeting message
- `MULTIPLY_BY_TWO(number)` - Multiplies a number by 2
- `CONCATENATE_WITH_PREFIX(prefix, text)` - Concatenates text with a prefix

## Usage in Google Sheets

After deploying, you can use these functions in your Google Sheets like any built-in function:

```
=HELLO("World")
=MULTIPLY_BY_TWO(5)
=CONCATENATE_WITH_PREFIX("Hello", "World")
```

## Google Apps Script Web App Deployment

For the QuickBooks integration to work, you need to deploy your Google Apps Script as a Web App:

### Step 1: Deploy as Web App

1. **Open your Google Apps Script project**:
   - Go to [script.google.com](https://script.google.com)
   - Find your project (or create one using `clasp create --type sheets`)

2. **Deploy the Web App**:
   - Click "Deploy" → "New deployment"
   - Choose type: "Web app"
   - Description: "QuickBooks Sync API"
   - Execute as: "Me"
   - Who has access: "Anyone" (for external API access)
   - Click "Deploy"

3. **Copy the Web App URL**:
   - After deployment, copy the URL (ends with `/exec`)
   - Example: `https://script.google.com/macros/s/AKfycbx.../exec`

### Step 2: Set Up API Key

1. **Run the setup function**:
   ```javascript
   // In your Google Apps Script editor, run this once:
   function setupQuickBooksIntegration() {
     const apiKey = Utilities.getUuid();
     PropertiesService.getScriptProperties().setProperty('QB_API_KEY', apiKey);
     console.log('API Key generated:', apiKey);
     return apiKey;
   }
   ```

2. **Copy the generated API key** from the execution log

### Step 3: Configure the Rust Service

Update your `config/config.toml`:

```toml
[google_sheets]
webapp_url = "https://script.google.com/macros/s/YOUR_ACTUAL_SCRIPT_ID/exec"
api_key = "YOUR_GENERATED_API_KEY"
spreadsheet_id = "YOUR_SPREADSHEET_ID"
sheet_name = "The Simple Buckets"
cell_address = "A1"
```

### Step 4: Test the Connection

```bash
cargo run --release -- --test-connection
```
