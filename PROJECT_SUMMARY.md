# QuickBooks Desktop Enterprise to Google Sheets Integration

## 🎯 **Project Overview**

This project enables automatic synchronization of QuickBooks Desktop Enterprise account balances to Google Sheets. Specifically designed to extract the balance of account **9445 (INCOME TAX)** and update it in a Google Sheets cell every hour.

## 🏗️ **Architecture**

```
QuickBooks Desktop Enterprise → Rust Service → Google Apps Script → Google Sheets
```

## 📦 **Components**

### 1. **Google Apps Script** (`src/Code.ts`)
- Receives data from Rust service via HTTP POST
- Updates Google Sheets cells with account balances
- Handles authentication and error logging
- Deployed as a Web App to receive external requests

### 2. **Rust Service** (`quickbooks-rust-service/`)
- Connects to QuickBooks Desktop Enterprise using COM interface
- Extracts account balance for specified account (9445)
- Sends data to Google Apps Script Web App
- Runs on configurable schedule (default: hourly)

## 🚀 **Setup Instructions**

### **Phase 1: Google Apps Script Setup**

1. **Access the updated code**:
   - Open Google Apps Script: https://script.google.com/d/1Xzoi9iquQrT_nJJxb_7y100v5LPzI4w4u2ICLpPoBO7-PzYGsJiMuZHh/edit
   - **Refresh the page** to see the latest functions
   - You should now see functions like `setupPermissions`, `checkPermissions`, etc.

2. **Set up permissions**:
   - In the editor, select `setupPermissions` from the function dropdown
   - Click the ▶️ "Run" button
   - When prompted, click "Review permissions" → "Allow"

3. **Generate API Key**:
   - Select `setupQuickBooksIntegration` from the dropdown
   - Click the ▶️ "Run" button
   - Copy the generated API key from the execution log

4. **Deploy as Web App**:
   - Click "Deploy" > "New Deployment"
   - Choose "Web app" type
   - Set "Execute as: Me" and "Who has access: Anyone"
   - Click "Deploy" and copy the Web App URL

5. **Test the setup**:
   - Select `testWebAppEndpoint` from the dropdown
   - Click the ▶️ "Run" button
   - Check that it shows success in the execution log

### **Phase 2: Rust Service Setup**

1. **Navigate to Rust service directory**:
   ```bash
   cd quickbooks-rust-service
   ```

2. **Run setup script**:
   ```bash
   ./setup.sh
   ```

3. **Configure the service**:
   Edit `config/config.toml`:
   ```toml
   [quickbooks]
   company_file = "C:\\Path\\To\\Your\\Company.qbw"
   account_number = "9445"
   account_name = "INCOME TAX"

   [google_sheets]
   webapp_url = "https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec"
   api_key = "YOUR_API_KEY_FROM_STEP_1"
   cell_address = "A1"
   sheet_name = "Sheet1"

   [schedule]
   cron_expression = "0 0 * * * *"  # Every hour
   ```

4. **Run the service**:
   ```bash
   cargo run --release
   ```

## 🔧 **Requirements**

### **Google Apps Script Side**
- ✅ Google Apps Script deployed as Web App
- ✅ API key configured for authentication
- ✅ Target Google Sheets document accessible

### **Rust Service Side (Windows)**
- QuickBooks Desktop Enterprise installed and running
- Company file accessible
- Admin permissions for QuickBooks
- Rust toolchain installed
- Network access to Google Apps Script

## 📊 **Data Flow**

1. **Rust service starts** and connects to QuickBooks
2. **Every hour**, the service:
   - Queries QuickBooks for account 9445 balance
   - Sends HTTP POST to Google Apps Script Web App
   - Google Apps Script updates the specified cell
   - Logs the update with timestamp

## 🔍 **Monitoring**

- **Rust service logs** show connection status and sync results
- **Google Apps Script logs** show received data and update status
- **Google Sheets** shows last update timestamp

## 🛠️ **Troubleshooting**

### **Service Feedback Commands**
```cmd
# Show help and all options
cargo run --release -- --help

# Show current configuration  
cargo run --release -- --show-config

# Run with detailed logging
cargo run --release -- --verbose

# Test connections only (with feedback)
cargo run --release -- --test-connection --verbose

# Test just QuickBooks (skip Google Sheets)
cargo run --release -- --test-connection --skip-sheets-test
```

### **Common Issues:**

1. **"Invalid API key"**: 
   - Open Google Apps Script editor
   - Select `setupQuickBooksIntegration` from dropdown, click ▶️ "Run"
   - Copy the generated API key to your config.toml

2. **"Sheet not found"**:
   - Select `checkPermissions` from dropdown, click ▶️ "Run"
   - Check the execution log for available sheet names
   - Update your config.toml with the correct sheet name

3. **"Permission denied"**:
   - Select `setupPermissions` from dropdown, click ▶️ "Run"
   - Click "Review permissions" → "Allow" when prompted

4. **"Failed to create QuickBooks session manager"**:
   - Install QuickBooks Desktop SDK
   - Use mock data for testing: set `company_file = "MOCK"`

5. **"STATUS_ACCESS_VIOLATION"**:
   - Run as Administrator
   - Install QuickBooks SDK properly
   - Use mock mode for testing

6. **"Page Not Found"**:
   - Deploy Google Apps Script as Web App
   - Use the correct Web App URL (ends with `/exec`)
   - Set permissions to "Anyone"

### **Mock Data Testing**
```toml
[quickbooks]
company_file = "MOCK"  # Uses simulated data for testing
```

## 📈 **Current Status**

✅ **Google Apps Script**: Deployed and ready  
✅ **Rust Service**: Built and configured  
🔄 **Next**: Configure and test on Windows with QuickBooks  

## 🎯 **Expected Result**

Once running, your Google Sheets will automatically update with the current balance of QuickBooks account 9445 (INCOME TAX) every hour, with timestamps showing when each update occurred.

---

**Ready to deploy!** Follow the setup instructions above to get your QuickBooks data flowing into Google Sheets automatically.
