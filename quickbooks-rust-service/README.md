# QuickBooks to Google Sheets Sync Service

This Rust service connects to QuickBooks Desktop Enterprise and syncs account data to Google Sheets.

## Features

- Connects to QuickBooks Desktop Enterprise using COM interface
- Extracts account balance for account 9445 (INCOME TAX)
- Sends data to Google Sheets via Apps Script Web App
- Runs on hourly schedule
- Comprehensive error handling and logging

## Setup

### **Prerequisites (Windows)**
1. **Windows 10/11** with QuickBooks Desktop Enterprise
2. **Visual Studio Build Tools** or Visual Studio Community
   - Download: https://visualstudio.microsoft.com/downloads/
   - Include: Windows 10/11 SDK, MSVC v143 compiler
3. **Rust toolchain**: stable-x86_64-pc-windows-msvc

### **Installation Steps**

1. **Install Rust** (if not already installed):
   ```cmd
   curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh
   ```

2. **Configure Rust for Windows**:
   ```cmd
   rustup toolchain install stable-x86_64-pc-windows-msvc
   rustup default stable-x86_64-pc-windows-msvc
   ```

3. **Build the service**:
   ```cmd
   cargo clean
   cargo build --release
   ```

4. **Configure the service**:
   - Copy `config/config.example.toml` to `config/config.toml`
   - Update the configuration with your Google Apps Script URL and API key

5. **Run the service**:
   ```cmd
   cargo run --release
   ```

### **If Build Fails**
See [WINDOWS_BUILD.md](WINDOWS_BUILD.md) for detailed troubleshooting.

## Configuration

The service uses a TOML configuration file at `config/config.toml`:

```toml
[quickbooks]
# How to connect to QuickBooks company file:
# Option 1: Specific file path
company_file = "C:\\Path\\To\\Your\\Company.qbw"
# Option 2: Use currently open file (recommended)
company_file = "AUTO"
# Option 3: Let QuickBooks prompt for file selection
company_file = ""

account_number = "9445"
account_name = "INCOME TAX"

[google_sheets]
# Your Google Apps Script Web App URL
webapp_url = "https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec"
# API key for authentication
api_key = "your-api-key-here"
# Google Sheets Document ID (the spreadsheet file itself)
# Get this from the URL: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
spreadsheet_id = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"
# Sheet name (the tab within the spreadsheet)
sheet_name = "Sheet1"
# Cell address where the balance will be updated
cell_address = "A1"

[schedule]
cron_expression = "0 0 * * * *"  # Every hour
```

### **Google Sheets Configuration Explained**

- **`spreadsheet_id`**: The ID of the Google Sheets **document** (the file itself)
  - Found in the URL: `https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit`
  - Example: `1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms`

- **`sheet_name`**: The name of the **tab** within the spreadsheet document
  - Examples: "Sheet1", "Income Tax", "QB Data"
  - If not specified, uses the active/first sheet

- **`cell_address`**: The specific cell where the account balance will be written
  - Examples: "A1", "B5", "C10"

## QuickBooks Connection Methods

### **Method 1: Specific Company File Path**
```toml
company_file = "C:\\Users\\YourUser\\Documents\\QuickBooks\\Your Company.qbw"
```
- **Pros**: Direct, no user interaction required
- **Cons**: Must know exact file path, file must exist
- **Use when**: Running as automated service with known file location

### **Method 2: Currently Open Company File (Recommended)**
```toml
company_file = "AUTO"
```
- **Pros**: Works with whatever company file is currently open in QuickBooks
- **Cons**: QuickBooks must be running with a company file open
- **Use when**: QuickBooks is manually opened by user each day

### **Method 3: Interactive File Selection**
```toml
company_file = ""
```
- **Pros**: User can choose from available company files
- **Cons**: Requires user interaction, not suitable for automation
- **Use when**: Multiple company files available, manual operation

## Prerequisites

1. **QuickBooks Desktop Enterprise** must be running
2. **Company file** must be accessible (either open or at specified path)
3. **User permissions** must allow external application access
4. **Network connectivity** to Google Apps Script

## Authentication & Security

### **QuickBooks Authentication Layers**

#### **1. Windows User Authentication**
- The service runs under a **Windows user account**
- That user must have **QuickBooks permissions**
- No configuration needed - uses current Windows user

#### **2. QuickBooks Application Permissions**
- **First time**: QuickBooks will prompt to allow external applications
- **User must click "Yes"** to allow the connection
- **Subsequent runs**: Connection is remembered (if saved)

#### **3. Company File Authentication (Optional)**
```toml
# If your company file has a password
company_file_password = "your-file-password"
```

#### **4. QuickBooks User Authentication (Multi-User Files)**
```toml
# For multi-user QuickBooks files
qb_username = "your-qb-username"
qb_password = "your-qb-password"
```

### **Authentication Flow**
1. Service starts as Windows user
2. Service connects to QuickBooks application
3. QuickBooks prompts for application permission (first time)
4. Service opens company file (with file password if needed)
5. Service authenticates as QuickBooks user (if multi-user)
6. Service extracts account data

### **Security Best Practices**
- Run service as **dedicated Windows user** with minimal permissions
- Use **environment variables** for sensitive passwords:
  ```bash
  set QB_FILE_PASSWORD=your-password
  set QB_USER_PASSWORD=your-user-password
  ```
- Enable **QuickBooks audit trail** to track access
- Use **HTTPS** for all Google Sheets communication

## QuickBooks Desktop Enterprise v24 Specific Notes

### **SDK Version**
- QuickBooks Desktop Enterprise v24 (2024) uses **QBFC17.QBSessionManager**
- The service automatically detects and tries the correct SDK version
- Falls back to older versions if needed

### **First-Time Setup**
1. **Run as Administrator** (required for initial COM registration)
2. **Allow External Applications** in QuickBooks:
   - Go to Edit → Preferences → Integrated Applications
   - Check "Allow integrated applications to access QuickBooks"
3. **Company File Access**:
   - For single-user files: QuickBooks must be running with the company file open
   - For multi-user files: Configure QB username/password in config

### **Multi-User Enterprise Files**
QuickBooks Enterprise v24 often uses multi-user setup:
```toml
[quickbooks]
company_file = "\\\\server\\path\\to\\company.qbw"  # Network path
qb_username = "your_qb_username"
qb_password = "your_qb_password"  # Or use environment variable
```

## Windows Service Installation

To run as a Windows service, you can use tools like `nssm` or `winsw`.

## Security Notes

- Store sensitive configuration in environment variables
- Use HTTPS for all API communications
- Implement proper QuickBooks user authentication
