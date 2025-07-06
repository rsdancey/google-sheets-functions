# Windows File Path Configuration Guide

## The Problem

When configuring QuickBooks file paths on Windows, you might encounter errors like:
```
invalid escape sequence
TOML parse error at line 6, column 20
```

This happens because Windows uses backslashes (`\`) in file paths, but TOML (and many programming languages) treat backslashes as escape characters.

## ✅ Solutions

### **Option 1: Let the App Auto-Convert (Recommended)**

The QuickBooks sync service now automatically converts single backslashes to double backslashes before parsing the TOML configuration.

**You can write:**
```toml
[quickbooks]
company_file = "C:\Users\YourName\Documents\QuickBooks\Company.qbw"
```

**The app will automatically convert it to:**
```toml
[quickbooks] 
company_file = "C:\\Users\\YourName\\Documents\\QuickBooks\\Company.qbw"
```

### **Option 2: Use Double Backslashes (Manual)**

If you prefer to escape the paths yourself:

```toml
[quickbooks]
company_file = "C:\\Users\\YourName\\Documents\\QuickBooks\\Company.qbw"
```

### **Option 3: Use Forward Slashes**

Windows also accepts forward slashes in many contexts:

```toml
[quickbooks]
company_file = "C:/Users/YourName/Documents/QuickBooks/Company.qbw"
```

### **Option 4: Use Raw Strings (TOML)**

TOML supports literal strings that don't process escape sequences:

```toml
[quickbooks]
company_file = '''C:\Users\YourName\Documents\QuickBooks\Company.qbw'''
```

## 🚀 **Recommended Configuration**

For most users, just use the normal Windows path with single backslashes:

```toml
[quickbooks]
# The app will handle the backslashes automatically
company_file = "C:\Users\YourName\Documents\QuickBooks\Your Company.qbw"
account_number = "9445"
account_name = "INCOME TAX"

[google_sheets]
webapp_url = "https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec"
api_key = "YOUR_API_KEY"
spreadsheet_id = "YOUR_SPREADSHEET_ID"
sheet_name = "The Simple Buckets"
cell_address = "A1"

[schedule]
cron_expression = "0 0 * * * *"  # Every hour
```

## 🔍 **Special Cases**

### **For Currently Open File**
```toml
company_file = "AUTO"  # Uses whatever file is open in QuickBooks
```

### **For Mock Testing**
```toml
company_file = "MOCK"  # Uses simulated data for testing
```

### **For File Selection Dialog**
```toml
company_file = ""  # Will prompt to select from available files
```

## ⚡ **Testing Your Configuration**

After setting up your configuration, test it:

```cmd
# Show your current configuration
cargo run --release -- --show-config

# Test the QuickBooks connection
cargo run --release -- --test-connection --skip-sheets-test --verbose
```

The service will log when it processes Windows paths:
```
INFO: Preprocessing Windows path: 'C:\Users\...' -> 'C:\\Users\\...'
```

## 🛠️ **Troubleshooting**

### **Error: "TOML parse error"**
- The automatic conversion didn't work
- Try using double backslashes manually: `C:\\Users\\...`
- Or use forward slashes: `C:/Users/...`

### **Error: "File not found"**
- Check that the file path exists
- Make sure QuickBooks is installed
- Verify the company file is accessible

### **Error: "Permission denied"**
- Run as Administrator
- Check QuickBooks application permissions
- Ensure the company file allows external access

The automatic path conversion makes it much easier to configure Windows file paths - just copy and paste the path from Windows Explorer!
