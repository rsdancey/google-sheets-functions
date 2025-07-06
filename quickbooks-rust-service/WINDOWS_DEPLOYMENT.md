# Windows Server 2019 Deployment Guide

## 🖥️ **Windows Server 2019 Specific Considerations**

### **Important Server Environment Differences**
Windows Server 2019 has different characteristics that affect our QuickBooks integration:

1. **Service Environment**: Typically runs as a Windows Service
2. **User Session Limitations**: May not have interactive desktop sessions  
3. **QuickBooks Desktop Installation**: Different licensing and installation requirements
4. **COM Security**: More restrictive COM security settings
5. **Network Service Account**: May run under different user contexts

### **Critical Server Challenges**
- **QuickBooks Desktop** is primarily designed for desktop environments, not servers
- **Interactive sessions** required for QuickBooks may not be available
- **COM automation** has stricter security in server environments
- **Service account permissions** need careful configuration

## 🔧 **Server-Specific Code Adjustments Needed**

### **1. Enhanced COM Initialization**
The current code needs server-specific COM initialization:

```rust
#[cfg(windows)]
fn initialize_com_for_server() -> Result<()> {
    unsafe {
        // Use COINIT_MULTITHREADED for server environments
        CoInitializeEx(
            None,
            COINIT_MULTITHREADED | COINIT_DISABLE_OLE1DDE,
        ).map_err(|e| anyhow::anyhow!("Failed to initialize COM: {}", e))?;
        
        // Set COM security for server environment
        CoInitializeSecurity(
            None,                        // pSecDesc
            -1,                         // cAuthSvc
            None,                       // asAuthSvc
            None,                       // pReserved1
            RPC_C_AUTHN_LEVEL_NONE,     // dwAuthnLevel
            RPC_C_IMP_LEVEL_IMPERSONATE, // dwImpLevel
            None,                       // pAuthList
            EOAC_NONE,                  // dwCapabilities
            None,                       // pReserved3
        ).map_err(|e| anyhow::anyhow!("Failed to set COM security: {}", e))?;
    }
    Ok(())
}
```

### **2. Windows Service Integration**
For server deployment, the application should run as a Windows Service:

```toml
# Additional dependencies needed for server
[target.'cfg(windows)'.dependencies]
windows = { version = "0.58", features = [
    "Win32_System_Com",
    "Win32_System_Ole", 
    "Win32_Foundation",
    "Win32_System_Variant",
    "Win32_System_Com_StructuredStorage",
    "Win32_System_Threading",
    "Win32_Security",
    "Win32_Globalization",
    "Win32_System_Services",           # Added for service detection
    "Win32_Security_Authorization",    # Added for user impersonation
    "Win32_System_WindowsProgramming", # Added for server utilities
    "implement",
] }
windows-service = "0.6"  # For Windows Service integration
```

## 🎯 **Recommended Server Deployment Approach**

### **Option 1: QuickBooks Online API (Recommended for Server)**
For Windows Server 2019, I **strongly recommend** switching to QuickBooks Online API instead of Desktop:

**Advantages:**
- ✅ **Server-friendly**: Designed for automated/headless operations
- ✅ **REST API**: No COM automation complexity
- ✅ **Better security**: OAuth 2.0 authentication
- ✅ **Cloud-based**: No local QuickBooks installation needed
- ✅ **More reliable**: Better for scheduled operations

**Implementation:**
```rust
// QuickBooks Online API approach
use reqwest::Client;
use serde_json::Value;

struct QuickBooksOnlineClient {
    client: Client,
    access_token: String,
    company_id: String,
}

impl QuickBooksOnlineClient {
    async fn get_accounts(&self) -> Result<Vec<Account>> {
        let url = format!(
            "https://quickbooks-api.intuit.com/v3/company/{}/accounts",
            self.company_id
        );
        
        let response = self.client
            .get(&url)
            .header("Authorization", format!("Bearer {}", self.access_token))
            .header("Accept", "application/json")
            .send()
            .await?;
            
        // Parse and return accounts
        Ok(accounts)
    }
}
```

### **Option 2: Desktop with Service Account Setup**
If you must use QuickBooks Desktop on Server 2019:

**Requirements:**
1. **Interactive User Session**: Set up a dedicated user account that stays logged in
2. **QuickBooks Service Mode**: Configure QuickBooks to run as a service
3. **Scheduled Tasks**: Use Windows Task Scheduler to run under user context
4. **Service Account**: Create dedicated service account with appropriate permissions

**Service Configuration:**
```cmd
# Run as scheduled task with user context
schtasks /create /tn "QuickBooks Sync Service" /tr "C:\path\to\qb_sync.exe" /sc minute /mo 15 /ru "DOMAIN\serviceaccount" /rp "password"
```

## 📊 **Server Configuration Recommendations**

### **Enhanced Configuration for Server**
```toml
# config.toml for Windows Server 2019
[quickbooks]
enabled = true
company_file = "C:\\QuickBooks\\Company.qbw"  # Full path required
application_name = "QB Sheets Sync Service"
connection_timeout = 30000  # Longer timeout for server
retry_attempts = 3
service_mode = true  # New flag for server behavior
user_session_required = true  # Indicates need for user session

[server]
run_as_service = true
service_account = "QB_SERVICE_ACCOUNT"
log_to_event_log = true
health_check_port = 8080
```

### **Logging Configuration**
```rust
// Server-appropriate logging
fn configure_server_logging() -> Result<()> {
    env_logger::Builder::from_default_env()
        .target(env_logger::Target::Stdout)  // For service logs
        .filter_level(log::LevelFilter::Info)
        .format_timestamp_secs()
        .init();
    
    // Also log to Windows Event Log for server monitoring
    Ok(())
}
```

## 🚨 **Important Server Considerations**

### **QuickBooks Desktop Limitations on Server**
- **Not officially supported** for server environments
- **Requires interactive user session** which servers typically don't have
- **Licensing issues** may arise with server installations
- **COM automation** is less reliable in service contexts

### **Alternative Approaches**
1. **QuickBooks Online API** - Most reliable for server deployment
2. **Desktop with RDP session** - Keep user logged in via Remote Desktop
3. **Scheduled task approach** - Run under user context periodically
4. **Hybrid approach** - Process on desktop, sync to server

## 💡 **My Recommendation**

For Windows Server 2019 deployment, I **strongly recommend**:

1. **Switch to QuickBooks Online API** for the most reliable server deployment
2. **If Desktop required**: Use scheduled tasks with dedicated user account
3. **Implement comprehensive logging** for server monitoring
4. **Add health check endpoints** for service monitoring
5. **Use configuration files** instead of environment variables

Would you like me to:
1. **Implement QuickBooks Online API integration** (recommended)
2. **Add Windows Service support** for the current Desktop approach
3. **Create server-specific configuration** options

The QuickBooks Online API would be much more suitable for your Windows Server 2019 environment!

## 🚀 **Deployment Steps**

### **Step 1: Install Rust on Windows**
```cmd
# Download from https://rustup.rs/ or use winget
winget install Rustlang.Rust.MSVC
```

### **Step 2: Install QuickBooks SDK**
1. Download from: https://developer.intuit.com/app/developer/qbdesktop/docs/get-started
2. Run the installer as Administrator
3. Follow the installation wizard

### **Step 3: Clone and Build the Project**
```cmd
git clone <your-repo-url>
cd quickbooks-rust-service
cargo build --release
```

### **Step 4: Configure for Production**
Edit `config/config.toml`:
```toml
[quickbooks]
# Use currently open QuickBooks file
company_file = "AUTO"
account_number = "9445"
account_name = "INCOME TAX"

# Optional: Add authentication if needed
# company_file_password = ""  # If your .qbw file has a password
# qb_username = ""           # For multi-user files
# qb_password = ""           # For multi-user files

[google_sheets]
# Your actual Google Apps Script Web App URL
webapp_url = "https://script.google.com/macros/s/YOUR_ACTUAL_SCRIPT_ID/exec"
# Your actual API key from setupQuickBooksIntegration()
api_key = "YOUR_ACTUAL_API_KEY"
# Your actual spreadsheet ID
spreadsheet_id = "YOUR_ACTUAL_SPREADSHEET_ID"
# Your actual sheet name
sheet_name = "The Simple Buckets"  # or whatever your sheet is named
cell_address = "A1"

[schedule]
# Run every hour
cron_expression = "0 0 * * * *"
```

### **Step 5: Test the Connection**
```cmd
# Test QuickBooks connection only
cargo run --release -- --test-connection --skip-sheets-test --verbose

# Test both QuickBooks and Google Sheets
cargo run --release -- --test-connection --verbose

# Run a single sync test
cargo run --release -- --test-connection
```

### **Step 6: Run the Service**
```cmd
# Run normally (will sync every hour)
cargo run --release

# Run with verbose logging
cargo run --release -- --verbose
```

## 🔧 **Troubleshooting**

### **"QuickBooks SDK not found"**
- Install the QuickBooks SDK as Administrator
- Restart your command prompt after installation

### **"Company file not found"**
- Make sure QuickBooks is running
- Ensure your company file is open
- Try specifying the full path to your .qbw file

### **"Access denied to account 9445"**
- Check that the account exists in your Chart of Accounts
- Ensure the QuickBooks user has access to this account
- Verify the account number is exactly "9445"

### **"Permission denied"**
- Run the command prompt as Administrator
- Check QuickBooks application permissions
- Ensure the QuickBooks company file allows external applications

## 🎯 **Account 9445 Specific Setup**

Your account structure is:
```
1000 (Parent Account)
└── BoA Accounts (Sub-account, no number)
    └── 9445 - INCOME TAX (Sub-sub-account)
```

The service will:
1. Query QuickBooks for account number "9445"
2. QuickBooks will automatically find it in the hierarchy
3. Return the balance for just that account (not aggregated)
4. Send the balance to your Google Sheets

## 📊 **Expected Results**

Once running with real QuickBooks data:
- You'll see the actual balance of account 9445
- The balance will be current as of the last QuickBooks transaction
- Updates will happen every hour (or whatever schedule you set)
- Google Sheets will show real-time QuickBooks data

## 🔄 **Service Management**

### **Run as Windows Service (Optional)**
To run continuously in the background:

1. Install a service wrapper like NSSM
2. Configure it to run your Rust executable
3. Set it to start automatically with Windows

### **Monitoring**
- Check the console output for sync status
- Monitor Google Sheets for updates
- Review Google Apps Script logs for any errors

## 🛡️ **Security Considerations**

- Keep your API key secure
- Use Windows firewall to restrict network access
- Run with minimal required privileges
- Backup your QuickBooks data regularly

Once deployed on Windows with QuickBooks running, you'll get the real balance of account 9445 instead of mock data!
