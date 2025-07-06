# QuickBooks Authentication Guide

## 🔐 **How QuickBooks Authentication Works**

QuickBooks Desktop Enterprise uses **multiple layers of authentication**. Here's exactly how the service handles credentials:

## **Authentication Layers**

### **Layer 1: Windows User Authentication**
- **What it is**: The service runs under a Windows user account
- **How it works**: Uses the **currently logged-in Windows user** or **service account**
- **Configuration**: No configuration needed - automatic
- **Requirements**: 
  - Windows user must have **QuickBooks permissions**
  - User must be able to **launch QuickBooks**

### **Layer 2: QuickBooks Application Permission**
- **What it is**: QuickBooks asks if external applications can connect
- **How it works**: 
  - **First time**: QuickBooks shows a dialog: "Allow [Application] to access QuickBooks?"
  - **User must click "Yes"** and optionally "Don't ask again"
  - **Subsequent times**: Connects automatically (if permission was saved)
- **Configuration**: No configuration needed - interactive prompt
- **Requirements**: User must be present for first-time setup

### **Layer 3: Company File Password (Optional)**
- **What it is**: Some company files are password-protected
- **How it works**: Service provides password when opening the file
- **Configuration**: 
  ```toml
  company_file_password = "your-file-password"
  ```
  Or environment variable:
  ```bash
  set QB_FILE_PASSWORD=your-password
  ```

### **Layer 4: QuickBooks User Authentication (Multi-User)**
- **What it is**: Multi-user QuickBooks files require user login
- **How it works**: Service logs in as a specific QuickBooks user
- **Configuration**:
  ```toml
  qb_username = "your-qb-username"
  qb_password = "your-qb-password"
  ```
  Or environment variables:
  ```bash
  set QB_USERNAME=your-username
  set QB_USER_PASSWORD=your-password
  ```

## **Real-World Scenarios**

### **Scenario 1: Single-User QuickBooks (Most Common)**
```toml
[quickbooks]
company_file = "AUTO"
account_number = "9445"
account_name = "INCOME TAX"
# No passwords needed - QuickBooks runs under current Windows user
company_file_password = ""
qb_username = ""
qb_password = ""
```

**What happens:**
1. Service runs as current Windows user
2. Service connects to QuickBooks (permission prompt first time)
3. Service uses currently open company file
4. No additional authentication needed

### **Scenario 2: Password-Protected Company File**
```toml
[quickbooks]
company_file = "C:\\QuickBooks\\SecureCompany.qbw"
company_file_password = "MyCompanyPassword123"
# OR use environment variable QB_FILE_PASSWORD
```

**What happens:**
1. Service runs as current Windows user
2. Service connects to QuickBooks
3. Service opens specific company file with password
4. No QuickBooks user authentication needed

### **Scenario 3: Multi-User QuickBooks on Network**
```toml
[quickbooks]
company_file = "\\\\SERVER\\QuickBooks\\Company.qbw"
qb_username = "ServiceUser"
qb_password = "ServicePassword123"
# OR use environment variables QB_USERNAME and QB_USER_PASSWORD
```

**What happens:**
1. Service runs as Windows user
2. Service connects to QuickBooks
3. Service opens company file from network
4. Service logs in as QuickBooks user "ServiceUser"

### **Scenario 4: Maximum Security (All Layers)**
```toml
[quickbooks]
company_file = "\\\\SERVER\\QuickBooks\\SecureCompany.qbw"
company_file_password = "FilePassword123"
qb_username = "ServiceUser"
qb_password = "UserPassword123"
```

**What happens:**
1. Service runs as Windows user
2. Service connects to QuickBooks
3. Service opens password-protected company file
4. Service logs in as QuickBooks user

## **Security Best Practices**

### **1. Use Environment Variables for Passwords**
Instead of putting passwords in the config file:
```bash
# Windows Command Prompt
set QB_FILE_PASSWORD=your-file-password
set QB_USERNAME=your-username
set QB_USER_PASSWORD=your-user-password
set SHEETS_API_KEY=your-api-key

# Then run the service
cargo run --release
```

### **2. Run as Dedicated Service Account**
1. Create a dedicated Windows user: `QBSyncService`
2. Give it minimal permissions
3. Install QuickBooks under this account
4. Run the service as this account

### **3. QuickBooks User Permissions**
Create a dedicated QuickBooks user with minimal permissions:
- **Read-only access** to financial data
- **No editing permissions**
- **Specific account access** only

### **4. Network Security**
- Use **HTTPS** for Google Sheets communication
- Store **API keys** in environment variables
- Enable **QuickBooks audit trail**

## **First-Time Setup Process**

### **Step 1: Windows User Setup**
1. Log in as the Windows user that will run the service
2. Make sure this user can launch QuickBooks
3. Open QuickBooks and your company file manually (first time)

### **Step 2: QuickBooks Permission**
1. Run the service for the first time
2. QuickBooks will show: "Allow Rust Service to access QuickBooks?"
3. Click **"Yes"** and **"Don't ask again"**
4. Service will now connect automatically

### **Step 3: Test Authentication**
1. Service will log connection attempts
2. Check logs for authentication success/failure
3. Verify account data is extracted correctly

## **Troubleshooting Authentication Issues**

### **"QuickBooks connection failed"**
- ✅ Make sure QuickBooks is running
- ✅ Check if company file is open (for AUTO mode)
- ✅ Verify Windows user has QuickBooks permissions

### **"Access denied"**
- ✅ Run QuickBooks as Administrator
- ✅ Check company file permissions
- ✅ Verify QuickBooks user has account access

### **"Invalid password"**
- ✅ Check company file password
- ✅ Verify QuickBooks user credentials
- ✅ Ensure passwords match exactly

### **"Application not authorized"**
- ✅ Delete QuickBooks application permissions
- ✅ Run service again to get fresh permission prompt
- ✅ Make sure to click "Don't ask again"

## QuickBooks Desktop Enterprise v24 Authentication

### **Enterprise-Specific Considerations**

**Multi-User Environment**: QuickBooks Enterprise v24 is typically deployed in multi-user environments with:
- Network-based company files
- User-based access control
- Administrator-managed permissions

**Required Permissions**:
1. **Windows User**: Must have local administrator rights for initial setup
2. **QuickBooks User**: Must have appropriate permissions in the company file
3. **Network Access**: If company file is on a network share

### **Authentication Flow for Enterprise v24**

1. **Initial Setup** (Run as Administrator):
   ```cmd
   # Run PowerShell as Administrator
   cd C:\path\to\rust\service
   .\target\release\qb_sync.exe --test-connection
   ```

2. **QuickBooks Permission Dialog**: 
   - First run will show "External Application Access" dialog
   - Choose "Yes, always allow access" for automated operation
   - This registers the application with QuickBooks

3. **Company File Authentication**:
   ```toml
   [quickbooks]
   company_file = "\\\\fileserver\\quickbooks\\YourCompany.qbw"
   company_file_password = "file_password"  # If file is password-protected
   qb_username = "your_qb_username"
   qb_password = "your_qb_password"
   ```

4. **Environment Variables** (Recommended for production):
   ```cmd
   set QB_FILE_PASSWORD=your_file_password
   set QB_USER_PASSWORD=your_qb_user_password
   ```

### **Troubleshooting Enterprise v24**

**"Application not authorized"**: 
- Run QuickBooks as Administrator
- Go to Edit → Preferences → Integrated Applications
- Add the application manually if needed

**Network file access issues**:
- Use UNC paths: `\\\\server\\share\\file.qbw`
- Ensure Windows user has network access
- Test file access manually first

**Multi-user conflicts**:
- Ensure exclusive access during sync operations
- Consider scheduling during off-hours
- Use QuickBooks user with appropriate permissions

## **Summary**

The service handles QuickBooks authentication through:
1. **Windows user** (automatic)
2. **QuickBooks application permission** (one-time prompt)
3. **Company file password** (optional, configured)
4. **QuickBooks user login** (optional, for multi-user files)

Most users only need to handle the **one-time permission prompt** - everything else is automatic!

## **Common Error: "Failed to create QuickBooks session manager"**

### **What This Error Means**
```
Caused by: Failed to create QuickBooks session manager with any supported version
```

This error indicates that the **QuickBooks SDK** is not installed or not properly registered on your Windows system.

### **Root Causes**
1. **QuickBooks SDK not installed** - The SDK is separate from QuickBooks Desktop
2. **COM objects not registered** - The SDK components aren't accessible
3. **QuickBooks not running** - The application must be running for COM access
4. **Wrong QuickBooks version** - SDK version mismatch

### **Solution Steps**

#### **Step 1: Install QuickBooks SDK**
1. **Download the SDK**:
   - Go to [Intuit Developer](https://developer.intuit.com/)
   - Download "QuickBooks Desktop SDK"
   - Choose the version that matches your QuickBooks Desktop Enterprise v24

2. **Install the SDK**:
   - Run the installer **as Administrator**
   - Accept all default settings
   - Restart your computer after installation

#### **Step 2: Verify SDK Installation**
1. **Check Registry Entries**:
   - Press `Win + R`, type `regedit`
   - Navigate to: `HKEY_CLASSES_ROOT\QBFC17.QBSessionManager`
   - If this key exists, the SDK is installed

2. **Check Program Files**:
   - Look for: `C:\Program Files\Common Files\Intuit\QuickBooks\QBFC17.DLL`
   - If this file exists, the SDK is properly installed

#### **Step 3: Register COM Objects (if needed)**
If the SDK is installed but COM objects aren't registered:

```cmd
# Run Command Prompt as Administrator
cd "C:\Program Files\Common Files\Intuit\QuickBooks"
regsvr32 QBFC17.DLL
```

#### **Step 4: Start QuickBooks Desktop**
1. **Launch QuickBooks Desktop Enterprise**
2. **Open your company file**
3. **Keep QuickBooks running** (don't close it)
4. **Try the service again**

#### **Step 5: Test with Different SDK Versions**
The service tries these versions in order:
- QBFC17 (QuickBooks 2024)
- QBFC16 (QuickBooks 2016+)
- QBFC15 (QuickBooks 2015)
- QBFC14 (QuickBooks 2014)
- QBFC13 (QuickBooks 2013)

If you have an older QuickBooks version, you might need an older SDK.

### **Enterprise v24 Specific Notes**

**For QuickBooks Desktop Enterprise v24:**
1. **Use QBFC17 SDK** (QuickBooks 2024 SDK)
2. **Install SDK after QuickBooks** - Install QuickBooks first, then the SDK
3. **Administrator Rights Required** - Both QuickBooks and the service need admin rights initially

### **Alternative: Test Without SDK**
To test if the issue is SDK-related, you can temporarily modify the service to skip QuickBooks connection:

1. **Edit your config.toml**:
   ```toml
   [quickbooks]
   company_file = "MOCK"  # This will use mock data
   ```

2. **Run the service**:
   ```cmd
   cargo run --release
   ```

If the service runs successfully with mock data, the issue is definitely the QuickBooks SDK.

### **Getting Help**
If you continue having issues:
1. **Check QuickBooks version**: Help → About QuickBooks
2. **Check Windows version**: `winver` command
3. **Check if QuickBooks is 32-bit or 64-bit**
4. **Download matching SDK version**

The SDK version must match your QuickBooks Desktop version for COM access to work.
