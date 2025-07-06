# QuickBooks SDK Installation Guide

## Problem: "Failed to create QuickBooks session manager"

This error means the QuickBooks SDK is not installed or not properly registered.

## Quick Solution

### 1. Download QuickBooks SDK

**For QuickBooks Desktop Enterprise v24:**
- Go to [Intuit Developer Downloads](https://developer.intuit.com/app/developer/qbdesktop/docs/get-started/download-and-install-the-sdk)
- Download "QuickBooks Desktop SDK v13 or later"
- Look for **QBFC17** (QuickBooks 2024 SDK)

### 2. Install SDK

1. **Run installer as Administrator**
2. **Accept all defaults**
3. **Restart computer after installation**

### 3. Verify Installation

Check if these files exist:
```
C:\Program Files\Common Files\Intuit\QuickBooks\QBFC17.DLL
C:\Program Files\Common Files\Intuit\QuickBooks\QBSessionManager.exe
```

### 4. Register COM Objects

If SDK is installed but still getting errors:

```cmd
# Run Command Prompt as Administrator
cd "C:\Program Files\Common Files\Intuit\QuickBooks"
regsvr32 QBFC17.DLL
regsvr32 /u QBFC17.DLL
regsvr32 QBFC17.DLL
```

### 5. Test Again

1. **Start QuickBooks Desktop Enterprise**
2. **Open your company file**
3. **Keep QuickBooks running**
4. **Run the service**:
   ```cmd
   cargo run --release -- --test-connection
   ```

## Alternative: Mock Data Test

To test if the issue is SDK-related, temporarily use mock data:

1. **Edit config.toml**:
   ```toml
   [quickbooks]
   company_file = "MOCK"
   ```

2. **Run service**:
   ```cmd
   cargo run --release
   ```

If it works with mock data, the issue is definitely QuickBooks SDK installation.

## Common Issues

### "Access Denied" during SDK installation
- **Solution**: Run installer as Administrator
- **Alternative**: Disable antivirus temporarily during installation

### "QBFC17.DLL not found"
- **Solution**: SDK not installed or wrong version
- **Fix**: Download correct SDK version for your QuickBooks

### "COM object not registered"
- **Solution**: Run the regsvr32 commands above
- **Alternative**: Reinstall the SDK

### "QuickBooks not running"
- **Solution**: Start QuickBooks Desktop before running the service
- **Requirement**: Company file must be open

## **Critical Error: STATUS_ACCESS_VIOLATION (0xc0000005)**

This error indicates a memory access violation on Windows, typically related to COM interface issues.

### **Immediate Fix: Use Mock Data Mode**

1. **Edit your config.toml**:
   ```toml
   [quickbooks]
   company_file = "MOCK"
   account_number = "9445"
   account_name = "INCOME TAX"
   ```

2. **Test the service**:
   ```cmd
   cargo run --release -- --test-connection
   ```

This will bypass the QuickBooks COM interface and use simulated data.

### **Root Causes of Access Violation**

1. **QuickBooks SDK not properly installed**
2. **COM interface permissions issue** 
3. **32-bit/64-bit architecture mismatch**
4. **QuickBooks not running as Administrator**
5. **Corrupted QuickBooks installation**

### **Permanent Solutions**

#### **Solution 1: Install QuickBooks SDK Properly**
1. **Download correct SDK version**:
   - For QB Enterprise v24: QuickBooks SDK v13+ (QBFC17)
   - Download from [Intuit Developer](https://developer.intuit.com/)

2. **Install as Administrator**:
   ```cmd
   # Run installer as Administrator
   # Restart computer after installation
   ```

3. **Register COM objects**:
   ```cmd
   # Run Command Prompt as Administrator
   cd "C:\Program Files\Common Files\Intuit\QuickBooks"
   regsvr32 QBFC17.DLL
   ```

#### **Solution 2: Run with Proper Permissions**
1. **Run QuickBooks as Administrator**
2. **Run Command Prompt as Administrator**
3. **Run the Rust service from that elevated prompt**:
   ```cmd
   cd C:\path\to\your\service
   cargo run --release
   ```

#### **Solution 3: Check Architecture Compatibility**
1. **Check QuickBooks architecture**:
   - Task Manager → Details → Look for QuickBooks process
   - Note if it's 32-bit or 64-bit

2. **Match Rust compilation target**:
   ```cmd
   # For 32-bit QuickBooks
   rustup target add i686-pc-windows-msvc
   cargo build --release --target i686-pc-windows-msvc
   
   # For 64-bit QuickBooks (default)
   cargo build --release --target x86_64-pc-windows-msvc
   ```

### **Debugging Steps**

#### **Step 1: Test Mock Mode First**
```toml
[quickbooks]
company_file = "MOCK"
```
If this works, the issue is QuickBooks COM interface.

#### **Step 2: Check QuickBooks Installation**
1. **Open QuickBooks Desktop Enterprise**
2. **Open your company file manually**
3. **Verify it works normally**

#### **Step 3: Check Windows Event Logs**
1. **Open Event Viewer** (eventvwr.msc)
2. **Check Application logs** for detailed error info
3. **Look for COM/QuickBooks related errors**

#### **Step 4: Try Different QuickBooks Modes**
1. **Single-user mode**: File → Switch to Single-user Mode
2. **Administrator mode**: Run QuickBooks as Administrator
3. **Compatibility mode**: Right-click QB → Properties → Compatibility

### **Quick Test Commands**

```cmd
# Test with mock data (should work)
cargo run --release -- --test-connection

# Test with skip sheets (to isolate QB issues)
cargo run --release -- --test-connection --skip-sheets-test
```

### **If All Else Fails**

**Use Mock Mode for Testing**:
- Edit config.toml: `company_file = "MOCK"`
- Test Google Sheets integration
- Fix QuickBooks issues separately

The access violation is a Windows/COM issue, not a Rust code issue. Mock mode will let you test the rest of the system while you resolve the QuickBooks SDK installation.
