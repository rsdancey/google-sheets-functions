# Windows Build Guide

## 🛠️ **Common Windows Build Issues and Solutions**

### **Issue 1: Missing Windows SDK**
**Error**: `error: Microsoft Visual C++ 14.0 is required`
**Solution**:
1. Install **Visual Studio Build Tools** or **Visual Studio Community**
2. Include **Windows 10/11 SDK** and **MSVC v143 compiler**
3. Download from: https://visualstudio.microsoft.com/downloads/

### **Issue 2: Rust Toolchain Issues**
**Error**: `error: toolchain 'stable-x86_64-pc-windows-msvc' is not installed`
**Solution**:
```cmd
rustup toolchain install stable-x86_64-pc-windows-msvc
rustup default stable-x86_64-pc-windows-msvc
```

### **Issue 3: Windows Crate Compilation**
**Error**: `error: failed to run custom build command for 'windows v0.58.0'`
**Solution**:
```cmd
cargo clean
cargo update
cargo build --release
```

### **Issue 4: COM Interface Errors**
**Error**: `error: cannot find type 'CLSID' in this scope`
**Solution**: Update Cargo.toml with correct Windows features:
```toml
[target.'cfg(windows)'.dependencies]
windows = { version = "0.58", features = [
    "Win32_System_Com",
    "Win32_System_Ole",
    "Win32_Foundation",
    "Win32_System_Variant",
    "Win32_System_Com_StructuredStorage",
] }
```

### **Issue 5: Missing QuickBooks SDK**
**Error**: `error: failed to run custom build command for 'qbfc'`
**Solution**:
- This service uses **COM interface** directly (no separate SDK needed)
- QuickBooks Desktop must be installed on the system
- If you see QBFC errors, remove any qbfc dependencies

## 🔧 **Step-by-Step Windows Build Process**

### **Step 1: Install Prerequisites**
1. **Install Rust**:
   ```cmd
   curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh
   ```

2. **Install Visual Studio Build Tools**:
   - Download from: https://visualstudio.microsoft.com/downloads/
   - Include: Windows 10/11 SDK, MSVC v143 compiler

3. **Install QuickBooks Desktop Enterprise**:
   - Must be installed on the same machine as the service
   - Service uses COM interface to connect

### **Step 2: Configure Rust for Windows**
```cmd
rustup toolchain install stable-x86_64-pc-windows-msvc
rustup default stable-x86_64-pc-windows-msvc
```

### **Step 3: Build the Service**
```cmd
# Navigate to the service directory
cd quickbooks-rust-service

# Clean any previous builds
cargo clean

# Build in release mode
cargo build --release
```

### **Step 4: Test the Build**
```cmd
# Run a quick check
cargo check

# Test run (will show connection errors without QuickBooks)
cargo run --release
```

## 🚨 **Troubleshooting Specific Errors**

### **Error: "windows-sys" build failed**
```cmd
# Update Rust
rustup update

# Clean and rebuild
cargo clean
cargo build --release
```

### **Error: "linking with `link.exe` failed"**
- Install **Visual Studio Build Tools**
- Ensure **Windows SDK** is installed
- Restart command prompt after installation

### **Error: "QBFC not found"**
- This service **doesn't use QBFC** - it uses COM directly
- If you see QBFC errors, you may have conflicting dependencies

### **Error: "Permission denied"**
- Run command prompt as **Administrator**
- Ensure antivirus isn't blocking the build

## Runtime Errors

### "Invalid GUID string" Error

**Error**: `thread 'main' panicked at windows-core-0.58.0\src\guid.rs:125:9: Invalid GUID string`

**Cause**: The service couldn't create a valid GUID for the QuickBooks COM interface.

**Solution**: This has been fixed in the latest version. The service now:
- Uses `CLSIDFromProgID()` to properly convert QuickBooks ProgIDs to CLSIDs
- Tries multiple QuickBooks SDK versions (QBFC16, QBFC15, QBFC14, QBFC13)
- Provides better error messages

**If you still see this error**:
1. Make sure you're using the latest version of the service
2. Ensure QuickBooks Desktop is installed and running
3. Try running QuickBooks as Administrator
4. Check that your QuickBooks installation includes the SDK components

## 📋 **Verified Working Configuration**

**Operating System**: Windows 10/11
**Rust Version**: 1.70+ (stable-x86_64-pc-windows-msvc)
**Visual Studio**: Build Tools 2022 with Windows 11 SDK
**QuickBooks**: Desktop Enterprise 2020+

## 🎯 **Quick Fix Commands**

If you're getting build errors, try these commands in order:

```cmd
# 1. Update everything
rustup update
cargo update

# 2. Clean build
cargo clean

# 3. Check for issues
cargo check

# 4. Build release
cargo build --release

# 5. If still failing, reinstall toolchain
rustup toolchain uninstall stable-x86_64-pc-windows-msvc
rustup toolchain install stable-x86_64-pc-windows-msvc
rustup default stable-x86_64-pc-windows-msvc
```

## 📧 **Still Having Issues?**

If you're still getting build errors, please share:
1. **Exact error message**
2. **Windows version**
3. **Rust version** (`rustc --version`)
4. **Visual Studio version**
5. **QuickBooks version**

The service has been tested on Windows 10/11 with QuickBooks Desktop Enterprise 2020-2024.
