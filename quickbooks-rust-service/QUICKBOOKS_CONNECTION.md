# QuickBooks Company File Connection Methods

## 🔍 **How the Service Finds Your Company File**

The service supports **three different methods** to connect to your QuickBooks company file:

## **Method 1: Specific File Path** 
```toml
company_file = "C:\\Users\\YourUser\\Documents\\QuickBooks\\Your Company.qbw"
```

### How it works:
1. Service connects to QuickBooks Desktop Enterprise
2. Service tells QuickBooks to open the **specific file path** you provided
3. QuickBooks opens that exact file
4. Service extracts account data from that file

### When to use:
- ✅ **Automated/unattended operation** (like a Windows Service)
- ✅ **Single company file** that's always the same
- ✅ **File path never changes**

### Requirements:
- File must exist at the exact path specified
- QuickBooks must have permission to access the file
- No user interaction required

---

## **Method 2: Currently Open File (Recommended)**
```toml
company_file = "AUTO"
```

### How it works:
1. User **manually opens QuickBooks** and opens their company file
2. Service connects to QuickBooks Desktop Enterprise
3. Service says "use whatever company file is currently open"
4. Service extracts account data from the open file

### When to use:
- ✅ **Daily manual operation** (user opens QuickBooks each morning)
- ✅ **Multiple company files** (different files on different days)
- ✅ **Flexible file locations** (files might move or be renamed)

### Requirements:
- QuickBooks must be running
- A company file must be open in QuickBooks
- User must have opened the correct company file

---

## **Method 3: Interactive Selection**
```toml
company_file = ""
```

### How it works:
1. Service connects to QuickBooks Desktop Enterprise
2. QuickBooks shows a **file selection dialog** to the user
3. User chooses which company file to use
4. Service extracts account data from selected file

### When to use:
- ✅ **Multiple company files** available
- ✅ **Manual operation** with user present
- ✅ **Testing/development** scenarios

### Requirements:
- User must be present to select file
- Not suitable for automated operation

---

## **Real-World Examples**

### **Scenario 1: Single Company, Automated**
```toml
# Perfect for running as Windows Service
company_file = "C:\\QuickBooks\\MyCompany2024.qbw"
```

### **Scenario 2: Daily Manual Operation**
```toml
# User opens QuickBooks each morning, service runs hourly
company_file = "AUTO"
```

### **Scenario 3: Multiple Companies**
```toml
# Different company files on different days
company_file = "AUTO"  # User opens the right file each day
```

---

## **QuickBooks Desktop Enterprise Specifics**

### **File Locations**
Common QuickBooks file locations:
- `C:\\Users\\[Username]\\Documents\\QuickBooks\\`
- `C:\\ProgramData\\Intuit\\QuickBooks\\`
- Network drives (for multi-user setups)

### **Multi-User Mode**
If QuickBooks is in multi-user mode:
- Files are typically on a server
- Path might be: `\\\\SERVER\\QuickBooks\\Company.qbw`
- All users need access to the same file

### **Permissions**
QuickBooks may prompt for permission when external applications connect:
- **First time**: User must allow the connection
- **Subsequent times**: Should connect automatically (if saved)

---

## **Troubleshooting Connection Issues**

### **"Company file not found"**
- Check the file path is correct
- Ensure the file exists at that location
- Verify file permissions

### **"QuickBooks connection failed"**
- Make sure QuickBooks Desktop Enterprise is running
- Check if a company file is open (for AUTO mode)
- Verify QuickBooks allows external connections

### **"Permission denied"**
- Run QuickBooks as Administrator
- Check Windows file permissions
- Ensure the service has appropriate user privileges

---

## **Recommended Setup**

For most users, we recommend:
1. **Set** `company_file = "AUTO"`
2. **Manually open** QuickBooks each morning
3. **Open your company file** in QuickBooks
4. **Let the service run** hourly throughout the day

This provides the best balance of **automation** and **flexibility**!
