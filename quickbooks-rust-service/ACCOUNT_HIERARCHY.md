# QuickBooks Account Hierarchy Guide

## Account Structure for Account 9445 (INCOME TAX)

Your QuickBooks account structure is:
```
1000 (Parent Account)
└── BoA Accounts (Sub-account, no number)
    └── 9445 - INCOME TAX (Sub-sub-account)
```

## How QuickBooks SDK Handles Sub-Accounts

### 1. **Account Query Methods**

When querying for account 9445, the QuickBooks SDK needs to:

1. **Search by Account Number**: Use `AccountQuery` with `AccountNumber` filter
2. **Handle Hierarchy**: The SDK automatically handles the parent-child relationships
3. **Return Full Path**: The response includes the full account hierarchy

### 2. **QBXML Structure for Account Query**

```xml
<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <AccountQueryRq requestID="1">
      <AccountNumber>9445</AccountNumber>
      <IncludeRetElement>AccountNumber</IncludeRetElement>
      <IncludeRetElement>Name</IncludeRetElement>
      <IncludeRetElement>FullName</IncludeRetElement>
      <IncludeRetElement>Balance</IncludeRetElement>
      <IncludeRetElement>ParentRef</IncludeRetElement>
    </AccountQueryRq>
  </QBXMLMsgsRq>
</QBXML>
```

### 3. **Expected Response Structure**

```xml
<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="13.0"?>
<QBXML>
  <QBXMLMsgsRs>
    <AccountQueryRs requestID="1" statusCode="0" statusSeverity="Info">
      <AccountRet>
        <ListID>80000001-1234567890</ListID>
        <TimeCreated>2023-01-15T10:30:00-08:00</TimeCreated>
        <TimeModified>2024-12-15T14:22:00-08:00</TimeModified>
        <EditSequence>1234567890</EditSequence>
        <Name>INCOME TAX</Name>
        <FullName>BoA Accounts:INCOME TAX</FullName>
        <IsActive>true</IsActive>
        <ParentRef>
          <ListID>80000002-1234567890</ListID>
          <FullName>BoA Accounts</FullName>
        </ParentRef>
        <Sublevel>2</Sublevel>
        <AccountType>OtherCurrentLiability</AccountType>
        <AccountNumber>9445</AccountNumber>
        <Balance>-18745.32</Balance>
      </AccountRet>
    </AccountQueryRs>
  </QBXMLMsgsRs>
</QBXML>
```

### 4. **Key Points for Implementation**

1. **Account Number is Unique**: Even though 9445 is a sub-account, the account number is unique across all accounts
2. **FullName shows Hierarchy**: The `FullName` field shows the full path: "BoA Accounts:INCOME TAX"
3. **Balance is Direct**: The balance returned is for account 9445 specifically, not aggregated from parents
4. **Account Type Matters**: INCOME TAX is typically an "OtherCurrentLiability" account type
5. **Negative Balance**: Tax liability accounts often have negative balances (money owed)

### 5. **Common Issues and Solutions**

#### **Issue**: Account not found
- **Cause**: Account number doesn't exist or is inactive
- **Solution**: Query all accounts first to verify structure

#### **Issue**: Permission denied
- **Cause**: User doesn't have access to this account
- **Solution**: Ensure QuickBooks user has full access to Chart of Accounts

#### **Issue**: Incorrect balance
- **Cause**: Querying parent account instead of sub-account
- **Solution**: Always query by specific account number (9445), not parent

### 6. **Testing Strategy**

1. **Verify Account Exists**: 
   ```rust
   // Query for the specific account number
   let account_query = create_account_query("9445");
   ```

2. **Check Account Hierarchy**:
   ```rust
   // Verify the parent structure
   assert!(account.full_name.contains("BoA Accounts"));
   assert_eq!(account.account_number, "9445");
   ```

3. **Validate Balance Type**:
   ```rust
   // INCOME TAX accounts typically have negative balances
   // (representing money owed for taxes)
   if account.account_type == "OtherCurrentLiability" {
       // Balance might be negative
   }
   ```

### 7. **Production Implementation Notes**

When implementing the real QuickBooks SDK integration:

1. **Use AccountQuery with AccountNumber filter** - most efficient
2. **Handle the response parsing** - extract balance from the XML response
3. **Account for data types** - QuickBooks balances are typically decimal/currency
4. **Error handling** - account might not exist, be inactive, or have access restrictions
5. **Session management** - properly begin/end QuickBooks sessions

### 8. **Mock Data Enhancement**

The current mock implementation simulates:
- Account 9445 specifically with realistic negative balance for tax liability
- Daily variation to simulate real-world balance changes
- Proper logging of the account hierarchy structure

This helps test the integration before connecting to real QuickBooks data.
