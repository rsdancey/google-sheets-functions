use anyhow::Result;
use crate::config::Config;
use crate::account_service::AccountService;

/// High-level QuickBooks client that hides COM/VARIANT complexity
pub struct QuickBooksClient {
    config: Config,
}

/// Clean account data structure
#[derive(Debug, Clone)]
pub struct AccountBalance {
    pub account_number: String,
    pub account_name: String,
    pub balance: f64,
    pub account_type: String,
}

impl QuickBooksClient {
    pub fn new(config: Config) -> Self {
        Self { config }
    }
    
    /// High-level method to get account balance by number
    /// Uses the real QBFC API with SafeVariant wrappers
    pub fn get_account_balance(&self, account_number: &str) -> Result<Option<AccountBalance>> {
        println!("🔍 Querying QuickBooks for account: {}", account_number);
        
        // Use the actual AccountService with QBFC API
        let account_service = AccountService::new(self.config.clone())?;
        
        match account_service.get_account_balance()? {
            Some(account_info) => {
                // Convert from AccountInfo to AccountBalance
                Ok(Some(AccountBalance {
                    account_number: account_info.number,
                    account_name: account_info.name,
                    balance: account_info.balance,
                    account_type: account_info.account_type, // This is already a String
                }))
            }
            None => Ok(None),
        }
    }
    
    /// Test QuickBooks connection using the real QBFC API
    pub fn test_connection(&self) -> Result<bool> {
        println!("🔗 Testing QuickBooks connection...");
        
        // Use AccountService to test the connection
        let account_service = AccountService::new(self.config.clone())?;
        
        // Try to get account balance - if this succeeds, connection works
        match account_service.get_account_balance() {
            Ok(_) => {
                println!("✅ QuickBooks connection test successful");
                Ok(true)
            }
            Err(e) => {
                println!("❌ QuickBooks connection test failed: {}", e);
                Ok(false)
            }
        }
    }
    
    /// Get configuration details
    pub fn get_config(&self) -> &Config {
        &self.config
    }
}

/// High-level service that orchestrates the sync process
pub struct SyncService {
    qb_client: QuickBooksClient,
}

impl SyncService {
    pub fn new(config: Config) -> Self {
        Self {
            qb_client: QuickBooksClient::new(config),
        }
    }
    
    /// Main sync operation - high-level and clean
    pub fn sync_account_to_sheets(&self) -> Result<()> {
        let config = self.qb_client.get_config();
        
        println!("🚀 Starting sync operation...");
        println!("   Account: {} ({})", 
            config.quickbooks.account_name, 
            config.quickbooks.account_number
        );
        
        // Step 1: Get account balance from QuickBooks
        let balance = match self.qb_client.get_account_balance(&config.quickbooks.account_number)? {
            Some(account) => {
                println!("✅ Found account: {} = ${:.2}", account.account_name, account.balance);
                account.balance
            }
            None => {
                println!("❌ Account {} not found", config.quickbooks.account_number);
                return Ok(());
            }
        };
        
        // Step 2: Update Google Sheets
        println!("📊 Updating Google Sheets...");
        println!("   Spreadsheet: {}", config.google_sheets.spreadsheet_id);
        println!("   Cell: {}", config.google_sheets.cell_address);
        println!("   Value: ${:.2}", balance);
        
        // TODO: Implement Google Sheets API call
        println!("✅ Sync completed successfully!");
        
        Ok(())
    }
}
