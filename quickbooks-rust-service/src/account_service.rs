use anyhow::{Context, Result};
use crate::config::Config;
use crate::request_processor::RequestProcessor2;
use crate::AccountInfo;
use std::sync::Arc;

/// Service for querying QuickBooks account data based on configuration
pub struct AccountService {
    config: Config,
    processor: Arc<RequestProcessor2>,
}

impl AccountService {
    pub fn new(config: Config) -> Result<Self> {
        let processor = Arc::new(RequestProcessor2::new()
            .context("Failed to create QBFC session manager")?);
        Ok(Self { config, processor })
    }

    /// Connect to QuickBooks and query the configured account
    pub fn get_account_balance(&self) -> Result<Option<AccountInfo>> {
        log::info!("Starting QuickBooks account query service...");
        // Get connection parameters from config
        let app_name = self.config.quickbooks.application_name
            .as_deref()
            .unwrap_or("QuickBooks Sync Service");
        let app_id = self.config.quickbooks.application_id
            .as_deref()
            .unwrap_or(""); // Empty string for testing
        // Open connection to QuickBooks
        self.processor.open_connection(app_id, app_name)
            .context("Failed to open connection to QuickBooks")?;
        log::info!("✅ Opened connection to QuickBooks");
        // Begin session with company file
        let company_file = if self.config.quickbooks.company_file == "AUTO" {
            "" // Empty string for currently open file
        } else {
            &self.config.quickbooks.company_file
        };
        // Use the same processor instance for begin_session and all further calls
        let session_ticket = self.processor.begin_session(company_file, crate::FileMode::DoNotCare)
            .context("Failed to begin QuickBooks session")?;
        log::info!("✅ Began QuickBooks session with ticket: {:?}", session_ticket);
        // Query the configured account
        let account_info = self.processor.query_account_by_number(session_ticket, &self.config.quickbooks.account_number)
            .context("Failed to query account information")?;
        if let Some(ref info) = account_info {
            log::info!("✅ Successfully retrieved account: {} ({}), Balance: ${:.2}", 
                       info.name, info.number, info.balance);
        } else {
            log::warn!("❌ Account {} not found", self.config.quickbooks.account_number);
        }
        // Clean up session
        self.processor.end_session()
            .context("Failed to end QuickBooks session")?;
        self.processor.close_connection()
            .context("Failed to close QuickBooks connection")?;
        log::info!("✅ QuickBooks session closed successfully");
        Ok(account_info)
    }
    
    /// Get the configured account number
    pub fn get_account_number(&self) -> &str {
        &self.config.quickbooks.account_number
    }
    
    /// Get the configured account name
    pub fn get_account_name(&self) -> &str {
        &self.config.quickbooks.account_name
    }
}

// All COM interop is now handled by winapi types via RequestProcessor2
