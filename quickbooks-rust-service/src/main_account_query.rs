use anyhow::{Result, Context};
use log::{info, error, warn};
use std::env;
use winapi::um::oaidl::IDispatch;

use quickbooks_sheets_sync::config::Config;
use quickbooks_sheets_sync::request_processor::{RequestProcessor2, AccountInfo};

#[derive(Debug, Clone)]
pub struct AccountData {
    pub name: String,
    pub account_number: String,
    pub account_type: String,
    pub balance: f64,
}

impl From<AccountInfo> for AccountData {
    fn from(info: AccountInfo) -> Self {
        Self {
            name: info.name,
            account_number: info.number,
            account_type: info.account_type,
            balance: info.balance,
        }
    }
}

/// Service that handles QuickBooks integration using configuration
pub struct QuickBooksService {
    config: Config,
    processor: Option<RequestProcessor2>,
    session: Option<*mut IDispatch>,
}

impl QuickBooksService {
    pub fn new(config: Config) -> Self {
        Self {
            config,
            processor: None,
            session: None,
        }
    }

    /// Initialize COM and connect to QuickBooks
    pub fn connect(&mut self) -> Result<()> {
        info!("Initializing COM system...");
        unsafe {
            // Replace windows::Win32::System::Com::CoInitialize(None)
            let hr = winapi::um::combaseapi::CoInitializeEx(std::ptr::null_mut(), winapi::um::objbase::COINIT_APARTMENTTHREADED);
            if hr < 0 {
                return Err(anyhow::anyhow!("Failed to initialize COM system: HRESULT=0x{:08X}", hr));
            }
        }

        info!("Creating QuickBooks request processor...");
        let processor = RequestProcessor2::new()
            .context("Failed to create QuickBooks request processor")?;

        // Get application info from config
        let default_id = "QuickBooks-Sheets-Sync".to_string();
        let app_id = self.config.quickbooks.application_id
            .as_ref()
            .unwrap_or(&default_id);
        let default_name = "QuickBooks Sheets Sync".to_string();
        let app_name = self.config.quickbooks.application_name
            .as_ref()
            .unwrap_or(&default_name);

        info!("Opening connection to QuickBooks with app: {}", app_name);
        processor.open_connection(app_id, app_name)?;

        self.processor = Some(processor);
        Ok(())
    }

    /// Begin a session with QuickBooks
    pub fn begin_session(&mut self) -> Result<()> {
        let processor = self.processor.as_ref()
            .ok_or_else(|| anyhow::anyhow!("Not connected to QuickBooks. Call connect() first."))?;

        info!("Beginning QuickBooks session...");
        
        // Use the company file from config, or empty string for currently open file
        let company_file = match self.config.quickbooks.company_file.as_str() {
            "AUTO" => "",
            path => path,
        };

        let session = processor.begin_session(company_file, quickbooks_sheets_sync::FileMode::DoNotCare)?;
        self.session = Some(session);
        
        info!("Successfully started QuickBooks session");
        Ok(())
    }

    /// Get account data for the configured account number
    pub fn get_account_data(&self) -> Result<AccountData> {
        let processor = self.processor.as_ref()
            .ok_or_else(|| anyhow::anyhow!("Not connected to QuickBooks"))?;
        
        let session = self.session.ok_or_else(|| anyhow::anyhow!("No active session. Call begin_session() first."))?;

        let account_number = &self.config.quickbooks.account_number;
        info!("Querying account data for account number: {}", account_number);

        match processor.query_account_by_number(session, account_number)? {
            Some(account_info) => {
                info!("Successfully retrieved account: {} ({})", account_info.name, account_info.number);
                Ok(account_info.into())
            }
            None => {
                warn!("Account {} not found in QuickBooks", account_number);
                Err(anyhow::anyhow!("Account {} not found", account_number))
            }
        }
    }

    /// Get the current company file name
    pub fn get_company_file_name(&self) -> Result<String> {
        let processor = self.processor.as_ref()
            .ok_or_else(|| anyhow::anyhow!("Not connected to QuickBooks"))?;
        
        processor.get_current_company_file_name()
            .context("Failed to get company file name")
    }

    /// End the session and disconnect
    pub fn disconnect(&mut self) -> Result<()> {
        if let (Some(processor), Some(session)) = (self.processor.as_ref(), self.session) {
            info!("Ending QuickBooks session...");
            processor.end_session()?;
            
            info!("Closing QuickBooks connection...");
            processor.close_connection()?;
            
            self.session = None;
        }

        if self.processor.is_some() {
            self.processor = None;
            info!("Uninitializing COM system...");
            unsafe {
                winapi::um::combaseapi::CoUninitialize();
            }
        }

        Ok(())
    }
}

impl Drop for QuickBooksService {
    fn drop(&mut self) {
        if let Err(e) = self.disconnect() {
            error!("Error during QuickBooks service cleanup: {}", e);
        }
    }
}

fn print_instructions() {
    println!("QuickBooks Account Query Service");
    println!("===============================");
    println!();
    println!("This service reads configuration from config/config.toml and queries");
    println!("the specified account from QuickBooks Desktop.");
    println!();
    println!("Prerequisites:");
    println!("   1. QuickBooks Desktop must be installed and running");
    println!("   2. A company file must be open in QuickBooks");
    println!("   3. The account number in config.toml must exist in QuickBooks");
    println!();
}

fn main() -> Result<()> {
    // Initialize logging
    let args: Vec<String> = env::args().collect();
    let verbose = args.contains(&"--verbose".to_string()) || args.contains(&"-v".to_string());

    if verbose {
        env_logger::builder()
            .filter_level(log::LevelFilter::Debug)
            .init();
    } else {
        env_logger::builder()
            .filter_level(log::LevelFilter::Info)
            .init();
    }

    print_instructions();

    // Load configuration
    info!("Loading configuration from config/config.toml...");
    let config = Config::load_from_file("config/config.toml")
        .context("Failed to load configuration file")?;

    info!("Configuration loaded successfully");
    info!("Target account: {} - {}", 
          config.quickbooks.account_number, 
          config.quickbooks.account_name);

    // Create and initialize the service
    let mut service = QuickBooksService::new(config);

    // Connect to QuickBooks
    info!("Connecting to QuickBooks Desktop...");
    service.connect()
        .context("Failed to connect to QuickBooks")?;

    // Begin session
    info!("Starting QuickBooks session...");
    service.begin_session()
        .context("Failed to begin QuickBooks session")?;

    // Get company file info
    match service.get_company_file_name() {
        Ok(company_file) => info!("Connected to company file: {}", company_file),
        Err(e) => warn!("Could not get company file name: {}", e),
    }

    // Query the configured account
    info!("Querying account data...");
    match service.get_account_data() {
        Ok(account_data) => {
            println!();
            println!("SUCCESS! Account data retrieved:");
            println!("================================");
            println!("Account Number: {}", account_data.account_number);
            println!("Account Name:   {}", account_data.name);
            println!("Account Type:   {}", account_data.account_type);
            println!("Balance:        ${:.2}", account_data.balance);
            println!();
        }
        Err(e) => {
            error!("Failed to retrieve account data: {}", e);
            println!();
            println!("Account query failed. Please check:");
            println!("1. The account number '{}' exists in your QuickBooks company file", service.config.quickbooks.account_number);
            println!("2. You have permission to access account data");
            println!("3. QuickBooks is not busy with other operations");
            return Err(e);
        }
    }

    // Clean up
    info!("Disconnecting from QuickBooks...");
    service.disconnect()?;

    println!("QuickBooks account query completed successfully!");
    Ok(())
}
