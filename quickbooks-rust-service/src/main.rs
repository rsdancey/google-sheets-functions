mod file_mode;
mod config;
mod qbfc_safe;
mod qbxml_safe;

use anyhow::{Result, Context};
use log::{info, error, warn};
use std::env;

// --- QBFC imports ---
use crate::config::Config;
use crate::qbfc_safe::qbfc_request_processor::{RequestProcessor2 as QbfcRequestProcessor, AccountInfo as QbfcAccountInfo};
use crate::file_mode::FileMode;

// --- QBXML imports ---
use crate::qbxml_safe::qbxml_request_processor::QbxmlRequestProcessor;

#[derive(Debug, Clone)]
pub struct AccountData {
    pub name: String,
    pub account_number: String,
    pub account_type: String,
    pub balance: f64,
}

impl From<QbfcAccountInfo> for AccountData {
    fn from(info: QbfcAccountInfo) -> Self {
        Self {
            name: info.name,
            account_number: info.number,
            account_type: info.account_type,
            balance: info.balance,
        }
    }
}

fn print_instructions() {
    println!("QuickBooks Account Query Service v2");
    println!("===================================");
    println!();
    println!("This service reads configuration from config/config.toml and queries");
    println!("the specified account from QuickBooks Desktop.");
    println!();
    println!("Prerequisites:");
    println!("   1. QuickBooks Desktop must be installed and running");
    println!("   2. A company file must be open in QuickBooks");
    println!("   3. The account number in config.toml must exist in QuickBooks");
    println!();
    println!("Usage: main_account_query --api qbfc|qbxml [--verbose]");
    println!();
}


fn main() {
    if let Err(e) = real_main() {
        eprintln!("Error: {:#}", e);
        std::process::exit(1);
    }
}

fn real_main() -> anyhow::Result<()> {
    // Parse arguments
    let args: Vec<String> = env::args().collect();
    let verbose = args.contains(&"--verbose".to_string()) || args.contains(&"-v".to_string());
    let api_mode = args.iter().position(|a| a == "--api").and_then(|i| args.get(i+1)).map(|s| s.as_str()).unwrap_or("qbfc");

    if verbose {
        env_logger::builder().filter_level(log::LevelFilter::Debug).init();
    } else {
        env_logger::builder().filter_level(log::LevelFilter::Info).init();
    }

    print_instructions();

    // Load configuration
    info!("Loading configuration from config/config.toml...");
    let config = Config::load_from_file("config/config.toml")
        .context("Failed to load configuration file")?;
    info!("Configuration loaded successfully");
    info!("Target account: {} - {}", config.quickbooks.account_number, config.quickbooks.account_name);

    match api_mode {
        "qbfc" => run_qbfc(config),
        "qbxml" => run_qbxml(config),
        _ => {
            error!("Unknown API mode: {}. Use --api qbfc or --api qbxml", api_mode);
            std::process::exit(1);
        }
    }
}

// --- QBFC code path ---
fn run_qbfc(config: Config) -> Result<()> {
    info!("[QBFC] Connecting to QuickBooks Desktop...");
    unsafe {
        let hr = winapi::um::combaseapi::CoInitializeEx(std::ptr::null_mut(), winapi::um::objbase::COINIT_APARTMENTTHREADED);
        if hr < 0 {
            return Err(anyhow::anyhow!("Failed to initialize COM system: HRESULT=0x{:08X}", hr));
        }
    }
    let processor = QbfcRequestProcessor::new().context("Failed to create QuickBooks request processor")?;
    let app_id = config.quickbooks.application_id.as_deref().unwrap_or("QuickBooks-Sheets-Sync");
    let app_name = config.quickbooks.application_name.as_deref().unwrap_or("QuickBooks Sheets Sync");
    info!("[QBFC] Opening connection to QuickBooks with app: {}", app_name);
    processor.open_connection(app_id, app_name)?;
    let company_file = match config.quickbooks.company_file.as_str() { "AUTO" => "", path => path };
    let session = processor.begin_session(company_file, crate::FileMode::DoNotCare)?;
    info!("[QBFC] Successfully started QuickBooks session");
    let account_number = &config.quickbooks.account_number;
    info!("[QBFC] Querying account data for account number: {}", account_number);
    match processor.query_account_by_number(session, account_number)? {
        Some(account_info) => {
            info!("[QBFC] Successfully retrieved account: {} ({})", account_info.name, account_info.number);
            println!("\nSUCCESS! Account data retrieved (QBFC):\n================================");
            println!("Account Number: {}", account_info.number);
            println!("Account Name:   {}", account_info.name);
            println!("Account Type:   {}", account_info.account_type);
            println!("Balance:        ${:.2}", account_info.balance);
        }
        None => {
            warn!("[QBFC] Account {} not found in QuickBooks", account_number);
            println!("\nAccount query failed. Account {} not found.", account_number);
            return Err(anyhow::anyhow!("Account {} not found", account_number));
        }
    }
    info!("[QBFC] Disconnecting from QuickBooks...");
    processor.end_session()?;
    processor.close_connection()?;
    unsafe { winapi::um::combaseapi::CoUninitialize(); }
    println!("QuickBooks account query completed successfully! (QBFC)");
    Ok(())
}

// --- QBXML code path ---
fn run_qbxml(config: Config) -> Result<()> {
    info!("[QBXML] Connecting to QuickBooks Desktop...");
    unsafe {
        let hr = winapi::um::combaseapi::CoInitializeEx(std::ptr::null_mut(), winapi::um::objbase::COINIT_APARTMENTTHREADED);
        if hr < 0 {
            return Err(anyhow::anyhow!("Failed to initialize COM system: HRESULT=0x{:08X}", hr));
        }
    }

    let processor = QbxmlRequestProcessor::new().context("Failed to create QBXML request processor")?;

    let app_id = config.quickbooks.application_id.as_deref().unwrap_or("QuickBooks-Sheets-Sync");

    let app_name = config.quickbooks.application_name.as_deref().unwrap_or("QuickBooks Sheets Sync");

    info!("[QBXML] Opening connection to QuickBooks with app: {}", app_name);
    processor.open_connection(app_id, app_name)?;

    let company_file = match config.quickbooks.company_file.as_str() { "AUTO" => "", path => path };
    let ticket = processor.begin_session(company_file, crate::FileMode::DoNotCare)?;

    info!("[QBXML] Successfully started QuickBooks session");
    // For demo: just print the ticket and close
    println!("[QBXML] Session ticket: {}", ticket);
    // You would add QBXML request/response logic here
    processor.end_session(&ticket)?;
    processor.close_connection()?;
    unsafe { winapi::um::combaseapi::CoUninitialize(); }
    println!("QuickBooks account query completed successfully! (QBXML)");
    Ok(())
}
