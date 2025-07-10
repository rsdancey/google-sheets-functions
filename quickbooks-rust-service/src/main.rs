mod file_mode;
mod config;
mod qbxml_safe;

use anyhow::{Result, Context};
use log::{info, error, warn};
use std::env;

use crate::config::Config;
use crate::file_mode::FileMode;
use crate::qbxml_safe::qbxml_request_processor::QbxmlRequestProcessor;
mod google_sheets;
use google_sheets::GoogleSheetsClient;

#[derive(Debug, Clone)]
pub struct AccountData {
    pub account_full_name: String,
    pub number: String,
    pub account_type: String,
    pub balance: f64,
}

fn print_instructions() {
    println!("QuickBooks Account Query Service v3");
    println!("===================================");
    println!();
    println!("This service reads configuration from config/config.toml and queries");
    println!("the specified account to retreive its balance from QuickBooks Desktop.");
    println!();
    println!("Prerequisites:");
    println!("   1. QuickBooks Desktop and the QuickBooks SDK v16 (or higher) must be installed and running");
    println!("   2. A company file must be open in QuickBooks");
    println!("   3. The FullName of the account in config.toml must exist in QuickBooks");
    println!();
    println!("Usage: main_account_query [--verbose]");
    println!();
}


#[tokio::main]
async fn main() {
    match real_main().await {
        Err(e) => {
            eprintln!("Error: {:#}", e);
            std::process::exit(1);
        },
        Ok(()) => {
            std::process::exit(0);
        }
    }
}

async fn real_main() -> anyhow::Result<()> {
    // Parse arguments
    let args: Vec<String> = env::args().collect();
    let verbose = args.iter().any(|a| a == "--verbose" || a == "-v");
    let mut account_full_name_arg: Option<&str> = None;
    for arg in &args {
        if let Some(val) = arg.strip_prefix("--account_full_name=") {
            account_full_name_arg = Some(val);
            break;
        }
    }
    let account_full_name_arg = match account_full_name_arg {
        Some(val) if !val.is_empty() => val,
        _ => return Err(anyhow::anyhow!("Missing required --account_full_name argument")),
    };

    if verbose {
        env_logger::builder().filter_level(log::LevelFilter::Debug).init();
    } else {
        env_logger::builder().filter_level(log::LevelFilter::Info).init();
    }

    print_instructions();

    // Load configuration
    info!("Loading configuration from config/config.toml...");
    let mut config = Config::load_from_file("config/config.toml")
        .context("Failed to load configuration file")?;
    info!("Configuration loaded successfully");
    info!("Target account: {}", &account_full_name_arg);

    run_qbxml(config, &account_full_name_arg).await
}

async fn run_qbxml(config: Config, account_full_name: &str) -> Result<()> {
    info!("Connecting to QuickBooks Desktop...");
    unsafe {
        let hr = winapi::um::combaseapi::CoInitializeEx(std::ptr::null_mut(), winapi::um::objbase::COINIT_APARTMENTTHREADED);
        if hr < 0 {
            return Err(anyhow::anyhow!("Failed to initialize COM system: HRESULT=0x{:08X}", hr));
        }
    }

    let processor = QbxmlRequestProcessor::new().context("Failed to create QBXML request processor")?;

    let app_id = config.quickbooks.application_id.as_deref().unwrap_or("QuickBooks-Sheets-Sync");

    let app_name = config.quickbooks.application_name.as_deref().unwrap_or("QuickBooks Sheets Sync");

    info!("Opening connection to QuickBooks with app: {}", app_name);
    processor.open_connection(app_id, app_name)?;

    let company_file = match config.quickbooks.company_file.as_str() { "AUTO" => "", path => path };
    let ticket = processor.begin_session(company_file, crate::FileMode::DoNotCare)?;
    let mut balance = 0.0;
    info!("Session ticket: '{}', length: {}", ticket, ticket.len());
    info!("Successfully started QuickBooks session");
    // Query account by FullName using the config value
    info!("[QBXML] requesting full account xml from QuickBooks");
    match processor.get_account_xml(&ticket) {
        Ok(Some(response_xml)) => {
            info!("[QBXML] response_xml contains valid account data");
            match processor.get_account_balance(&response_xml, account_full_name) {
                Ok(Some(account_balance)) => {
                    info!("[QBXML] Account balance is: {:?}", account_balance);
                    balance = account_balance;
                },
                Ok(None) => {
                    info!("[QBXML] No valid account balance");
                },
                Err(e) => {
                    eprintln!("[QBXML] Error parsing response_xml: {:#}", e);
                }
            }
        },
        Ok(None) => {
            info!("[QBXML] No response_xml received");
        },
        Err(e) => {
            eprintln!("[QBXML] Error querying Quickbooks: {:#}", e);
        }
    }
    info!("Ending session and terminating ticket: {}", ticket);
    processor.end_session(&ticket)?;
    info!("Closing connection to QuickBooks");
    processor.close_connection()?;
    unsafe { winapi::um::combaseapi::CoUninitialize(); }
    println!("run_qbxml complete");

    // Send balance to Google Sheets
    let gs_cfg = &config.google_sheets;
    let gs_client = GoogleSheetsClient::new(
        gs_cfg.webapp_url.clone(),
        gs_cfg.api_key.clone(),
        gs_cfg.spreadsheet_id.clone(),
        gs_cfg.sheet_name.clone(),
        gs_cfg.cell_address.clone(),
    );
    gs_client.send_balance(account_full_name, balance).await?;
    Ok(())
}
