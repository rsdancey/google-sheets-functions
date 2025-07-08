use anyhow::{Result, Context};
use log::info;
use serde::Deserialize;
use std::env;

use quickbooks_sheets_sync::{QuickBooksClient, QuickBooksConfig};

#[derive(Debug, Deserialize)]
pub struct Config {
    quickbooks: QuickBooksConfig,
}


#[derive(Debug, Clone)]
pub struct AccountData {
    pub name: String,
    pub account_number: String,
    pub account_type: String,
    pub balance: f64,
    pub description: Option<String>,
}

impl Config {
    fn load() -> Result<Self> {
        let config_path = "config.json";
        let config_str = std::fs::read_to_string(config_path)
            .map_err(|e| anyhow::anyhow!(
                "Failed to read config file '{}': {}. \
                Make sure the file exists and you have permission to read it.",
                config_path, e
            ))?;

        serde_json::from_str(&config_str)
            .map_err(|e| anyhow::anyhow!(
                "Failed to parse config file '{}': {}. \
                Check that the file contains valid JSON in the expected format.",
                config_path, e
            ))
    }
}

fn print_instructions() {
    println!("QuickBooks Desktop Integration Service v1.10");
    println!("==========================================");
    println!();
    println!("Prerequisites before running this service:");
    println!("   1. QuickBooks Desktop must be installed");
    println!("   2. QuickBooks Desktop must be RUNNING");
    println!("   3. A company file must be OPEN in QuickBooks");
    println!("   4. QuickBooks should be in a ready state (not processing anything)");
    println!();
    println!("If you see 'Session manager functionality test failed', please:");
    println!("   • Open QuickBooks Desktop");
    println!("   • Open a company file");
    println!("   • Wait for QuickBooks to finish loading completely");
    println!("   • Then run this service again");
    println!();
    println!("Starting QuickBooks integration test...");
    println!("========================================");
}

fn print_help() {
    println!("QuickBooks Sheets Sync - Enhanced Registration and XML Processing Test");
    println!();
    println!("This program tests QuickBooks Desktop integration, registration, and XML processing.");
    println!();
    println!("USAGE:");
    println!("    cargo run [OPTIONS]");
    println!();
    println!("OPTIONS:");
    println!("    --verbose, -v    Enable verbose logging");
    println!("    --help, -h       Show this help message");
    println!();
    println!("PREREQUISITES:");
    println!("    1. QuickBooks Desktop Enterprise must be installed and running");
    println!("    2. A company file must be open in QuickBooks");
    println!("    3. QuickBooks should be in single-user mode (not multi-user)");
    println!("    4. This program may need to run as Administrator");
    println!("    5. Make sure no other QuickBooks integrations are running");
    println!();
    println!("WHAT THIS DOES:");
    println!("    1. Tests if QuickBooks SDK components are available");
    println!("    2. Attempts to register the application with QuickBooks");
    println!("    3. Tests XML processing with account data retrieval");
    println!("    4. Shows success/failure and next steps");
    println!();
    println!("BASED ON:");
    println!("    Official Intuit SDK sample patterns (SDKTestPlus3.vb)");
}

fn main() -> Result<()> {
    // Enable error backtrace if requested
    if env::var("RUST_BACKTRACE").is_err() {
        env::set_var("RUST_BACKTRACE", "1");
    }
    // Initialize logging
    let args: Vec<String> = env::args().collect();
    let verbose = args.contains(&"--verbose".to_string()) || args.contains(&"-v".to_string());

    if verbose {
        env_logger::builder()
            .filter_level(log::LevelFilter::Debug)
            .init();
    } else {
        env_logger::init();
    }

    // Show help if requested
    if args.contains(&"--help".to_string()) || args.contains(&"-h".to_string()) {
        print_help();
        return Ok(());
    }

    print_instructions();

    // Load configuration
    let config = Config::load()
        .context("Failed to load configuration")?;

    // Create QuickBooks client
    let mut qb = QuickBooksClient::new(&config.quickbooks)?;

    // Connect to QuickBooks
    info!("Attempting to connect to QuickBooks...");
    qb.connect("")?;  // Empty string means use currently open company file
    info!("Connected to QuickBooks and started session");

    // Get company file
    let company_file = qb.get_company_file_name()?;
    info!("Current company file: {}", company_file);

    Ok(())
}
