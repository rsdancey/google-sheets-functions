use anyhow::{Context, Result};
use log::{error, info};
use std::env;

mod config;
mod quickbooks;
mod sheets;

use config::Config;
use quickbooks::QuickBooksClient;

#[derive(Debug, Clone)]
pub struct AccountData {
    pub name: String,
    pub account_number: String,
    pub account_type: String,
    pub balance: f64,
    pub description: Option<String>,
}

fn print_instructions() {
    println!("🔧 QuickBooks Desktop Integration Service");
    println!("==========================================");
    println!();
    println!("📋 Prerequisites before running this service:");
    println!("   1. QuickBooks Desktop must be installed");
    println!("   2. QuickBooks Desktop must be RUNNING");
    println!("   3. A company file must be OPEN in QuickBooks");
    println!("   4. QuickBooks should be in a ready state (not processing anything)");
    println!();
    println!("⚠️  If you see 'Session manager functionality test failed', please:");
    println!("   • Open QuickBooks Desktop");
    println!("   • Open a company file");
    println!("   • Wait for QuickBooks to finish loading completely");
    println!("   • Then run this service again");
    println!();
    println!("Starting QuickBooks integration test...");
    println!("========================================");
}

#[tokio::main]
async fn main() -> Result<()> {
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

    // Initialize QuickBooks client
    let qb_client = QuickBooksClient::new(&config.quickbooks)
        .context("Failed to initialize QuickBooks client")?;

    // Test 1: Check if QuickBooks SDK is available
    info!("📋 Step 1: Testing QuickBooks SDK availability...");
    qb_client.test_connection().await
        .context("QuickBooks SDK test failed")?;

    // Test 2: Attempt to register with QuickBooks
    info!("📋 Step 2: Attempting to connect and register with QuickBooks...");
    match qb_client.register_with_quickbooks().await {
        Ok(()) => {
            info!("🎉 SUCCESS! Application connected and registered with QuickBooks.");
            
            // Test 3: Test XML processing (account data retrieval)
            info!("🗒️ Step 3: Testing XML processing (account data retrieval)...");
            match qb_client.test_account_data_retrieval().await {
                Ok(xml_response) => {
                    info!("✅ Account data retrieval successful!");
                    info!("📄 XML Response received ({} characters)", xml_response.len());
                    
                    // Show first 500 characters of response for debugging
                    if xml_response.len() > 500 {
                        info!("📋 Response preview: {}...", &xml_response[..500]);
                    } else {
                        info!("📋 Full response: {}", xml_response);
                    }
                }
                Err(e) => {
                    error!("❌ Account data retrieval failed: {}", e);
                    info!("💡 This may be expected if:");
                    info!("   1. Company file has no accounts");
                    info!("   2. Permissions are restricted");
                    info!("   3. QBXML version incompatibility");
                }
            }
            
            info!("📝 Next steps:");
            info!("   1. QuickBooks should have shown an authorization dialog");
            info!("   2. The application is now registered and can connect");
            info!("   3. Account data retrieval has been tested");
        }
        Err(e) => {
            error!("❌ Connection and registration failed: {}", e);
            info!("💡 Troubleshooting tips:");
            info!("   1. Make sure QuickBooks Desktop is running");
            info!("   2. Make sure a company file is open");
            info!("   3. QuickBooks may show an authorization dialog - approve it");
            info!("   4. Try running QuickBooks as Administrator");
            return Err(e);
        }
    };

    Ok(())
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
    println!("    Official Intuit C++ SDK sample patterns (sdktest.cpp)");
}
