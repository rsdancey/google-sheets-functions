use anyhow::{Context, Result};
use log::{error, info};
use std::env;

mod config;
mod quickbooks_clean as quickbooks;

use config::Config;
use quickbooks::QuickBooksClient;

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

    info!("🚀 QuickBooks Sheets Sync - Clean Registration Test");
    
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
    info!("📋 Step 2: Attempting to register with QuickBooks...");
    match qb_client.register_with_quickbooks().await {
        Ok(()) => {
            info!("🎉 SUCCESS! Application has been registered with QuickBooks.");
            info!("📝 Next steps:");
            info!("   1. QuickBooks should have shown an authorization dialog");
            info!("   2. The application is now registered and can connect");
            info!("   3. You can now implement account data retrieval");
        }
        Err(e) => {
            error!("❌ Registration failed: {}", e);
            info!("💡 Troubleshooting tips:");
            info!("   1. Make sure QuickBooks Desktop is running");
            info!("   2. Make sure a company file is open");
            info!("   3. QuickBooks may show an authorization dialog - approve it");
            info!("   4. Try running QuickBooks as Administrator");
            return Err(e);
        }
    }

    Ok(())
}

fn print_help() {
    println!("QuickBooks Sheets Sync - Clean Registration Test");
    println!();
    println!("This program tests basic QuickBooks Desktop integration and registration.");
    println!();
    println!("USAGE:");
    println!("    cargo run [OPTIONS]");
    println!();
    println!("OPTIONS:");
    println!("    --verbose, -v    Enable verbose logging");
    println!("    --help, -h       Show this help message");
    println!();
    println!("PREREQUISITES:");
    println!("    1. QuickBooks Desktop Enterprise must be installed");
    println!("    2. QuickBooks must be running with a company file open");
    println!("    3. Run this program to test basic connection and registration");
    println!();
    println!("WHAT THIS DOES:");
    println!("    1. Tests if QuickBooks SDK components are available");
    println!("    2. Attempts to register the application with QuickBooks");
    println!("    3. Shows success/failure and next steps");
}
