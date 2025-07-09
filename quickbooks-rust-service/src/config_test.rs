use anyhow::{Result, Context};
use crate::config::Config;

/// Simple demonstration of reading config and connecting to QuickBooks
pub fn test_account_config() -> Result<()> {
    println!("=== QuickBooks Account Configuration Test ===");
    
    // Load configuration from config.toml
    let config = Config::load()
        .context("Failed to load configuration from config.toml")?;
    
    println!("✅ Configuration loaded successfully!");
    println!();
    
    // Display the account configuration that was read from config.toml
    println!("Account Configuration:");
    println!("  Account Number: {}", config.quickbooks.account_number);
    println!("  Account Name: {}", config.quickbooks.account_name);
    println!("  Company File: {}", config.quickbooks.company_file);
    println!("  Application Name: {}", 
        config.quickbooks.application_name.as_deref().unwrap_or("Not specified"));
    println!();
    
    // Display Google Sheets configuration
    println!("Google Sheets Configuration:");
    println!("  Spreadsheet ID: {}", config.google_sheets.spreadsheet_id);
    println!("  Sheet Name: {}", config.google_sheets.sheet_name.as_deref().unwrap_or("Not specified"));
    println!("  Cell Address: {}", config.google_sheets.cell_address);
    println!();
    
    println!("🎯 Target: Sync account '{}' ({}) balance to Google Sheets cell {}", 
        config.quickbooks.account_name,
        config.quickbooks.account_number,
        config.google_sheets.cell_address
    );
    
    println!();
    println!("✅ Configuration test completed successfully!");
    println!("Next step: Implement QBFC API account query using this configuration.");
    
    Ok(())
}
