use anyhow::{Context, Result};
use log::{error, info, warn};
use serde::{Deserialize, Serialize};
use std::env;
use std::time::Duration;
use tokio::time;
use tokio_cron_scheduler::{Job, JobScheduler};

mod config;
mod quickbooks;
mod sheets;

use config::Config;
use quickbooks::QuickBooksClient;
use sheets::GoogleSheetsClient;

#[derive(Debug, Serialize, Deserialize)]
pub struct AccountData {
    pub account_number: String,
    pub account_name: String,
    pub balance: f64,
    pub timestamp: String,
}

#[tokio::main]
async fn main() -> Result<()> {
    // Parse command line arguments first to determine log level
    let args: Vec<String> = env::args().collect();
    let verbose = args.contains(&"--verbose".to_string()) || args.contains(&"-v".to_string());
    let help = args.contains(&"--help".to_string()) || args.contains(&"-h".to_string());

    // Initialize logging once at the start
    let log_result = if verbose {
        env_logger::builder()
            .filter_level(log::LevelFilter::Debug)
            .try_init()
    } else {
        env_logger::try_init()
    };
    
    // Log initialization result (only if it succeeded)
    if log_result.is_ok() {
        // Logger was successfully initialized
    } else {
        // Logger was already initialized (this is fine)
    }

    // Parse other command line arguments
    if help {
        println!("QuickBooks to Google Sheets Sync Service");
        println!("Usage: qb_sync.exe [OPTIONS]");
        println!();
        println!("OPTIONS:");
        println!("  --test-connection    Test connections and exit");
        println!("  --test-methods       Test VARIANT conversions and response parsing");
        println!("  --test-account       Test account balance retrieval");
        println!("  --skip-sheets-test   Skip Google Sheets connection test");
        println!("  --show-config        Show current configuration and exit");
        println!("  --verbose, -v        Enable verbose logging");
        println!("  --help, -h           Show this help message");
        println!();
        println!("Examples:");
        println!("  qb_sync.exe                    # Run the service normally");
        println!("  qb_sync.exe --test-connection  # Test connections only");
        println!("  qb_sync.exe --test-methods     # Test internal methods");
        println!("  qb_sync.exe --verbose          # Run with detailed logging");
        return Ok(());
    }

    // Parse other command line arguments
    let skip_sheets_test = args.contains(&"--skip-sheets-test".to_string());
    let show_config = args.contains(&"--show-config".to_string());
    let test_account = args.contains(&"--test-account".to_string());

    // Load configuration
    let config = Config::load()
        .context("Failed to load configuration")?;

    // Show configuration if requested (without sensitive data)
    if show_config {
        info!("Configuration loaded:");
        info!("  QuickBooks company_file: {}", config.quickbooks.company_file);
        info!("  QuickBooks account_number: {}", config.quickbooks.account_number);
        info!("  QuickBooks application_name: {:?}", config.quickbooks.application_name);
        info!("  QuickBooks application_id: {:?}", config.quickbooks.application_id);
        info!("  Google Sheets webapp_url: {}", config.google_sheets.webapp_url);
        info!("  Google Sheets spreadsheet_id: {}", config.google_sheets.spreadsheet_id);
        info!("  Google Sheets sheet_name: {:?}", config.google_sheets.sheet_name);
        info!("  Google Sheets cell_address: {}", config.google_sheets.cell_address);
        info!("  API key configured: {}", if config.google_sheets.api_key.is_empty() { "NO" } else { "YES" });
        return Ok(());
    }

    // Initialize clients
    let qb_client = QuickBooksClient::new(&config.quickbooks)
        .context("Failed to initialize QuickBooks client")?;
    
    // If testing account balance, run the account balance test before other connections
    if test_account {
        info!("Testing account balance retrieval...");
        
        match qb_client.get_account_balance(&config.quickbooks.account_number).await {
            Ok(account_data) => {
                info!("✅ Account balance test successful!");
                info!("Account: {} ({})", account_data.account_number, account_data.account_name);
                info!("Balance: ${:.2}", account_data.balance);
                info!("Timestamp: {}", account_data.timestamp);
            }
            Err(e) => {
                error!("❌ Account balance test failed: {}", e);
                return Err(e);
            }
        }
        
        info!("Account balance test completed");
        return Ok(());
    }
    
    let sheets_client = GoogleSheetsClient::new(&config.google_sheets)
        .context("Failed to initialize Google Sheets client")?;

    // Test initial connection
    test_quickbooks_connection(&qb_client, &config).await?;

    if !skip_sheets_test {
        info!("Testing Google Sheets connection...");
        test_sheets_connection(&sheets_client, &config).await?;
    } else {
        warn!("Skipping Google Sheets connection test");
    }

    // Set up scheduler
    let scheduler = JobScheduler::new().await?;

    // Clone data for the job closure
    let qb_client_clone = qb_client.clone();
    let sheets_client_clone = sheets_client.clone();
    let config_clone = config.clone();

    // Create the sync job
    let sync_job = Job::new_async(config.schedule.cron_expression.as_str(), move |_uuid, _l| {
        let qb_client = qb_client_clone.clone();
        let sheets_client = sheets_client_clone.clone();
        let config = config_clone.clone();
        
        Box::pin(async move {
            if let Err(e) = sync_account_data(&qb_client, &sheets_client, &config).await {
                error!("Sync job failed: {}", e);
            }
        })
    })?;

    scheduler.add(sync_job).await?;
    scheduler.start().await?;

    info!("Scheduler started with cron expression: {}", config.schedule.cron_expression);
    
    // Run initial sync
    sync_account_data(&qb_client, &sheets_client, &config).await?;

    info!("✅ Initial sync completed successfully!");
    info!("🕐 Service now running on schedule: {}", config.schedule.cron_expression);
    info!("📊 Will sync account {} to Google Sheets every time the schedule triggers", config.quickbooks.account_number);
    info!("🛑 Press Ctrl+C to stop the service");
    
    // Keep the service running with periodic status updates
    let mut counter = 0;
    loop {
        time::sleep(Duration::from_secs(60)).await;
        counter += 1;
        
        // Show status every 5 minutes
        if counter % 5 == 0 {
            info!("🔄 Service running normally ({} minutes elapsed)", counter);
        }
    }
}

async fn sync_account_data(
    qb_client: &QuickBooksClient,
    sheets_client: &GoogleSheetsClient,
    config: &Config,
) -> Result<()> {
    let start_time = std::time::Instant::now();

    // Get account balance from QuickBooks
    let account_data = qb_client
        .get_account_balance(&config.quickbooks.account_number)
        .await
        .context("Failed to get account balance from QuickBooks")?;

    info!("✅ Retrieved account data: {} = ${:.2}", 
          account_data.account_name, 
          account_data.balance);

    // Send data to Google Sheets
    info!("📝 Updating Google Sheets...");
    sheets_client
        .update_account_data(&account_data, config)
        .await
        .context("Failed to update Google Sheets")?;

    let elapsed = start_time.elapsed();
    info!("✅ Sync completed successfully in {:.2}s", elapsed.as_secs_f64());
    info!("📊 Account {} ({}) = ${:.2} updated in cell {}", 
          account_data.account_number,
          account_data.account_name,
          account_data.balance,
          config.google_sheets.cell_address);

    Ok(())
}

async fn test_quickbooks_connection(
    qb_client: &QuickBooksClient,
    config: &Config,
) -> Result<()> {
    // First, test basic QuickBooks availability
    match qb_client.test_connection().await {
        Ok(()) => {
            // Connection test passed
        }
        Err(e) => {
            error!("QuickBooks application connection failed: {}", e);
            warn!("Make sure QuickBooks Desktop Enterprise is running");
            return Err(e);
        }
    }
    
    // Next, prepare QuickBooks connection strategy
    match qb_client.connect_to_quickbooks().await {
        Ok(()) => {
            // Connection strategy prepared
        }
        Err(e) => {
            error!("QuickBooks connection configuration failed: {}", e);
            warn!("Connection method: {}", 
                  if config.quickbooks.company_file.is_empty() {
                      "Interactive file selection"
                  } else if config.quickbooks.company_file == "AUTO" {
                      "Currently open company file"
                  } else {
                      "Specific file path"
                  });
            
            if config.quickbooks.company_file != "AUTO" && !config.quickbooks.company_file.is_empty() {
                warn!("Company file: {}", config.quickbooks.company_file);
                warn!("Make sure the file exists and is accessible");
            } else {
                warn!("Make sure QuickBooks is running with a company file open");
            }
            
            return Err(e);
        }
    }
    
    Ok(())
}

async fn test_sheets_connection(
    sheets_client: &GoogleSheetsClient,
    config: &Config,
) -> Result<()> {
    match sheets_client.test_connection().await {
        Ok(()) => {
            info!("Google Sheets connection successful");
            Ok(())
        }
        Err(e) => {
            error!("Google Sheets connection failed: {}", e);
            warn!("Check your Google Apps Script URL and API key");
            warn!("URL: {}", config.google_sheets.webapp_url);
            Err(e)
        }
    }
}
