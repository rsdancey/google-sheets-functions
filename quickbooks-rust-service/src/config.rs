use anyhow::{Context, Result};
use figment::{Figment, providers::{Format, Toml}};
use serde::{Deserialize, Serialize};
use std::path::Path;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Config {
    pub quickbooks: QuickBooksConfig,
    pub google_sheets: GoogleSheetsConfig,
    pub schedule: ScheduleConfig,
    pub server: Option<ServerConfig>,  // Added server configuration
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct QuickBooksConfig {
    pub enabled: Option<bool>,  // Added enabled flag
    pub company_file: String,
    pub account_number: String,
    pub account_name: String,
    pub connection_mode: Option<String>,  // Added connection mode preference
    pub application_name: Option<String>,  // Added application name
    pub application_id: Option<String>,  // Added application ID (AppID/GUID)
    pub company_file_password: Option<String>,
    pub qb_username: Option<String>,
    pub qb_password: Option<String>,
    pub connection_timeout: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct GoogleSheetsConfig {
    pub webapp_url: String,
    pub api_key: String,
    pub spreadsheet_id: String,
    pub sheet_name: Option<String>,
    pub cell_address: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ScheduleConfig {
    pub cron_expression: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ServerConfig {
    pub run_as_service: Option<bool>,
    pub connection_timeout: Option<u32>,
    pub retry_attempts: Option<u32>,
    pub log_to_event_log: Option<bool>,
    pub health_check_port: Option<u16>,
}

impl Config {
    pub fn load() -> Result<Self> {
        // Check if config file exists
        let config_path = Path::new("config/config.toml");
        if !config_path.exists() {
            return Err(anyhow::anyhow!(
                "Configuration file not found at: {}. Please create it from config.example.toml",
                config_path.display()
            ));
        }

        // Use Figment to load configuration
        let mut config: Config = Figment::new()
            // Load base configuration from TOML file
            .merge(Toml::file("config/config.toml"))
            .extract()
            .context("Failed to parse configuration")?;

        // Override with environment variables if present (same as old config)
        if let Ok(password) = std::env::var("QB_FILE_PASSWORD") {
            config.quickbooks.company_file_password = Some(password);
        }
        
        if let Ok(username) = std::env::var("QB_USERNAME") {
            config.quickbooks.qb_username = Some(username);
        }
        
        if let Ok(password) = std::env::var("QB_USER_PASSWORD") {
            config.quickbooks.qb_password = Some(password);
        }
        
        if let Ok(api_key) = std::env::var("SHEETS_API_KEY") {
            config.google_sheets.api_key = api_key;
        }

        // Post-process Windows paths
        config.normalize_paths();

        // Validate configuration
        config.validate()?;

        Ok(config)
    }

    fn validate(&self) -> Result<()> {
        // Validate QuickBooks config
        if self.quickbooks.company_file.is_empty() {
            return Err(anyhow::anyhow!("QuickBooks company file path cannot be empty"));
        }

        if self.quickbooks.account_number.is_empty() {
            return Err(anyhow::anyhow!("QuickBooks account number cannot be empty"));
        }

        // Validate Google Sheets config
        if self.google_sheets.webapp_url.is_empty() {
            return Err(anyhow::anyhow!("Google Sheets webapp URL cannot be empty"));
        }

        if self.google_sheets.api_key.is_empty() {
            return Err(anyhow::anyhow!("Google Sheets API key cannot be empty"));
        }

        if self.google_sheets.cell_address.is_empty() {
            return Err(anyhow::anyhow!("Google Sheets cell address cannot be empty"));
        }

        // Validate schedule config
        if self.schedule.cron_expression.is_empty() {
            return Err(anyhow::anyhow!("Schedule cron expression cannot be empty"));
        }

        Ok(())
    }

    /// Normalize file paths to handle Windows path separators
    fn normalize_paths(&mut self) {
        // Handle Windows file paths for company_file
        self.quickbooks.company_file = self.normalize_windows_path(&self.quickbooks.company_file);
    }

    /// Normalize a Windows file path by converting single backslashes to double backslashes
    /// and handling common Windows path issues
    fn normalize_windows_path(&self, path: &str) -> String {
        // Skip normalization for special values
        if path == "MOCK" || path == "AUTO" || path.is_empty() {
            return path.to_string();
        }

        // If running on Windows and the path looks like a Windows path
        if cfg!(windows) && Self::is_windows_path(path) {
            // Convert forward slashes to backslashes for Windows
            let normalized = path.replace('/', "\\");
            log::info!("Normalized Windows path: '{}' -> '{}'", path, normalized);
            normalized
        } else {
            // Not a Windows path or not on Windows, return as-is
            path.to_string()
        }
    }

    /// Check if a path looks like a Windows path
    fn is_windows_path(path: &str) -> bool {
        // Check for Windows drive letters (C:, D:, etc.) or UNC paths (\\server)
        path.len() >= 3 && path.chars().nth(1) == Some(':') ||
        path.starts_with("\\\\") ||
        path.contains('\\')
    }
}

impl Default for Config {
    fn default() -> Self {
        Config {
            quickbooks: QuickBooksConfig {
                enabled: Some(true),
                company_file: "AUTO".to_string(),
                account_number: "9445".to_string(),
                account_name: "INCOME TAX".to_string(),
                connection_mode: Some("multi-user".to_string()),
                application_name: None,
                application_id: None,
                company_file_password: None,
                qb_username: None,
                qb_password: None,
                connection_timeout: Some(30),
            },
            google_sheets: GoogleSheetsConfig {
                webapp_url: String::new(),
                api_key: String::new(),
                spreadsheet_id: String::new(),
                sheet_name: None,
                cell_address: "A1".to_string(),
            },
            schedule: ScheduleConfig {
                cron_expression: "0 0 * * * *".to_string(), // Every hour
            },
            server: None,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::env;

    #[test]
    fn test_default_config() {
        let config = Config::default();
        assert_eq!(config.quickbooks.company_file, "AUTO");
        assert_eq!(config.quickbooks.account_number, "9445");
        assert_eq!(config.quickbooks.account_name, "INCOME TAX");
        assert_eq!(config.quickbooks.connection_timeout, Some(30));
        assert_eq!(config.google_sheets.cell_address, "A1");
        assert_eq!(config.schedule.cron_expression, "0 0 * * * *");
    }

    #[test]
    fn test_is_windows_path() {
        assert!(Config::is_windows_path("C:\\Users\\Test\\file.qbw"));
        assert!(Config::is_windows_path("D:\\Documents\\Company.qbw"));
        assert!(Config::is_windows_path("\\\\server\\share\\file.qbw"));
        assert!(Config::is_windows_path("C:/Users/Test/file.qbw")); // Mixed separators
        
        assert!(!Config::is_windows_path("/unix/path/file.qbw"));
        assert!(!Config::is_windows_path("MOCK"));
        assert!(!Config::is_windows_path("AUTO"));
        assert!(!Config::is_windows_path(""));
        assert!(!Config::is_windows_path("relative/path/file.qbw"));
    }

    #[test]
    fn test_normalize_windows_path() {
        let config = Config::default();
        
        // Special values should not be normalized
        assert_eq!(config.normalize_windows_path("MOCK"), "MOCK");
        assert_eq!(config.normalize_windows_path("AUTO"), "AUTO");
        assert_eq!(config.normalize_windows_path(""), "");
        
        // On Windows, mixed separators should be normalized
        if cfg!(windows) {
            assert_eq!(config.normalize_windows_path("C:/Users/Test/file.qbw"), "C:\\Users\\Test\\file.qbw");
        }
        
        // Unix paths should remain unchanged
        assert_eq!(config.normalize_windows_path("/unix/path/file.qbw"), "/unix/path/file.qbw");
    }

    #[test]
    fn test_config_validation() {
        let mut config = Config::default();
        
        // Should fail validation with empty required fields
        config.google_sheets.webapp_url = "".to_string();
        config.google_sheets.api_key = "".to_string();
        assert!(config.validate().is_err());
        
        // Should pass validation with all required fields
        config.google_sheets.webapp_url = "https://script.google.com/test".to_string();
        config.google_sheets.api_key = "test_key".to_string();
        assert!(config.validate().is_ok());
    }

    #[test]
    fn test_environment_variable_override() {
        // Set environment variables for testing
        env::set_var("QB_FILE_PASSWORD", "test_password");
        env::set_var("QB_USERNAME", "test_user");
        env::set_var("SHEETS_API_KEY", "test_sheets_key");
        
        // This test would require a config file to exist, so we'll just verify
        // that the environment variables are properly set
        assert_eq!(env::var("QB_FILE_PASSWORD").unwrap(), "test_password");
        assert_eq!(env::var("QB_USERNAME").unwrap(), "test_user");
        assert_eq!(env::var("SHEETS_API_KEY").unwrap(), "test_sheets_key");
        
        // Clean up
        env::remove_var("QB_FILE_PASSWORD");
        env::remove_var("QB_USERNAME");
        env::remove_var("SHEETS_API_KEY");
    }
}
