use anyhow::{Context, Result};
use figment::{Figment, providers::{Format, Toml}};
use serde::{Deserialize, Serialize};
use std::path::Path;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Config {
    pub quickbooks: QuickBooksConfig,
    pub google_sheets: GoogleSheetsConfig,
    pub schedule: ScheduleConfig,
    pub server: Option<ServerConfig>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct QuickBooksConfig {
    pub enabled: Option<bool>,
    pub company_file: String,
    pub account_number: String,
    pub account_name: String,
    pub connection_mode: Option<String>,
    pub application_name: Option<String>,
    pub application_id: Option<String>,
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
    pub port: Option<u16>,
    pub host: Option<String>,
}

impl Config {
    pub fn load_from_file<P: AsRef<Path>>(path: P) -> Result<Self> {
        let figment = Figment::from(Toml::file(path));
        figment.extract().context("Failed to parse config file")
    }
}
