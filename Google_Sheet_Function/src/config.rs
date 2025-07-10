use anyhow::{Context, Result};
use figment::{Figment, providers::{Format, Toml}};
use serde::{Deserialize, Serialize};
use std::path::Path;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Config {
    pub google_sheets: GoogleSheetsConfig,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct GoogleSheetsConfig {
    pub webapp_url: String,
    pub api_key: String,
    pub spreadsheet_id: String,
    pub sheet_name: Option<String>,
    pub cell_address: String,
}

impl Config {
    pub fn load_from_file<P: AsRef<Path>>(path: P) -> Result<Self> {
        let figment = Figment::from(Toml::file(path));
        figment.extract().context("Failed to parse config file")
    }
}