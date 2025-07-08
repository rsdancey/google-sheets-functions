mod com_helpers;
pub mod request_processor;

use anyhow::{Result, anyhow};
use serde::Deserialize;
use windows::Win32::System::Com::{CoInitialize, CoUninitialize};
use windows::Win32::UI::WindowsAndMessaging::{FindWindowW, GetWindowThreadProcessId};
use windows_core::{HSTRING, PCWSTR};

#[derive(Debug, Clone, Deserialize)]
pub struct QuickBooksConfig {
    pub app_name: String,
}

#[derive(Debug, Clone, PartialEq)]
pub enum ConnectionType {
    Local,           // localQBD
    LocalWithUI,     // localQBDLaunchUI
    Remote,          // remoteQBD
    RemoteQBOE,      // remoteQBOE
    Unknown,         // For default/unspecified
}

#[derive(Debug, Clone, PartialEq)]
pub enum FileMode {
    SingleUser,
    MultiUser,
    DoNotCare,
}

/// Authentication preferences for QuickBooks connection
#[derive(Debug, Clone)]
pub struct AuthPreferences {
    pub is_read_only: bool,
    pub enterprise_enabled: bool,
    pub premier_enabled: bool,
    pub pro_enabled: bool,
    pub simple_enabled: bool,
    pub force_auth_dialog: bool,
}

#[allow(dead_code)]
pub struct QuickBooksClient {
    config: QuickBooksConfig,
request_processor: Option<RequestProcessor2>,
    session_ticket: Option<String>,
    company_file: Option<String>,
    is_com_initialized: bool,
    connection_type: ConnectionType,
    file_mode: FileMode,
    auth_preferences: AuthPreferences,
}

impl QuickBooksClient {
    pub fn new(config: &QuickBooksConfig) -> Result<Self> {
        Ok(Self {
            config: config.clone(),
            request_processor: None,
            session_ticket: None,
            company_file: None,
            is_com_initialized: false,
            connection_type: ConnectionType::Local,
            file_mode: FileMode::DoNotCare,
            auth_preferences: AuthPreferences {
                is_read_only: false,
                enterprise_enabled: true,
                premier_enabled: false,
                pro_enabled: false,
                simple_enabled: false,
                force_auth_dialog: false,
            },
        })
    }

    #[allow(dead_code)]
    pub fn is_quickbooks_running(&self) -> bool {
        unsafe {
            // Try various known QuickBooks window classes
            let window_classes = [
                HSTRING::from("QBPOS:WndClass"),       // POS version
                HSTRING::from("QBPosWindow"),          // POS alternative
                HSTRING::from("QuickBooks:WndClass"),  // Desktop version
                HSTRING::from("QBHLPR:WndClass"),      // Helper window
                HSTRING::from("QBHelperWndClass"),     // Alternative helper
                HSTRING::from("QuickBooksWindows"),    // Multi-user mode
            ];

            for class_name in window_classes.iter() {
                let window = FindWindowW(PCWSTR::from_raw(class_name.as_ptr()), PCWSTR::null());
                if window.0 != 0 {
                    let mut process_id = 0u32;
                    GetWindowThreadProcessId(window, Some(&mut process_id));
                    if process_id != 0 {
                        log::debug!("Found QuickBooks window with class: {}", class_name);
                        return true;
                    }
                }
            }

            log::warn!("No QuickBooks windows found. Checked classes: {:?}", window_classes);
            false
        }
    }

    pub fn connect(&mut self, qb_file: &str) -> Result<()> {
        log::debug!("Attempting to connect to QuickBooks");

        // Initialize COM
        unsafe { CoInitialize(None)? };
        self.is_com_initialized = true;

        // Create request processor
        let request_processor = RequestProcessor2::new()
            .map_err(|e| anyhow!("Failed to create request processor: {:?}", e))?;

        // Open connection
        request_processor.open_connection("", &self.config.app_name)
            .map_err(|e| anyhow!("Failed to open connection: {:?}", e))?;

        // Begin session
        let ticket = request_processor.begin_session(qb_file, self.file_mode.clone())
            .map_err(|e| anyhow!("Failed to begin session: {:?}", e))?;

        self.session_ticket = Some(ticket);
        self.request_processor = Some(request_processor);

        Ok(())
    }

    pub fn get_company_file_name(&self) -> Result<String> {
        if let Some(rp) = &self.request_processor {
            if let Some(ticket) = &self.session_ticket {
                rp.get_current_company_file_name(ticket)
                    .map_err(|e| anyhow!("Failed to get company file name: {:?}", e))
            } else {
                Err(anyhow!("No active session"))
            }
        } else {
            Err(anyhow!("Request processor not initialized"))
        }
    }
}

impl Drop for QuickBooksClient {
    fn drop(&mut self) {
        if let Some(rp) = self.request_processor.take() {
            if let Some(ticket) = self.session_ticket.take() {
                let _ = rp.end_session(&ticket);
                let _ = rp.close_connection();
            }
        }

        if self.is_com_initialized {
            unsafe { CoUninitialize(); }
        }
    }
}
