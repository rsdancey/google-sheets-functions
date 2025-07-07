use anyhow::Result;
use log::{info, error};
use std::ptr::null_mut;

#[cfg(windows)]
use anyhow::anyhow;
use windows::{
    core::{BSTR, PCWSTR, Interface},
    Win32::{
        Foundation::HWND,
        System::Com::{CoInitializeEx, CoUninitialize, COINIT_MULTITHREADED},
        UI::WindowsAndMessaging::{FindWindowW, GetWindowThreadProcessId},
    },
};

use crate::config::QuickBooksConfig;

#[derive(Debug)]
pub struct QuickBooksClient {
    config: QuickBooksConfig,
    session_ticket: Option<String>,
    company_file: Option<String>,
}

impl QuickBooksClient {
    pub fn new(config: &QuickBooksConfig) -> Result<Self> {
        Ok(Self {
            config: config.clone(),
            session_ticket: None,
            company_file: None,
        })
    }

    #[cfg(windows)]
    pub fn is_quickbooks_running(&self) -> bool {
        unsafe {
            // First try searching for the Point of Sale window class
            let window = FindWindowW(
                PCWSTR::from_raw(wide_str!("QBPOS:WndClass").as_ptr()),
                PCWSTR::null(),
            );
            
            if !window.is_invalid() {
                let mut process_id = 0u32;
                GetWindowThreadProcessId(window, Some(&mut process_id));
                return process_id != 0;
            }
            
            // Try standard QuickBooks Desktop window class
            let window = FindWindowW(
                PCWSTR::from_raw(wide_str!("QBPosWindow").as_ptr()),
                PCWSTR::null(),
            );
            
            if !window.is_invalid() {
                let mut process_id = 0u32;
                GetWindowThreadProcessId(window, Some(&mut process_id));
                return process_id != 0;
            }

            false
        }
    }

    #[cfg(windows)]
    pub fn connect(&mut self) -> Result<()> {
        if !self.is_quickbooks_running() {
            return Err(anyhow!("QuickBooks is not running"));
        }

        unsafe {
            // Initialize COM
            CoInitializeEx(None, COINIT_MULTITHREADED)?;

            // Create session manager instance
            let session_manager = windows::core::factory::<_, windows::core::IUnknown>("QBXMLRP2.RequestProcessor.1")?;

            // Call OpenConnection2
            let app_id = BSTR::from("Your Application ID");
            let app_name = BSTR::from("Your Application Name");
            session_manager.OpenConnection2(
                &app_id,
                &app_name,
                1, // localQBD mode
            )?;

            Ok(())
        }
    }

    #[cfg(windows)]
    pub fn begin_session(&mut self) -> Result<()> {
        unsafe {
            let session_manager = windows::core::factory::<_, windows::core::IUnknown>("QBXMLRP2.RequestProcessor.1")?;
            
            let mut ticket = BSTR::new();
            session_manager.BeginSession(
                null_mut(), // Use currently open company file
                0, // No special flags
                &mut ticket,
            )?;

            self.session_ticket = Some(ticket.to_string());
            Ok(())
        }
    }

    #[cfg(windows)]
    pub fn get_company_file_name(&mut self) -> Result<String> {
        unsafe {
            let session_manager = windows::core::factory::<_, windows::core::IUnknown>("QBXMLRP2.RequestProcessor.1")?;
            
            // Get Company object
            let company = session_manager.GetCurrentCompany()?;
            
            // Get FileName property
            let file_name = company.GetFileName()?;
            let file_name_str = file_name.to_string();
            
            self.company_file = Some(file_name_str.clone());
            Ok(file_name_str)
        }
    }

    #[cfg(windows)]
    pub fn cleanup(&mut self) -> Result<()> {
        if let Some(ticket) = self.session_ticket.take() {
            unsafe {
                let session_manager = windows::core::factory::<_, windows::core::IUnknown>("QBXMLRP2.RequestProcessor.1")?;
                
                // End session
                session_manager.EndSession(BSTR::from(&ticket))?;
                
                // Close connection
                session_manager.CloseConnection()?;
                
                // Uninitialize COM
                CoUninitialize();
            }
        }
        Ok(())
    }
}

impl Drop for QuickBooksClient {
    fn drop(&mut self) {
        let _ = self.cleanup();
    }
}

struct ComGuard;

#[cfg(windows)]
impl ComGuard {
    fn new() -> Result<Self> {
        unsafe {
            let result = CoInitializeEx(None, COINIT_MULTITHREADED);
            match result {
                Ok(_) => Ok(ComGuard),
                Err(e) => Err(anyhow!("Failed to initialize COM: {}", e))
            }
        }
    }
}

#[cfg(windows)]
impl Drop for ComGuard {
    fn drop(&mut self) {
        unsafe {
            CoUninitialize();
        }
    }
}
