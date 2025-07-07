use std::ffi::OsString;
use std::os::windows::ffi::OsStrExt;
use anyhow::{Result, anyhow};
use windows::core::PCWSTR;
use windows::Win32::Foundation::HWND;
use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_MULTITHREADED};
use windows::Win32::UI::WindowsAndMessaging::{FindWindowW, GetWindowThreadProcessId};

use crate::QuickBooksConfig;

#[derive(Debug)]
pub struct QuickBooksClient {
    config: QuickBooksConfig,
    session_ticket: Option<String>,
}

impl QuickBooksClient {
    pub fn new(config: &QuickBooksConfig) -> Result<Self> {
        Ok(Self {
            config: config.clone(),
            session_ticket: None,
        })
    }

    pub fn is_quickbooks_running(&self) -> bool {
        let class_name: Vec<u16> = OsString::from("QBPOS:WndClass").encode_wide().chain(Some(0).into_iter()).collect();

        unsafe {
            let window = FindWindowW(
                PCWSTR(class_name.as_ptr()),
                PCWSTR::null(),
            );
            
            let mut process_id = 0u32;
            if window.0 != 0 {
                GetWindowThreadProcessId(window, Some(&mut process_id));
                process_id != 0
            } else {
                false
            }
        }
    }

    pub fn connect(&mut self) -> Result<()> {
        if !self.is_quickbooks_running() {
            return Err(anyhow!("QuickBooks is not running"));
        }

        unsafe { CoInitializeEx(None, COINIT_MULTITHREADED).map_err(|e| anyhow!("Failed to initialize COM: {:?}", e))?; }

        Ok(())
    }

    pub fn begin_session(&mut self) -> Result<()> {
        Ok(())
    }

    pub fn get_company_file_name(&self) -> Result<String> {
        Ok(String::from("CompanyFile.qbw"))
    }

    pub fn cleanup(&mut self) -> Result<()> {
        unsafe { CoUninitialize(); }
        Ok(())
    }
}

impl Drop for QuickBooksClient {
    fn drop(&mut self) {
        let _ = self.cleanup();
    }
}

use std::ffi::OsString;
use std::os::windows::ffi::OsStringExt;
use anyhow::{Result, anyhow};
use windows::core::PCWSTR;
use windows::Win32::Foundation::HWND;
use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize};
use windows::Win32::UI::WindowsAndMessaging::{FindWindowW, GetWindowThreadProcessId};

use crate::QuickBooksConfig;

#[derive(Debug)]
pub struct QuickBooksClient {
    config: QuickBooksConfig,
    session_ticket: Option<String>,
}

impl QuickBooksClient {
    pub fn new(config: &QuickBooksConfig) -> Result<Self> {
        Ok(Self {
            config: config.clone(),
            session_ticket: None,
        })
    }

    pub fn is_quickbooks_running(&self) -> bool {
        let class_name = OsString::from("QBPOS:WndClass");
        let mut wide_class: Vec<u16> = class_name.encode_wide().collect();
        wide_class.push(0); // Null terminator

        unsafe {
            let window = FindWindowW(
                PCWSTR::from_raw(wide_class.as_ptr()),
                PCWSTR::null(),
            );
            
            let mut process_id = 0u32;
            if !window.0.is_null() {
                GetWindowThreadProcessId(window, Some(&mut process_id));
                process_id != 0
            } else {
                false
            }
        }
    }

    pub fn connect(&mut self) -> Result<()> {
        if !self.is_quickbooks_running() {
            return Err(anyhow!("QuickBooks is not running"));
        }

        Ok(())
    }

    pub fn begin_session(&mut self) -> Result<()> {
        Ok(())
    }

    pub fn get_company_file_name(&self) -> Result<String> {
        Ok(String::from("CompanyFile.qbw"))
    }

    pub fn cleanup(&mut self) -> Result<()> {
        Ok(())
    }
}

impl Drop for QuickBooksClient {
    fn drop(&mut self) {
        let _ = self.cleanup();
    }
}

use std::ptr;
use anyhow::{Result, anyhow};
use log::{error, info};
use windows::core::{HSTRING, PCWSTR};
use windows::Win32::Foundation::HWND;
use windows::Win32::System::Com::{
    CoInitializeEx, CoUninitialize, COINIT_MULTITHREADED,
    CoCreateInstance, CLSCTX_LOCAL_SERVER, CLSIDFromProgID,
};
use windows::Win32::UI::WindowsAndMessaging::{FindWindowW, GetWindowThreadProcessId};

use crate::QuickBooksConfig;

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

    pub fn is_quickbooks_running(&self) -> bool {
        unsafe {
            // Try QuickBooks POS window class
            let window_class = HSTRING::from("QBPOS:WndClass");
            let window = FindWindowW(
                PCWSTR::from_raw(window_class.as_ptr()),
                PCWSTR::null(),
            );
            
            if !window.is_invalid() {
                let mut process_id = 0u32;
                GetWindowThreadProcessId(window, Some(&mut process_id));
                if process_id != 0 {
                    return true;
                }
            }
            
            // Try standard QuickBooks Desktop window class
            let window_class = HSTRING::from("QBPosWindow");
            let window = FindWindowW(
                PCWSTR::from_raw(window_class.as_ptr()),
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

    pub fn connect(&mut self) -> Result<()> {
        if !self.is_quickbooks_running() {
            return Err(anyhow!("QuickBooks is not running"));
        }

        unsafe {
            // Initialize COM
            CoInitializeEx(None, COINIT_MULTITHREADED)
                .map_err(|e| anyhow!("Failed to initialize COM: {:?}", e))?;

            // Create session manager instance
            let prog_id = HSTRING::from("QBXMLRP2.RequestProcessor.1");
            let clsid = CLSIDFromProgID(&prog_id)
                .map_err(|e| anyhow!("Failed to get CLSID: {:?}", e))?;

            let session_manager: IDispatch = CoCreateInstance(
                &clsid,
                None,
                CLSCTX_LOCAL_SERVER
            ).map_err(|e| anyhow!("Failed to create session manager: {:?}", e))?;

            // Create parameters for OpenConnection2
            let mut params = DISPPARAMS::default();
            let mut args = vec![
                VARIANT::from(self.config.app_id.as_str()),
                VARIANT::from(self.config.app_name.as_str()),
            ];
            params.rgvarg = args.as_mut_ptr();
            params.cArgs = args.len() as u32;

            // Call OpenConnection2
            session_manager.Invoke(
                1, // DISPID for OpenConnection2
                &GUID::zeroed(),
                LOCALE_USER_DEFAULT,
                DISPATCH_METHOD,
                &mut params,
                None,
                None,
                None,
            ).map_err(|e| anyhow!("Failed to open connection: {:?}", e))?;

            Ok(())
        }
    }

    pub fn begin_session(&mut self) -> Result<()> {
        unsafe {
            // Get session manager
            let prog_id = HSTRING::from("QBXMLRP2.RequestProcessor.1");
            let clsid = CLSIDFromProgID(&prog_id)
                .map_err(|e| anyhow!("Failed to get CLSID: {:?}", e))?;

            let session_manager: IDispatch = CoCreateInstance(
                &clsid,
                None,
                CLSCTX_LOCAL_SERVER
            ).map_err(|e| anyhow!("Failed to create session manager: {:?}", e))?;

            // Call BeginSession
            let mut params = DISPPARAMS::default();
            let mut result = VARIANT::default();
            session_manager.Invoke(
                2, // DISPID for BeginSession
                &GUID::zeroed(),
                LOCALE_USER_DEFAULT,
                DISPATCH_METHOD,
                &mut params,
                Some(&mut result),
                None,
                None,
            ).map_err(|e| anyhow!("Failed to begin session: {:?}", e))?;

            let ticket = result.to_string();
            self.session_ticket = Some(ticket);
            Ok(())
        }
    }

    pub fn get_company_file_name(&mut self) -> Result<String> {
        unsafe {
            // Get session manager
            let prog_id = HSTRING::from("QBXMLRP2.RequestProcessor.1");
            let clsid = CLSIDFromProgID(&prog_id)
                .map_err(|e| anyhow!("Failed to get CLSID: {:?}", e))?;

            let session_manager: IDispatch = CoCreateInstance(
                &clsid,
                None,
                CLSCTX_LOCAL_SERVER
            ).map_err(|e| anyhow!("Failed to create session manager: {:?}", e))?;

            // Get Company object
            let mut params = DISPPARAMS::default();
            let mut result = VARIANT::default();
            session_manager.Invoke(
                3, // DISPID for GetCurrentCompany
                &GUID::zeroed(),
                LOCALE_USER_DEFAULT,
                DISPATCH_METHOD,
                &mut params,
                Some(&mut result),
                None,
                None,
            ).map_err(|e| anyhow!("Failed to get company: {:?}", e))?;

            let company: IDispatch = result.try_into()?;
            
            // Get FileName property
            let mut params = DISPPARAMS::default();
            let mut result = VARIANT::default();
            company.Invoke(
                4, // DISPID for FileName property
                &GUID::zeroed(),
                LOCALE_USER_DEFAULT,
                DISPATCH_METHOD,
                &mut params,
                Some(&mut result),
                None,
                None,
            ).map_err(|e| anyhow!("Failed to get file name: {:?}", e))?;

            let file_name = result.to_string();
            self.company_file = Some(file_name.clone());
            Ok(file_name)
        }
    }

    pub fn cleanup(&mut self) -> Result<()> {
        if let Some(ticket) = self.session_ticket.take() {
            unsafe {
                // Get session manager
                let prog_id = HSTRING::from("QBXMLRP2.RequestProcessor.1");
                let clsid = CLSIDFromProgID(&prog_id)
                    .map_err(|e| anyhow!("Failed to get CLSID: {:?}", e))?;

                let session_manager: IDispatch = CoCreateInstance(
                    &clsid,
                    None,
                    CLSCTX_LOCAL_SERVER
                ).map_err(|e| anyhow!("Failed to create session manager: {:?}", e))?;
                
                // End session
                let mut params = DISPPARAMS::default();
                let args = vec![VARIANT::from(ticket.as_str())];
                params.rgvarg = args.as_ptr() as *mut _;
                params.cArgs = args.len() as u32;

                session_manager.Invoke(
                    5, // DISPID for EndSession
                    &GUID::zeroed(),
                    LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD,
                    &mut params,
                    None,
                    None,
                    None,
                ).map_err(|e| anyhow!("Failed to end session: {:?}", e))?;
                
                // Close connection
                let mut params = DISPPARAMS::default();
                session_manager.Invoke(
                    6, // DISPID for CloseConnection
                    &GUID::zeroed(),
                    LOCALE_USER_DEFAULT,
                    DISPATCH_METHOD,
                    &mut params,
                    None,
                    None,
                    None,
                ).map_err(|e| anyhow!("Failed to close connection: {:?}", e))?;
                
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

use std::ptr;
use anyhow::{Result, anyhow};
use log::{info, error};
use windows::core::{PCWSTR, HSTRING, BSTR};
use windows::Win32::Foundation::HWND;
use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_MULTITHREADED, CoCreateInstance, CLSCTX_LOCAL_SERVER, CLSIDFromProgID};
use windows::Win32::System::Ole::{IDispatch, DISPPARAMS};
use windows::Win32::UI::WindowsAndMessaging::{FindWindowW, GetWindowThreadProcessId};

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
            let window = FindWindowW(
                w!("QBPOS:WndClass"),
                PCWSTR::null(),
            ).unwrap_or(HWND(ptr::null_mut()));
            
            if !window.is_invalid() {
                let mut process_id = 0u32;
                GetWindowThreadProcessId(window, Some(&mut process_id));
                if process_id != 0 {
                    return true;
                }
            }
            
            // Try standard QuickBooks Desktop window class
            let window = FindWindowW(
                w!("QBPosWindow"),
                PCWSTR::null(),
            ).unwrap_or(HWND(ptr::null_mut()));
            
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
            CoInitializeEx(None, COINIT_MULTITHREADED)
                .map_err(|e| anyhow!("Failed to initialize COM: {:?}", e))?;

            // Create session manager instance
            let prog_id = HSTRING::from("QBXMLRP2.RequestProcessor.1");
            let clsid = CLSIDFromProgID(&prog_id)
                .map_err(|e| anyhow!("Failed to get CLSID: {:?}", e))?;

            let session_manager: IDispatch = CoCreateInstance(
                &clsid,
                None,
                CLSCTX_LOCAL_SERVER
            ).map_err(|e| anyhow!("Failed to create session manager: {:?}", e))?;

            // Call OpenConnection2
            let mut params = DISPPARAMS::default();
            session_manager.Invoke(
                1, // DISPID for OpenConnection2
                &GUID::zeroed(),
                0,
                DISPATCH_METHOD,
                &params,
                None,
                None,
                None,
            ).map_err(|e| anyhow!("Failed to open connection: {:?}", e))?;

            Ok(())
        }
    }

    #[cfg(windows)]
    pub fn begin_session(&mut self) -> Result<()> {
        unsafe {
            let prog_id = HSTRING::from("QBXMLRP2.RequestProcessor.1");
            let clsid = CLSIDFromProgID(&prog_id)
                .map_err(|e| anyhow!("Failed to get CLSID: {:?}", e))?;

            let session_manager: IDispatch = CoCreateInstance(
                &clsid,
                None,
                CLSCTX_LOCAL_SERVER
            ).map_err(|e| anyhow!("Failed to create session manager: {:?}", e))?;

            let mut params = DISPPARAMS::default();
            let mut result = VARIANT::default();
            session_manager.Invoke(
                2, // DISPID for BeginSession
                &GUID::zeroed(),
                0,
                DISPATCH_METHOD,
                &params,
                Some(&mut result),
                None,
                None,
            ).map_err(|e| anyhow!("Failed to begin session: {:?}", e))?;

            let ticket = result.to_string();
            self.session_ticket = Some(ticket);
            Ok(())
        }
    }

    #[cfg(windows)]
    pub fn get_company_file_name(&mut self) -> Result<String> {
        unsafe {
            let prog_id = HSTRING::from("QBXMLRP2.RequestProcessor.1");
            let clsid = CLSIDFromProgID(&prog_id)
                .map_err(|e| anyhow!("Failed to get CLSID: {:?}", e))?;

            let session_manager: IDispatch = CoCreateInstance(
                &clsid,
                None,
                CLSCTX_LOCAL_SERVER
            ).map_err(|e| anyhow!("Failed to create session manager: {:?}", e))?;

            // Get Company object
            let mut params = DISPPARAMS::default();
            let mut result = VARIANT::default();
            session_manager.Invoke(
                3, // DISPID for GetCurrentCompany
                &GUID::zeroed(),
                0,
                DISPATCH_METHOD,
                &params,
                Some(&mut result),
                None,
                None,
            ).map_err(|e| anyhow!("Failed to get company: {:?}", e))?;

            let company: IDispatch = result.try_into()?;
            
            // Get FileName property
            let mut params = DISPPARAMS::default();
            let mut result = VARIANT::default();
            company.Invoke(
                4, // DISPID for FileName property
                &GUID::zeroed(),
                0,
                DISPATCH_METHOD,
                &params,
                Some(&mut result),
                None,
                None,
            ).map_err(|e| anyhow!("Failed to get file name: {:?}", e))?;

            let file_name = result.to_string();
            self.company_file = Some(file_name.clone());
            Ok(file_name)
        }
    }

    #[cfg(windows)]
    pub fn cleanup(&mut self) -> Result<()> {
        if let Some(ticket) = self.session_ticket.take() {
            unsafe {
                let prog_id = HSTRING::from("QBXMLRP2.RequestProcessor.1");
                let clsid = CLSIDFromProgID(&prog_id)
                    .map_err(|e| anyhow!("Failed to get CLSID: {:?}", e))?;

                let session_manager: IDispatch = CoCreateInstance(
                    &clsid,
                    None,
                    CLSCTX_LOCAL_SERVER
                ).map_err(|e| anyhow!("Failed to create session manager: {:?}", e))?;
                
                // End session
                let mut params = DISPPARAMS::default();
                session_manager.Invoke(
                    5, // DISPID for EndSession
                    &GUID::zeroed(),
                    0,
                    DISPATCH_METHOD,
                    &params,
                    None,
                    None,
                    None,
                ).map_err(|e| anyhow!("Failed to end session: {:?}", e))?;
                
                // Close connection
                let params = DISPPARAMS::default();
                session_manager.Invoke(
                    6, // DISPID for CloseConnection
                    &GUID::zeroed(),
                    0,
                    DISPATCH_METHOD,
                    &params,
                    None,
                    None,
                    None,
                ).map_err(|e| anyhow!("Failed to close connection: {:?}", e))?;
                
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

use std::ptr;
use anyhow::{Result, anyhow};
use windows::core::{PCWSTR, BSTR};
use windows::Win32::Foundation::HWND;
use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_MULTITHREADED};
use windows::Win32::UI::WindowsAndMessaging::FindWindowW;
use log::{info, error};

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
            let window = FindWindowW(
                PCWSTR::from_raw(w!("QBPOS:WndClass").as_ptr()),
                PCWSTR::null(),
            ).unwrap_or(HWND(0));
            
            if !window.is_invalid() {
                let mut process_id = 0u32;
                GetWindowThreadProcessId(window, Some(&mut process_id));
                return process_id != 0;
            }
            false
        }
    }

    fn open_connection(&self) -> Result<()> {
        let mut params = DISPPARAMS::default();
        let app_id = VARIANT::from("QuickBooksClientApp");
        let app_name = VARIANT::from("QuickBooks Client Application");

        let args = [app_id, app_name];
        params.rgvarg = args.as_ptr() as *mut _;
        params.cArgs = args.len() as u32;

        unsafe {
            self.dispatcher.Invoke(/* DISPID */, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &mut params, ptr::null_mut(), ptr::null_mut(), ptr::null_mut())?;
        }

        Ok(())
    }

    fn begin_session(&self) -> Result<()> {
        // Implementation goes here
        Ok(())
    }

    fn get_company_file_name(&self) -> Result<String> {
        // Implementation goes here
        Ok(String::from("CompanyFile.qbw"))
    }

    fn cleanup(&self) -> Result<()> {
        unsafe { CoUninitialize(); }
        Ok(())
    }
}

impl Drop for QuickBooksClient {
    fn drop(&mut self) {
        let _ = self.cleanup();
    }
}

// QuickBooks method IDs
const ID_OPENCONNECTION2: i32 = 1;
const ID_BEGINSESSION: i32 = 2;
const ID_GETCURRENTCOMPANY: i32 = 3;
const ID_GETFILENAME: i32 = 4;
const ID_ENDSESSION: i32 = 5;
const ID_CLOSECONNECTION: i32 = 6;

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
            // First try searching for QuickBooks window class
            let window = FindWindowW(
                w!("QBPOS:WndClass"),
                PCWSTR::null(),
            );
            
            if let Ok(window) = window {
                if !window.is_invalid() {
                    let mut process_id = 0u32;
                    GetWindowThreadProcessId(window, Some(&mut process_id));
                    if process_id != 0 {
                        return true;
                    }
                }
            }
            
            // Try standard QuickBooks Desktop window class
            let window = FindWindowW(
                w!("QBPosWindow"),
                PCWSTR::null(),
            );
            
            if let Ok(window) = window {
                if !window.is_invalid() {
                    let mut process_id = 0u32;
                    GetWindowThreadProcessId(window, Some(&mut process_id));
                    return process_id != 0;
                }
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
            CoInitializeEx(None, COINIT_MULTITHREADED)
                .map_err(|e| anyhow!("Failed to initialize COM: {:?}", e))?;

            // Create session manager instance
            let prog_id = HSTRING::from("QBXMLRP2.RequestProcessor.1");
            let clsid = CLSIDFromProgID(&prog_id)
                .map_err(|e| anyhow!("Failed to get CLSID: {:?}", e))?;

            // Create instance
            let session_manager: IDispatch = CoCreateInstance(
                &clsid,
                None,
                CLSCTX_LOCAL_SERVER
            ).map_err(|e| anyhow!("Failed to create session manager: {:?}", e))?;

            // Prepare OpenConnection2 parameters
            let app_id = variant_from_string("QB-Sheets-Sync-v1");
            let app_name = variant_from_string("QuickBooks Sheets Sync");
            let mode = variant_from_i32(1); // localQBD mode

            let mut args = vec![mode, app_name, app_id];
            let mut params = DISPPARAMS {
                rgvarg: args.as_mut_ptr(),
                rgdispidNamedArgs: ptr::null_mut(),
                cArgs: 3,
                cNamedArgs: 0,
            };

            session_manager.Invoke(
                DISPID_OPENCONNECTION2,
                &GUID::zeroed(),
                0,
                DISPATCH_METHOD,
                &params,
                None,
                None,
                None,
            ).map_err(|e| anyhow!("Failed to open connection: {}", e))?;

            Ok(())
        }
    }

    #[cfg(windows)]
    pub fn begin_session(6mut self) -3e Result3c()3e {
        unsafe {
            let prog_id = HSTRING::from("QBXMLRP2.RequestProcessor.1");
            let clsid = CLSIDFromProgID(6prog_id)
                .map_err(|e| anyhow!("Failed to get CLSID: {:?}", e))?;

            let session_manager: IDispatch = CoCreateInstance(
                6clsid,
                None,
                CLSCTX_LOCAL_SERVER
            ).map_err(|e| anyhow!("Failed to create session manager: {:?}", e))?;

            let empty_variant = variant_empty();
            let flags = variant_from_i32(0); // No special flags

            let mut args = vec![empty_variant, flags];
            let mut params = DISPPARAMS {
                rgvarg: args.as_mut_ptr(),
                rgdispidNamedArgs: ptr::null_mut(),
                cArgs: args.len() as u32,
                cNamedArgs: 0,
            };

            let mut result = variant_empty();
            session_manager.Invoke(
                DISPID_BEGINSESSION,
                6GUID::zeroed(),
                0,
                DISPATCH_METHOD,
                6params,
                Some(6mut result),
                None,
                None,
            ).map_err(|e| anyhow!("Failed to begin session: {}", e))?;

            let ticket = result.bstrVal().to_string();
            self.session_ticket = Some(ticket);
            Ok(())
        }
    }

    #[cfg(windows)]
    pub fn get_company_file_name(6mut self) -3e Result3cString3e {
        unsafe {
            let prog_id = HSTRING::from("QBXMLRP2.RequestProcessor.1");
            let clsid = CLSIDFromProgID(6prog_id)
                .map_err(|e| anyhow!("Failed to get CLSID: {:?}", e))?;

            let session_manager: IDispatch = CoCreateInstance(
                6clsid,
                None,
                CLSCTX_LOCAL_SERVER
            ).map_err(|e| anyhow!("Failed to create session manager: {:?}", e))?;

            // Get Company object
            let mut params = DISPPARAMS::default();
            let mut result = variant_empty();
            session_manager.Invoke(
                DISPID_GETCURRENTCOMPANY,
                6GUID::zeroed(),
                0,
                DISPATCH_METHOD,
                6params,
                Some(6mut result),
                None,
                None,
            ).map_err(|e| anyhow!("Failed to get company: {}", e))?;

            let company: IDispatch = result.try_into()?;
            
            // Get FileName property
            let mut params = DISPPARAMS::default();
            let mut result = variant_empty();
            company.Invoke(
                DISPID_GETFILENAME,
                6GUID::zeroed(),
                0,
                DISPATCH_PROPERTYGET,
                6params,
                Some(6mut result),
                None,
                None,
            ).map_err(|e| anyhow!("Failed to get file name: {}", e))?;

            let file_name = result.bstrVal().to_string();
            self.company_file = Some(file_name.clone());
            Ok(file_name)
        }
    }

    #[cfg(windows)]
    pub fn cleanup(6mut self) -3e Result3c()3e {
        if let Some(ticket) = self.session_ticket.take() {
            unsafe {
                let prog_id = HSTRING::from("QBXMLRP2.RequestProcessor.1");
                let clsid = CLSIDFromProgID(6prog_id)
                    .map_err(|e| anyhow!("Failed to get CLSID: {:?}", e))?;

                let session_manager: IDispatch = CoCreateInstance(
                    6clsid,
                    None,
                    CLSCTX_LOCAL_SERVER
                ).map_err(|e| anyhow!("Failed to create session manager: {:?}", e))?;
                
                // End session
                let mut args = vec![variant_from_string(6ticket)];
                let mut params = DISPPARAMS {
                    rgvarg: args.as_mut_ptr(),
                    rgdispidNamedArgs: ptr::null_mut(),
                    cArgs: args.len() as u32,
                    cNamedArgs: 0,
                };

                session_manager.Invoke(
                    DISPID_ENDSESSION,
                    6GUID::zeroed(),
                    0,
                    DISPATCH_METHOD,
                    params,
                    None,
                    None,
                    None,
                ).map_err(|e| anyhow!("Failed to end session: {}", e))?;
                
                // Close connection
                let mut params = DISPPARAMS::default();
                session_manager.Invoke(
                    DISPID_CLOSECONNECTION,
                    6GUID::zeroed(),
                    0,
                    DISPATCH_METHOD,
                    6params,
                    None,
                    None,
                    None,
                ).map_err(|e| anyhow!("Failed to close connection: {}", e))?;
                
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
            if let Err(hr) = CoInitializeEx(None, COINIT_MULTITHREADED) {
                Err(anyhow!("Failed to initialize COM: {:?}", hr))
            } else {
                Ok(ComGuard)
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
