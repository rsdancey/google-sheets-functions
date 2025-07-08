use anyhow::{Result, anyhow};
use std::mem::ManuallyDrop;
use std::ptr;
use windows_core::{w, BSTR, HSTRING};
use windows::core::PCWSTR;
use windows::Win32::System::Com::{
    CoInitializeEx, CoUninitialize, COINIT_MULTITHREADED,
    CoCreateInstance, CLSCTX_LOCAL_SERVER, CLSIDFromProgID,
    IDispatch, DISPPARAMS, EXCEPINFO, DISPATCH_METHOD,
};
use windows::Win32::System::Variant::{VARIANT, VARENUM, VT_BSTR};
use windows::Win32::UI::WindowsAndMessaging::{FindWindowW, GetWindowThreadProcessId};

use crate::QuickBooksConfig;

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
            let window_class = w!("QBPOS:WndClass");
            let window = FindWindowW(window_class, PCWSTR::null());

            if window.0 != 0 {
                let mut process_id = 0u32;
                GetWindowThreadProcessId(window, Some(&mut process_id));
                if process_id != 0 {
                    return true;
                }
            }

            let window_class = w!("QBPosWindow");
            let window = FindWindowW(window_class, PCWSTR::null());

            if window.0 != 0 {
                let mut process_id = 0u32;
                GetWindowThreadProcessId(window, Some(&mut process_id));
                return process_id != 0;
            }

            false
        }
    }

pub fn connect(&mut self) -> Result<()> {
        if !self.is_quickbooks_running() {
            return Err(anyhow!("ERROR: QuickBooks is not running. Please start QuickBooks and open a company file."));
        }

        unsafe {
            // Initialize COM with detailed error handling
            match CoInitializeEx(None, COINIT_MULTITHREADED) {
                Ok(_) => log::debug!("COM initialized successfully"),
                Err(e) => {
                    let error_msg = format!("COM initialization failed with error code 0x{:08X}", e.code().0);
                    log::error!("{}", error_msg);
                    return Err(anyhow!(error_msg));
                }
            };

            // Get CLSID with detailed error handling
            let prog_id = HSTRING::from("QBXMLRP2.RequestProcessor.1");
            let clsid = match CLSIDFromProgID(&prog_id) {
                Ok(id) => id,
                Err(e) => {
                    let error_msg = format!(
                        "Failed to get CLSID for QBXMLRP2.RequestProcessor.1. Error: 0x{:08X}. \
                        This usually means the QuickBooks SDK is not properly installed or registered.",
                        e.code().0
                    );
                    log::error!("{}", error_msg);
                    CoUninitialize();
                    return Err(anyhow!(error_msg));
                }
            };

            // Create instance with detailed error handling
            let session_manager: IDispatch = match CoCreateInstance(
                &clsid,
                None,
                CLSCTX_LOCAL_SERVER
            ) {
                Ok(instance) => instance,
                Err(e) => {
                    let error_msg = format!(
                        "Failed to create QuickBooks session manager. Error: 0x{:08X}. \
                        This could mean QuickBooks is not running or the SDK components are not registered.",
                        e.code().0
                    );
                    log::error!("{}", error_msg);
                    CoUninitialize();
                    return Err(anyhow!(error_msg));
                }
            };
            CoInitializeEx(None, COINIT_MULTITHREADED)
                .map_err(|e| anyhow!("Failed to initialize COM: {:?}", e))?;

            let prog_id = HSTRING::from("QBXMLRP2.RequestProcessor.1");
            let clsid = match CLSIDFromProgID(&prog_id) {
                Ok(id) => id,
                Err(e) => {
                    let error_msg = format!(
                        "Failed to get CLSID for session initialization. Error: 0x{:08X}. \
                        The QuickBooks SDK components may not be properly registered.",
                        e.code().0
                    );
                    log::error!("{}", error_msg);
                    return Err(anyhow!(error_msg));
                }
            };

            let session_manager: IDispatch = match CoCreateInstance(
                &clsid,
                None,
                CLSCTX_LOCAL_SERVER
            ) {
                Ok(instance) => instance,
                Err(e) => {
                    let error_msg = format!(
                        "Failed to create session manager for BeginSession. Error: 0x{:08X}. \
                        Check if QuickBooks is running and accessible.",
                        e.code().0
                    );
                    log::error!("{}", error_msg);
                    return Err(anyhow!(error_msg));
                }
            };

            let mut params = DISPPARAMS::default();
            let mut args = vec![
                create_bstr_variant(&self.config.app_id),
                create_bstr_variant(&self.config.app_name),
            ];
            params.rgvarg = args.as_mut_ptr();
            params.cArgs = args.len() as u32;

            let mut result = VARIANT::default();
            let mut exc_info = EXCEPINFO::default();
            let mut arg_err = 0u32;

            session_manager.Invoke(
                1,  // DISPID for OpenConnection2
                &Default::default(),  // GUID for IID_NULL
                0,  // LOCALE_SYSTEM_DEFAULT
                DISPATCH_METHOD,
                &mut params,
                Some(&mut result),
                Some(&mut exc_info),
                Some(&mut arg_err),
            ).map_err(|e| anyhow!("Failed to open connection: {:?}", e))?;

            Ok(())
        }
    }

pub fn begin_session(&mut self) -> Result<()> {
        unsafe {
            log::debug!("Starting QuickBooks session");
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
            let mut exc_info = EXCEPINFO::default();
            let mut arg_err = 0u32;

            session_manager.Invoke(
                2,  // DISPID for BeginSession
                &Default::default(),  // GUID for IID_NULL
                0,  // LOCALE_SYSTEM_DEFAULT
                DISPATCH_METHOD,
                &mut params,
                Some(&mut result),
                Some(&mut exc_info),
                Some(&mut arg_err),
            ).map_err(|e| anyhow!("Failed to begin session: {:?}", e))?;

            let ticket = variant_to_string(&result)?;
            self.session_ticket = Some(ticket);
            Ok(())
        }
    }

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

            let mut params = DISPPARAMS::default();
            let mut result = VARIANT::default();
            let mut exc_info = EXCEPINFO::default();
            let mut arg_err = 0u32;

            session_manager.Invoke(
                3,  // DISPID for GetCurrentCompany
                &Default::default(),  // GUID for IID_NULL
                0,  // LOCALE_SYSTEM_DEFAULT
                DISPATCH_METHOD,
                &mut params,
                Some(&mut result),
                Some(&mut exc_info),
                Some(&mut arg_err),
            ).map_err(|e| anyhow!("Failed to get company: {:?}", e))?;

            let file_name = variant_to_string(&result)?;
            self.company_file = Some(file_name.clone());
            Ok(file_name)
        }
    }

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

                let mut params = DISPPARAMS::default();
                let mut args = vec![create_bstr_variant(&ticket)];
                params.rgvarg = args.as_mut_ptr();
                params.cArgs = args.len() as u32;

                let mut result = VARIANT::default();
                let mut exc_info = EXCEPINFO::default();
                let mut arg_err = 0u32;

                session_manager.Invoke(
                    5,  // DISPID for EndSession
                    &Default::default(),  // GUID for IID_NULL
                    0,  // LOCALE_SYSTEM_DEFAULT
                    DISPATCH_METHOD,
                    &mut params,
                    Some(&mut result),
                    Some(&mut exc_info),
                    Some(&mut arg_err),
                ).map_err(|e| anyhow!("Failed to end session: {:?}", e))?;

                let mut params = DISPPARAMS::default();
                let mut result = VARIANT::default();
                let mut exc_info = EXCEPINFO::default();
                let mut arg_err = 0u32;

                session_manager.Invoke(
                    6,  // DISPID for CloseConnection
                    &Default::default(),  // GUID for IID_NULL
                    0,  // LOCALE_SYSTEM_DEFAULT
                    DISPATCH_METHOD,
                    &mut params,
                    Some(&mut result),
                    Some(&mut exc_info),
                    Some(&mut arg_err),
                ).map_err(|e| anyhow!("Failed to close connection: {:?}", e))?;

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

fn create_bstr_variant(s: &str) -> VARIANT {
    unsafe {
        let mut variant = VARIANT::default();
        let bstr = ManuallyDrop::new(BSTR::from(s));
        
        // Get raw pointers to the union fields and write values
        let var_union_ptr = ptr::addr_of_mut!(variant.Anonymous);
        let var_union2_ptr = ptr::addr_of_mut!((*var_union_ptr).Anonymous);
        let var_union3_ptr = ptr::addr_of_mut!((*var_union2_ptr).Anonymous);
        
        // Write the variant type
        ptr::write(ptr::addr_of_mut!((*var_union2_ptr).vt), VARENUM(VT_BSTR.0));
        
        // Write the BSTR value
        ptr::write(ptr::addr_of_mut!((*var_union3_ptr).bstrVal), bstr);
        
        variant
    }
}

fn variant_to_string(variant: &VARIANT) -> Result<String> {
    unsafe {
        let var_union_ptr = ptr::addr_of!(variant.Anonymous);
        let var_union2_ptr = ptr::addr_of!((*var_union_ptr).Anonymous);
        
        if (*var_union2_ptr).vt == VARENUM(VT_BSTR.0) {
            let var_union3_ptr = ptr::addr_of!((*var_union2_ptr).Anonymous);
            let bstr = &(*var_union3_ptr).bstrVal;
            return Ok(bstr.to_string());
        }
        Err(anyhow!("Failed to convert VARIANT to string"))
    }
}
