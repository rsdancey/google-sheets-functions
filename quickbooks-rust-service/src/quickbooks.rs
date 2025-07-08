use anyhow::{Result, anyhow};
use std::mem::ManuallyDrop;
use std::ptr;
use windows_core::{BSTR, HSTRING, IUnknown};
use windows::core::{PCWSTR, PCSTR};
use windows::Win32::System::Com::{
    CoInitializeEx, CoUninitialize, COINIT_MULTITHREADED,
    CoCreateInstance, CLSCTX_ALL, CLSIDFromProgID,
    IDispatch, DISPPARAMS, EXCEPINFO, DISPATCH_METHOD,
};
use windows::Win32::System::Registry::{HKEY, HKEY_LOCAL_MACHINE, KEY_READ, RegOpenKeyExA, RegCloseKey};
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

            // Also check for the database manager which runs in multi-user mode
            let db_class = HSTRING::from("QBDBMgrWindow");
            let db_window = FindWindowW(PCWSTR::from_raw(db_class.as_ptr()), PCWSTR::null());
            if db_window.0 != 0 {
                let mut process_id = 0u32;
                GetWindowThreadProcessId(db_window, Some(&mut process_id));
                if process_id != 0 {
                    log::debug!("Found QuickBooks Database Manager window");
                    return true;
                }
            }

            log::warn!("No QuickBooks windows found. Checked classes: {:?}", window_classes);
            false
        }
    }

    pub fn connect(&mut self) -> Result<()> {
        // Assume QuickBooks is running
        log::debug!("Attempting to connect to QuickBooks");

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

            // Try to create Session Manager
            let session_manager_id = "QBXMLRP2.SessionManager";
            log::debug!("Attempting to create Session Manager");

            // Check registry for Session Manager
            let mut hkey = HKEY::default();
            let key_path = format!("SOFTWARE\\Classes\\{}\\CLSID", session_manager_id) + "\0";
            let result = RegOpenKeyExA(
                HKEY_LOCAL_MACHINE,
                PCSTR(key_path.as_ptr()),
                0,
                KEY_READ,
                &mut hkey
            );

            if result.is_err() {
                let msg = format!("QuickBooks Session Manager not found in registry. Please ensure QuickBooks and the SDK v16 are properly installed.");
                log::error!("{}", msg);
                CoUninitialize();
                return Err(anyhow!(msg));
            }
            RegCloseKey(hkey);

            // Create Session Manager
            let prog_id = HSTRING::from(session_manager_id);
            let clsid = match CLSIDFromProgID(&prog_id) {
                Ok(clsid) => clsid,
                Err(e) => {
                    let msg = format!("Failed to get CLSID for Session Manager: 0x{:08X}", e.code().0);
                    log::error!("{}", msg);
                    CoUninitialize();
                    return Err(anyhow!(msg));
                }
            };

            let session_manager = match CoCreateInstance::<Option<&IUnknown>, IDispatch>(&clsid, None, CLSCTX_ALL) {
                Ok(sm) => sm,
                Err(e) => {
                    let code = e.code().0;
                    let msg = if code == -2147221164i32 {
                        format!(
                            "Failed to create Session Manager - COM class not registered (0x80040154). \
                            This usually means either:\n\
                            1. QuickBooks is not installed\n\
                            2. QuickBooks SDK is not installed\n\
                            3. The SDK components are not properly registered"
                        )
                    } else {
                        format!("Failed to create Session Manager: 0x{:08X}", code)
                    };
                    log::error!("{}", msg);
                    CoUninitialize();
                    return Err(anyhow!(msg));
                }
            };

            // Now get the Request Processor from the Session Manager
            let mut params = DISPPARAMS::default();
            let mut result = VARIANT::default();
            let mut exc_info = EXCEPINFO::default();
            let mut arg_err = 0u32;

            match session_manager.Invoke(
                1,  // DISPID for GetRequestProcessor
                &Default::default(),
                0,
                DISPATCH_METHOD,
                &mut params,
                Some(&mut result),
                Some(&mut exc_info),
                Some(&mut arg_err),
            ) {
                Ok(_) => {
                    match result.try_into() {
                        Ok(request_processor) => {
                            // Set up connection parameters for the Request Processor
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

                            // Open connection using the Request Processor
                            match request_processor.Invoke(
                                1,  // DISPID for OpenConnection2
                                &Default::default(),
                                0,
                                DISPATCH_METHOD,
                                &mut params,
                                Some(&mut result),
                                Some(&mut exc_info),
                                Some(&mut arg_err),
                            ) {
                                Ok(_) => {
                                    log::debug!("Successfully connected to QuickBooks");
                                    Ok(())
                                }
                                Err(e) => {
                                    let msg = format!("Failed to open connection: {:?}", e);
                                    log::error!("{}", msg);
                                    CoUninitialize();
                                    Err(anyhow!(msg))
                                }
                            }
                        }
                        Err(e) => {
                            let msg = format!("Failed to get Request Processor from Session Manager: {:?}", e);
                            log::error!("{}", msg);
                            CoUninitialize();
                            Err(anyhow!(msg))
                        }
                    }
                }
                Err(e) => {
                    let msg = format!("Failed to get Request Processor: {:?}", e);
                    log::error!("{}", msg);
                    CoUninitialize();
                    Err(anyhow!(msg))
                }
            }

            // If we get here, none of the ProgIDs worked
            CoUninitialize();
            Err(anyhow!(last_error.unwrap_or_else(|| 
                "Failed to connect with any known QuickBooks SDK components".to_string()
            )))
        }
    }

    pub fn begin_session(&mut self) -> Result<()> {
        unsafe {
            log::debug!("Starting QuickBooks session");
            let prog_id = HSTRING::from("QBXMLRP2.RequestProcessor.1");
            let clsid = CLSIDFromProgID(&prog_id)
                .map_err(|e| anyhow!("Failed to get CLSID: {:?}", e))?;

            let session_manager: IDispatch = CoCreateInstance::<Option<&IUnknown>, IDispatch>(&clsid, None, CLSCTX_ALL)
                .map_err(|e| anyhow!("Failed to create session manager: {:?}", e))?;

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

            let session_manager: IDispatch = CoCreateInstance::<Option<&IUnknown>, IDispatch>(&clsid, None, CLSCTX_ALL)
                .map_err(|e| anyhow!("Failed to create session manager: {:?}", e))?;

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

                let session_manager: IDispatch = CoCreateInstance::<Option<&IUnknown>, IDispatch>(&clsid, None, CLSCTX_ALL)
                    .map_err(|e| anyhow!("Failed to create session manager: {:?}", e))?;

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
        
        let var_union_ptr = ptr::addr_of_mut!(variant.Anonymous);
        let var_union2_ptr = ptr::addr_of_mut!((*var_union_ptr).Anonymous);
        let var_union3_ptr = ptr::addr_of_mut!((*var_union2_ptr).Anonymous);
        
        ptr::write(ptr::addr_of_mut!((*var_union2_ptr).vt), VARENUM(VT_BSTR.0));
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
