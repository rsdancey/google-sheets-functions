use anyhow::{Result, anyhow};
use std::mem::ManuallyDrop;
use std::ptr;
use windows_core::{BSTR, HSTRING, IUnknown};
use windows::core::PCWSTR;
use windows::Win32::System::Com::{
    CoInitializeEx, CoUninitialize, COINIT_MULTITHREADED,
    CoCreateInstance, CLSCTX_ALL, CLSIDFromProgID,
    IDispatch, DISPPARAMS, EXCEPINFO, DISPATCH_METHOD,
};
use windows::Win32::System::Registry::HKEY;
use windows::Win32::System::Variant::{VARIANT, VARENUM, VT_BSTR};
use windows::Win32::UI::WindowsAndMessaging::{FindWindowW, GetWindowThreadProcessId};

use crate::QuickBooksConfig;

pub struct QuickBooksClient {
    config: QuickBooksConfig,
    session_ticket: Option<String>,
    company_file: Option<String>,
    session_manager: Option<IDispatch>,
    request_processor: Option<IDispatch>,
    is_com_initialized: bool,
}

impl QuickBooksClient {
    pub fn new(config: &QuickBooksConfig) -> Result<Self> {
        Ok(Self {
            config: config.clone(),
            session_ticket: None,
            company_file: None,
            session_manager: None,
            request_processor: None,
            is_com_initialized: false,
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
                Ok(_) => {
                    log::debug!("COM initialized successfully");
                    self.is_com_initialized = true;
                },
                Err(e) => {
                    let error_msg = format!("COM initialization failed with error code 0x{:08X}", e.code().0);
                    log::error!("{}", error_msg);
                    return Err(anyhow!(error_msg));
                }
            };

            // Try to create Request Processor
            let request_processor_id = "QBXMLRP2.RequestProcessor";
            log::debug!("Attempting to create Session Manager");

            // Create Request Processor
            log::debug!("Creating Request Processor");

            // Create Request Processor
            let prog_id = HSTRING::from(request_processor_id);
            let clsid = match CLSIDFromProgID(&prog_id) {
                Ok(clsid) => clsid,
                Err(e) => {
                    let msg = format!("Failed to get CLSID for Request Processor: 0x{:08X}", e.code().0);
                    log::error!("{}", msg);
                    CoUninitialize();
                    return Err(anyhow!(msg));
                }
            };

            let request_processor = match CoCreateInstance::<Option<&IUnknown>, IDispatch>(&clsid, None, CLSCTX_ALL) {
                Ok(rp) => {
                    self.request_processor = Some(rp.clone());
                    rp
                },
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

            // Open connection using the Request Processor
            let mut params = DISPPARAMS::default();
            let mut args = vec![
                create_bstr_variant(&self.config.app_name),  // Args are in reverse order for COM
                create_bstr_variant(&self.config.app_id),
            ];
            params.rgvarg = args.as_mut_ptr();
            params.cArgs = args.len() as u32;

            let mut result = VARIANT::default();
            let mut exc_info = EXCEPINFO::default();
            let mut arg_err = 0;

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
                    let code = e.code().0;
                    // Log the exception info
                    if !exc_info.bstrDescription.is_empty() {
                        log::error!("Exception description: {:?}", exc_info.bstrDescription);
                    }
                    if !exc_info.bstrSource.is_empty() {
                        log::error!("Exception source: {:?}", exc_info.bstrSource);
                    }
                    if exc_info.scode != 0 {
                        log::error!("Exception scode: 0x{:08X}", exc_info.scode);
                    }
                    
                    let msg = if let Some(info) = e.info() {
                        format!("Failed to open connection: 0x{:08X} - {:?}", code, info)
                    } else {
                        format!("Failed to open connection: 0x{:08X}", code)
                    };
                    log::error!("{}", msg);
                    CoUninitialize();
                    Err(anyhow!(msg))
                }
            }
        }
    }

    pub fn begin_session(&mut self, qb_file: &str) -> Result<()> {
        unsafe {
            log::debug!("Starting QuickBooks session");
            
            if let Some(request_processor) = &self.request_processor {
                let mut params = DISPPARAMS::default();
                let mut args = vec![
                    create_bstr_variant(qb_file),
create_bstr_variant("DoNotCare")
                ];
                params.rgvarg = args.as_mut_ptr();
                params.cArgs = args.len() as u32;

                let mut result = VARIANT::default();
                let mut exc_info = EXCEPINFO::default();
                let mut arg_err = 0u32;

                // Begin session
                match request_processor.Invoke(
                    2,  // DISPID for BeginSession
                    &Default::default(),
                    0,
                    DISPATCH_METHOD,
                    &mut params,
                    Some(&mut result),
                    Some(&mut exc_info),
                    Some(&mut arg_err),
                ) {
                    Ok(_) => {
                        let ticket = variant_to_string(&result)?;
                        self.session_ticket = Some(ticket);

                        // Get company file name (optional but useful for debugging)
                        let mut params = DISPPARAMS::default();
                        if let Some(ticket) = &self.session_ticket {
                            let mut args = vec![create_bstr_variant(ticket)];
                            params.rgvarg = args.as_mut_ptr();
                            params.cArgs = args.len() as u32;

                            let mut result = VARIANT::default();
                            match request_processor.Invoke(
                                3,  // DISPID for GetCurrentCompanyFileName
                                &Default::default(),
                                0,
                                DISPATCH_METHOD,
                                &mut params,
                                Some(&mut result),
                                Some(&mut exc_info),
                                Some(&mut arg_err),
                            ) {
                                Ok(_) => {
                                    if let Ok(company_file) = variant_to_string(&result) {
                                        log::debug!("Connected to company file: {}", company_file);
                                        self.company_file = Some(company_file);
                                    }
                                }
                                Err(e) => {
                                    log::warn!("Could not get company file name: {:?}", e);
                                }
                            }
                        }
                        Ok(())
                    }
                    Err(e) => {
                        let msg = format!("Failed to begin session: {:?}", e);
                        log::error!("{}", msg);
                        Err(anyhow!(msg))
                    }
                }
            } else {
                Err(anyhow!("Request Processor not initialized. Call connect() first."))
            }
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

}
impl Drop for QuickBooksClient {
    fn drop(&mut self) {
        unsafe {
            // First try to close any active session
            if let Some(ticket) = self.session_ticket.take() {
                if let Some(request_processor) = self.request_processor.take() {
                    // Try to end the session
                    let mut params = DISPPARAMS::default();
                    let mut args = vec![create_bstr_variant(&ticket)];
                    params.rgvarg = args.as_mut_ptr();
                    params.cArgs = args.len() as u32;

                    let mut result = VARIANT::default();
                    let mut exc_info = EXCEPINFO::default();
                    let mut arg_err = 0u32;

                    let _ = request_processor.Invoke(
                        5,  // DISPID for EndSession
                        &Default::default(),
                        0,
                        DISPATCH_METHOD,
                        &mut params,
                        Some(&mut result),
                        Some(&mut exc_info),
                        Some(&mut arg_err),
                    );

                    // Try to close connection
                    let mut params = DISPPARAMS::default();
                    let _ = request_processor.Invoke(
                        6,  // DISPID for CloseConnection
                        &Default::default(),
                        0,
                        DISPATCH_METHOD,
                        &mut params,
                        None,
                        None,
                        None,
                    );
                }
            }

            // Release COM objects in reverse order of creation
            if let Some(request_processor) = self.request_processor.take() {
                drop(request_processor);
            }
            if let Some(session_manager) = self.session_manager.take() {
                drop(session_manager);
            }

            // Finally uninitialize COM if we initialized it
            if self.is_com_initialized {
                CoUninitialize();
            }
        }
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
