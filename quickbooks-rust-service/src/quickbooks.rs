use anyhow::{Result, anyhow};
use std::mem::ManuallyDrop;
use windows_core::{BSTR, HSTRING, IUnknown};
use windows::core::PCWSTR;
use windows::Win32::System::Com::{
    CoInitialize, CoUninitialize,
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

    pub fn connect(&mut self, qb_file: &str) -> Result<()> {
        // Assume QuickBooks is running
        log::debug!("Attempting to connect to QuickBooks");

        unsafe {
            // Initialize COM with detailed error handling
            match CoInitialize(None) {
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

            // Create Request Processor exactly as in the sample
            log::debug!("Creating Request Processor");
            let prog_id = HSTRING::from("QBXMLRP2.RequestProcessor.2");
            let clsid = match CLSIDFromProgID(&prog_id) {
                Ok(clsid) => clsid,
                Err(e) => {
                    let msg = format!("Failed to get CLSID for Request Processor: 0x{:08X}", e.code().0);
                    log::error!("{}", msg);
                    CoUninitialize();
                    return Err(anyhow!(msg));
                }
            };

            // Create instance of Request Processor
            let request_processor = match CoCreateInstance::<Option<&IUnknown>, IDispatch>(&clsid, None, CLSCTX_LOCAL_SERVER) {
                Ok(rp) => {
                    self.request_processor = Some(rp.clone());
                    rp
                },
                Err(e) => {
                    let msg = format!("Failed to create Request Processor: 0x{:08X}", e.code().0);
                    log::error!("{}", msg);
                    CoUninitialize();
                    return Err(anyhow!(msg));
                }
            };

            // Open connection
            log::debug!("Opening connection");
            // Get IDs of methods we need
            let method_name = HSTRING::from("OpenConnection");
            let mut dispid = -1i32;
            let names = [PCWSTR::from_raw(method_name.as_ptr())];
            
            request_processor.GetIDsOfNames(
                &windows::core::GUID::zeroed(),
                names.as_ptr(),
                1,
                0x0409, // LCID_ENGLISH_US
                &mut dispid,
            )?;

            // Set up parameters exactly like sdktest.cpp
            let mut params = DISPPARAMS::default();
            let app_id = ManuallyDrop::new(BSTR::from(""));
            let app_name = ManuallyDrop::new(BSTR::from(&self.config.app_name));
            let mut args = vec![
                create_bstr_variant(""),             // appID (first parameter)
                create_bstr_variant(&self.config.app_name),  // appName (second parameter)
            ];
            params.rgvarg = args.as_mut_ptr();
            params.cArgs = args.len() as u32;

            let mut result = VARIANT::default();
            let mut exc_info = EXCEPINFO::default();
            let mut arg_err = 0;

            match request_processor.Invoke(
                dispid,
                &windows::core::GUID::zeroed(),
                0x0409,  // LCID_ENGLISH_US
                DISPATCH_METHOD,
                &mut params,
                Some(&mut result),
                Some(&mut exc_info),
                Some(&mut arg_err),
            ) {
                Ok(_) => {
                    log::debug!("Successfully opened connection");

                    // Begin session immediately after successful connection
                    log::debug!("Beginning session with company file: {}", qb_file);
                    log::debug!("COM array for BeginSession: [0]=BSTR('qbXMLModeEnter'), [1]=BSTR('')");
                    let mut params = DISPPARAMS::default();
                    // For COM, parameters are passed in reverse order
                    let mut args = vec![
                        create_bstr_variant("qbXMLModeEnter"),  // Mode (last parameter first)
                        create_bstr_variant(""),               // Empty company file (first parameter)
                    ];
                    for (i, arg) in args.iter().enumerate() {
                        log::debug!("BeginSession arg {}: BSTR", i);
                    }
                    params.rgvarg = args.as_mut_ptr();
                    params.cArgs = args.len() as u32;

                    let mut result = VARIANT::default();
                    let mut exc_info = EXCEPINFO::default();
                    let mut arg_err = 0;

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
                            // Store the session ticket
                            if let Ok(ticket) = variant_to_string(&result) {
                                self.session_ticket = Some(ticket);
                                log::debug!("Successfully started session");
                                Ok(())
                            } else {
                                let msg = "Failed to get session ticket";
                                log::error!("{}", msg);
                                CoUninitialize();
                                Err(anyhow!(msg))
                            }
                        }
                        Err(e) => {
                            let code = e.code().0;
                            if !exc_info.bstrDescription.is_empty() {
                                log::error!("Exception description: {:?}", exc_info.bstrDescription);
                            }
                            if !exc_info.bstrSource.is_empty() {
                                log::error!("Exception source: {:?}", exc_info.bstrSource);
                            }
                            if exc_info.scode != 0 {
                                log::error!("Exception scode: 0x{:08X}", exc_info.scode);
                            }
                            let msg = format!("Failed to begin session: 0x{:08X}", code);
                            log::error!("{}", msg);
                            CoUninitialize();
                            Err(anyhow!(msg))
                        }
                    }
                }
                Err(e) => {
                    let code = e.code().0;
                    if !exc_info.bstrDescription.is_empty() {
                        log::error!("Exception description: {:?}", exc_info.bstrDescription);
                    }
                    if !exc_info.bstrSource.is_empty() {
                        log::error!("Exception source: {:?}", exc_info.bstrSource);
                    }
                    if exc_info.scode != 0 {
                        log::error!("Exception scode: 0x{:08X}", exc_info.scode);
                    }
                    let msg = format!("Failed to open connection: 0x{:08X}", code);
                    log::error!("{}", msg);
                    CoUninitialize();
                    Err(anyhow!(msg))
                }
            }
        }
    }


    pub fn get_company_file_name(&mut self) -> Result<String> {
        unsafe {
            // Use QBXMLRP2.RequestProcessor.2 as shown in Windows registry for 64-bit QuickBooks Enterprise v24
            let prog_id = HSTRING::from("QBXMLRP2.RequestProcessor.2");
            let clsid = CLSIDFromProgID(&prog_id)
                .map_err(|e| anyhow!("Failed to get CLSID: {:?}", e))?;

            let session_manager: IDispatch = CoCreateInstance::<Option<&IUnknown>, IDispatch>(&clsid, None, CLSCTX_LOCAL_SERVER)
                .map_err(|e| anyhow!("Failed to create session manager: {:?}", e))?;

            let mut params = DISPPARAMS::default();
            let mut result = VARIANT::default();
            let mut exc_info = EXCEPINFO::default();
            let mut arg_err = 0u32;

            // GetCurrentCompany takes no parameters, so no need to reverse order
            log::debug!("Calling GetCurrentCompany()");
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
                    // For COM, parameters are passed in reverse order
                    let mut args = vec![create_bstr_variant(&ticket)];  // Single parameter, no need to reverse
                    params.rgvarg = args.as_mut_ptr();
                    params.cArgs = args.len() as u32;

                    let mut exc_info = EXCEPINFO::default();
                    let mut arg_err = 0u32;

                    log::debug!("COM array for EndSession: [0]=BSTR('{}')", ticket);
                    let _ = request_processor.Invoke(
                        5,  // DISPID for EndSession
                        &Default::default(),
                        0,
                        DISPATCH_METHOD,
                        &mut params,
                        None,  // No result needed
                        Some(&mut exc_info),
                        Some(&mut arg_err),
                    );

                    // Try to close connection - takes no parameters
                    let mut params = DISPPARAMS::default();
                    log::debug!("COM array for CloseConnection: []");
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
    let bstr = ManuallyDrop::new(BSTR::from(s));
    unsafe {
        let mut variant = VARIANT::default();
        let variant_anon = &mut variant.Anonymous;
        let variant_union = &mut variant_anon.Anonymous;
        
        variant_union.vt = VARENUM(VT_BSTR.0);
        variant_union.Anonymous.bstrVal = bstr;
        
        variant
    }
}

fn variant_to_string(variant: &VARIANT) -> Result<String> {
    unsafe {
        let variant_anon = &variant.Anonymous;
        let variant_union = &variant_anon.Anonymous;
        
        if variant_union.vt == VARENUM(VT_BSTR.0) {
            let bstr = &variant_union.Anonymous.bstrVal;
            return Ok(bstr.to_string());
        }
        Err(anyhow!("Failed to convert VARIANT to string"))
    }
}
