use windows::core::{HSTRING, IUnknown, PCWSTR, PCSTR};
use windows::Win32::System::Com::{
    CLSIDFromProgID, CoCreateInstance,
    CLSCTX_LOCAL_SERVER, CLSCTX_INPROC_SERVER, CLSCTX_ALL,
    IDispatch, DISPATCH_METHOD, EXCEPINFO,
};
use windows::Win32::System::Registry::*;
use std::path::Path;
use windows::Win32::System::Variant::VARIANT;
use crate::com_helpers::{create_bstr_variant, create_dispparams, create_empty_dispparams, variant_to_string};
use crate::FileMode;

/// Type-safe wrapper for QBXMLRP2 RequestProcessor2
pub struct RequestProcessor2 {
    inner: IDispatch,
    // Cache method IDs after first lookup
    open_connection_id: i32,
    begin_session_id: i32,
    end_session_id: i32,
    close_connection_id: i32,
    process_request_id: i32,
}

impl RequestProcessor2 {
    fn check_sdk_installation() -> windows::core::Result<()> {
        log::info!("Checking QuickBooks SDK installation status");
        unsafe {
            // Check Program Files paths
            let program_files_paths = [
                r"C:\Program Files\Common Files\Intuit\QuickBooks\QBXMLRPSync.dll",
                r"C:\Program Files\Common Files\Intuit\QuickBooks\QBXMLRP2.dll",
                r"C:\Program Files (x86)\Common Files\Intuit\QuickBooks\QBXMLRPSync.dll",
                r"C:\Program Files (x86)\Common Files\Intuit\QuickBooks\QBXMLRP2.dll",
            ];

            for path in program_files_paths.iter() {
                if Path::new(path).exists() {
                    log::debug!("Found SDK component: {}", path);
                } else {
                    log::warn!("Missing SDK component: {}", path);
                }
            }

            // Check if SDK is registered in Add/Remove Programs
            let mut hklm = HKEY::default();
            let _ = RegOpenKeyExA(
                HKEY_LOCAL_MACHINE,
                PCSTR::from_raw(b"SOFTWARE/Microsoft/Windows/CurrentVersion/Uninstall/QuickBooks SDK v15.0\0".as_ptr()),
                0,
                KEY_READ,
                &mut hklm
            ).map(|_| {
                log::debug!("Found QuickBooks SDK in Add/Remove Programs");
            }).map_err(|e| {
                log::warn!("QuickBooks SDK not found in Add/Remove Programs: 0x{:08X}", e.code().0);
            });
        }
        Ok(())
    }

    fn check_registry_paths() -> windows::core::Result<()> {
        log::info!("Starting detailed COM registration check");
        unsafe {
            let paths = [
                r"SOFTWARE\Classes\QBXMLRP2.RequestProcessor.2",
                r"SOFTWARE\Classes\WOW6432Node\QBXMLRP2.RequestProcessor.2",
                r"SOFTWARE\Classes\CLSID\{62989BF0-0AA7-11D4-8754-00A0C9AC7AC3}",
                r"SOFTWARE\Classes\WOW6432Node\CLSID\{62989BF0-0AA7-11D4-8754-00A0C9AC7AC3}",
            ];

            let mut hklm = HKEY::default();
            RegOpenKeyExA(
                HKEY_LOCAL_MACHINE,
                PCSTR::null(),
                0,
                KEY_READ,
                &mut hklm
            )?;

            for path in paths.iter() {
                log::debug!("Checking registry path: {}", path);
                let mut key = HKEY::default();
                match RegOpenKeyExA(
                    hklm,
                    PCSTR::from_raw(path.as_ptr() as *const u8),
                    0,
                    KEY_READ,
                    &mut key
                ) {
                    Ok(_) => {
                        let mut found_any_values = false;
                        log::debug!("Found registry key: {}", path);
                        let mut buf = [0u8; 260];
                        let mut size = buf.len() as u32;
                        // Check InprocServer32 key
                        if RegQueryValueExA(
                            key,
                            PCSTR::from_raw(b"InprocServer32\0".as_ptr()),
                            None,
                            None,
                            Some(buf.as_mut_ptr() as *mut u8),
                            Some(&mut size),
                        ).is_ok() {
                            if let Ok(path) = String::from_utf8(buf[..size as usize].to_vec()) {
                                log::debug!("DLL path: {}", path);
                                if let Some(path) = Path::new(&path).parent() {
                                    log::debug!("DLL directory exists: {}", path.exists());
                                }
                                found_any_values = true;
                            }
                        }

                        // Check additional values
                        let values_to_check = [
                            "LocalServer32",
                            "ProgID",
                            "VersionIndependentProgID",
                            "Implemented Categories",
                        ];

                        for value_name in values_to_check.iter() {
                            let mut buf = [0u8; 260];
                            let mut size = buf.len() as u32;
                            if RegQueryValueExA(
                                key,
                                PCSTR::from_raw(format!("{value_name}\0").as_ptr() as *const u8),
                                None,
                                None,
                                Some(buf.as_mut_ptr() as *mut u8),
                                Some(&mut size),
                            ).is_ok() {
                                if let Ok(value) = String::from_utf8(buf[..size as usize].to_vec()) {
                                    log::debug!("Found {} = {}", value_name, value.trim_end_matches('\0'));
                                    found_any_values = true;
                                }
                            }
                        }

                        if !found_any_values {
                            log::warn!("Registry key exists but has no relevant values: {}", path);
                        }

                        let _ = RegCloseKey(key);
                    },
                    Err(e) => {
                        log::warn!("Registry key not found: {} (error: 0x{:08X})", path, e.code().0);
                    }
                }
            }
            let _ = RegCloseKey(hklm);
        }
        Ok(())
    }

    pub fn new() -> windows::core::Result<Self> {
        Self::check_sdk_installation()?;
        Self::check_registry_paths()?;

        let prog_id = HSTRING::from("QBXMLRP2.RequestProcessor.2");
        log::debug!("Attempting to get CLSID for ProgID: {}", prog_id);
        let clsid = unsafe { CLSIDFromProgID(&prog_id)? };
        log::debug!("Got CLSID: {:?}", clsid);
        // Try different server types
        let server_types = [
            (CLSCTX_LOCAL_SERVER, "CLSCTX_LOCAL_SERVER"),
            (CLSCTX_INPROC_SERVER, "CLSCTX_INPROC_SERVER"),
            (CLSCTX_ALL, "CLSCTX_ALL"),
        ];

        let mut last_error = None;
        let mut dispatch = None;

        for (server_type, type_name) in server_types.iter() {
            log::debug!("Attempting to create COM instance with {}", type_name);
            match unsafe {
                CoCreateInstance::<Option<&IUnknown>, IDispatch>(
                    &clsid,
                    None,
                    *server_type
                )
            } {
                Ok(disp) => {
                    log::debug!("Successfully created COM instance with {}", type_name);
                    dispatch = Some(disp);
                    break;
                }
                Err(e) => {
                    log::error!("CoCreateInstance with {} failed with error code: 0x{:08X}", type_name, e.code().0);
                    last_error = Some(e);
                }
            }
        }

        let dispatch = dispatch.ok_or_else(|| last_error.unwrap())?;
        log::debug!("Successfully created COM instance");

        // Get all method IDs upfront
        let open_connection_id = Self::get_method_id(&dispatch, "OpenConnection")?;
        let begin_session_id = Self::get_method_id(&dispatch, "BeginSession")?;
        let end_session_id = Self::get_method_id(&dispatch, "EndSession")?;
        let close_connection_id = Self::get_method_id(&dispatch, "CloseConnection")?;
        let process_request_id = Self::get_method_id(&dispatch, "ProcessRequest")?;

        Ok(Self {
            inner: dispatch,
            open_connection_id,
            begin_session_id,
            end_session_id,
            close_connection_id,
            process_request_id,
        })
    }

    fn get_method_id(dispatch: &IDispatch, name: &str) -> windows::core::Result<i32> {
        let mut dispid = -1i32;
        let method_name = HSTRING::from(name);
        let names = [PCWSTR::from_raw(method_name.as_ptr())];
        
        unsafe {
            dispatch.GetIDsOfNames(
                &windows::core::GUID::zeroed(),
                names.as_ptr(),
                1,
                0x0409, // LCID_ENGLISH_US
                &mut dispid,
            )?;
        }
        Ok(dispid)
    }

    pub fn open_connection(&self, app_id: &str, app_name: &str) -> windows::core::Result<()> {
        let app_id_var = create_bstr_variant(app_id);
        let app_name_var = create_bstr_variant(app_name);
        let args = [app_id_var, app_name_var];

        let mut params = create_dispparams(&args);
        
        unsafe {
            self.inner.Invoke(
                self.open_connection_id,
                &windows::core::GUID::zeroed(),
                0x0409,
                DISPATCH_METHOD,
                &mut params,
                None,
                None,
                None,
            )
        }
    }

    pub fn begin_session(&self, company_file: &str, file_mode: FileMode) -> windows::core::Result<String> {
        let file_var = create_bstr_variant(company_file);
        let mode_var = create_bstr_variant(match file_mode {
            FileMode::SingleUser => "qbFileOpenSingleUser",
            FileMode::MultiUser => "qbFileOpenMultiUser",
            FileMode::DoNotCare => "qbFileOpenDoNotCare",
        });
        let args = [file_var, mode_var];

        let mut params = create_dispparams(&args);
        let mut result = VARIANT::default();
        
        unsafe {
            self.inner.Invoke(
                self.begin_session_id,
                &windows::core::GUID::zeroed(),
                0x0409,
                DISPATCH_METHOD,
                &mut params,
                Some(&mut result),
                None,
                None,
            )?;
            
            variant_to_string(&result)
        }
    }

    pub fn end_session(&self, ticket: &str) -> windows::core::Result<()> {
        let ticket_var = create_bstr_variant(ticket);
        let args = [ticket_var];
        let mut params = create_dispparams(&args);
        
        unsafe {
            self.inner.Invoke(
                self.end_session_id,
                &windows::core::GUID::zeroed(),
                0x0409,
                DISPATCH_METHOD,
                &mut params,
                None,
                None,
                None,
            )
        }
    }

    pub fn close_connection(&self) -> windows::core::Result<()> {
        let mut params = create_empty_dispparams();
        
        unsafe {
            self.inner.Invoke(
                self.close_connection_id,
                &windows::core::GUID::zeroed(),
                0x0409,
                DISPATCH_METHOD,
                &mut params,
                None,
                None,
                None,
            )
        }
    }

    pub fn process_request(&self, ticket: &str, request: &str) -> windows::core::Result<String> {
        let ticket_var = create_bstr_variant(ticket);
        let request_var = create_bstr_variant(request);
        let args = [ticket_var, request_var];
        let mut params = create_dispparams(&args);
        let mut result = VARIANT::default();
        
        unsafe {
            self.inner.Invoke(
                self.process_request_id,
                &windows::core::GUID::zeroed(),
                0x0409,
                DISPATCH_METHOD,
                &mut params,
                Some(&mut result),
                None,
                None,
            )?;
            
            variant_to_string(&result)
        }
    }

    pub fn get_current_company_file_name(&self, ticket: &str) -> windows::core::Result<String> {
        let mut params = create_dispparams(&[create_bstr_variant(ticket)]);
        let mut result = VARIANT::default();
        let mut exc_info = EXCEPINFO::default();
        let mut arg_err = 0;
        
        unsafe {
            self.inner.Invoke(
                3,  // DISPID for GetCurrentCompanyFileName
                &windows::core::GUID::zeroed(),
                0x0409,
                DISPATCH_METHOD,
                &mut params,
                Some(&mut result),
                Some(&mut exc_info),
                Some(&mut arg_err),
            )?;
            
            variant_to_string(&result)
        }
    }
}

impl Drop for RequestProcessor2 {
    fn drop(&mut self) {
        // The IDispatch will be automatically dropped, which releases the COM object
    }
}
