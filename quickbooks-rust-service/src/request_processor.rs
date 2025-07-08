use std::path::Path;
use std::ffi::CString;
use std::ptr;
use windows::core::{GUID, HSTRING, PCWSTR, PCSTR, PSTR};
use windows::Win32::System::Registry::{HKEY, HKEY_LOCAL_MACHINE, HKEY_CURRENT_USER};
use windows::Win32::System::Com::{CLSIDFromProgID, CoCreateInstance, CLSCTX_LOCAL_SERVER, IDispatch, EXCEPINFO, DISPATCH_METHOD, DISPATCH_FLAGS};

#[derive(Debug)]
enum QBXMLRPConnectionType {
    Unknown = 0,
    LocalQBD = 1,
    LocalQBDLaunchUI = 2,
    RemoteQBD = 3,
    RemoteQBOE = 4,
}
use windows::Win32::System::Registry::{RegOpenKeyExA, RegQueryValueExA, RegEnumKeyExA, RegCloseKey, KEY_READ};
use windows::Win32::System::Variant::VARIANT;
use crate::com_helpers::{create_bstr_variant, create_dispparams, create_empty_dispparams, variant_to_string};
use crate::FileMode;

/// Type-safe wrapper for QBXMLRP2 RequestProcessor2
pub struct RequestProcessor2 {
    inner: IDispatch,
    // Cache method IDs after first lookup
    open_connection_id: i32,
    open_connection2_id: i32,
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

            // Check SDK 16.0 installation paths
            // Try to find actual SDK registration
            let base_paths = [
                r"SOFTWARE\Intuit",
                r"SOFTWARE\WOW6432Node\Intuit",
            ];

            for base_path in base_paths.iter() {
                let mut key = HKEY::default();
                if RegOpenKeyExA(
                    HKEY_LOCAL_MACHINE,
                    PCSTR::from_raw(base_path.as_bytes().as_ptr()),
                    0,
                    KEY_READ,
                    &mut key
                ).is_ok() {
                    log::info!("Scanning for SDK keys in: {}", base_path);
                    let mut index = 0u32;
                    let mut name_buf = [0u8; 260];
                    let mut name_size = name_buf.len() as u32;

                    while RegEnumKeyExA(
                        key,
                        index,
                        PSTR(name_buf.as_mut_ptr() as *mut u8),
                        &mut name_size,
                        None,
                        PSTR(ptr::null_mut()),
                        None,
                        None,
                    ).is_ok() {
                        if let Ok(subkey_name) = String::from_utf8(name_buf[..name_size as usize].to_vec()) {
                            if subkey_name.contains("QBSDK") || subkey_name.contains("QuickBooks") {
                                log::info!("Found SDK-related key: {}", subkey_name);
                            }
                        }
                        index += 1;
                        name_size = name_buf.len() as u32;
                    }
                    let _ = RegCloseKey(key);
                }
            }

            let sdk_paths = [
                r"SOFTWARE\Intuit\QBSDK16.0",
                r"SOFTWARE\WOW6432Node\Intuit\QBSDK16.0",
                r"SOFTWARE\Intuit\QuickBooks\QBSDKComponentRegister",
                r"SOFTWARE\WOW6432Node\Intuit\QuickBooks\QBSDKComponentRegister",
            ];

            for path in sdk_paths.iter() {
                let mut key = HKEY::default();
                match RegOpenKeyExA(
                    HKEY_LOCAL_MACHINE,
                    PCSTR::from_raw(path.as_bytes().as_ptr()),
                    0,
                    KEY_READ,
                    &mut key
                ) {
                    Ok(_) => {
                        log::info!("Found SDK path: {}", path);
                        let mut buf = [0u8; 260];
                        let mut size = buf.len() as u32;
                        
                        // Check for installation path
                        if RegQueryValueExA(
                            key,
                            PCSTR::from_raw(b"InstallPath\0".as_ptr()),
                            None,
                            None,
                            Some(buf.as_mut_ptr() as *mut u8),
                            Some(&mut size),
                        ).is_ok() {
                            if let Ok(install_path) = String::from_utf8(buf[..size as usize].to_vec()) {
                                log::info!("SDK Install Path: {}", install_path.trim_end_matches('\0'));
                            }
                        }
                        
                        // Check for version info
                        if RegQueryValueExA(
                            key,
                            PCSTR::from_raw(b"Version\0".as_ptr()),
                            None,
                            None,
                            Some(buf.as_mut_ptr() as *mut u8),
                            Some(&mut size),
                        ).is_ok() {
                            if let Ok(version) = String::from_utf8(buf[..size as usize].to_vec()) {
                                log::info!("SDK Version: {}", version.trim_end_matches('\0'));
                            }
                        }
                        
                        let _ = RegCloseKey(key);
                    },
                    Err(e) => {
                        log::warn!("SDK path not found: {} (error: 0x{:08X})", path, e.code().0);
                    }
                }
            }

            // Check if SDK is registered in Add/Remove Programs
            let mut hklm = HKEY::default();
            let _ = RegOpenKeyExA(
                HKEY_LOCAL_MACHINE,
                PCSTR::from_raw(b"SOFTWARE/Microsoft/Windows/CurrentVersion/Uninstall/qbsdk16.0(x64)\0".as_ptr()),
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
            log::info!("Checking HKEY_CURRENT_USER registry paths");
            let hkcu_paths = [
                r"SOFTWARE\Classes\QBXMLRP2.RequestProcessor.2",
                r"SOFTWARE\Classes\QBXMLRP2.RequestProcessor2",
                r"SOFTWARE\Classes\CLSID\{62989BF0-0AA7-11D4-8754-00A0C9AC7AC3}",
                r"SOFTWARE\Classes\CLSID\{71F531F5-8E67-4A7A-9161-15733FFC3206}",
            ];

            let mut hkcu = HKEY::default();
            RegOpenKeyExA(
                HKEY_CURRENT_USER,
                PCSTR::null(),
                0,
                KEY_READ,
                &mut hkcu
            )?;

            for path in hkcu_paths.iter() {
                let mut key = HKEY::default();
                match RegOpenKeyExA(
                    hkcu,
                    PCSTR::from_raw(path.as_bytes().as_ptr()),
                    0,
                    KEY_READ,
                    &mut key
                ) {
                    Ok(_) => {
                        log::debug!("Found registry key (HKCU): {}", path);
                        let _ = RegCloseKey(key);
                    },
                    Err(e) => {
                        log::warn!("Registry key not found (HKCU): {} (error: 0x{:08X})", path, e.code().0);
                    }
                }
            }
            let _ = RegCloseKey(hkcu);

            log::info!("Checking HKEY_LOCAL_MACHINE registry paths");
            let paths = [
                // Check both naming variations
                r"SOFTWARE\Classes\QBXMLRP2.RequestProcessor.2",
                r"SOFTWARE\Classes\QBXMLRP2.RequestProcessor2",
                r"SOFTWARE\Classes\WOW6432Node\QBXMLRP2.RequestProcessor.2",
                r"SOFTWARE\Classes\WOW6432Node\QBXMLRP2.RequestProcessor2",
                // Check both CLSIDs we've seen
                r"SOFTWARE\Classes\CLSID\{62989BF0-0AA7-11D4-8754-00A0C9AC7AC3}",
                r"SOFTWARE\Classes\CLSID\{71F531F5-8E67-4A7A-9161-15733FFC3206}",
                r"SOFTWARE\Classes\WOW6432Node\CLSID\{62989BF0-0AA7-11D4-8754-00A0C9AC7AC3}",
                r"SOFTWARE\Classes\WOW6432Node\CLSID\{71F531F5-8E67-4A7A-9161-15733FFC3206}",
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
                    PCSTR::from_raw(path.as_bytes().as_ptr()),
                    0,
                    KEY_READ,
                    &mut key
                ) {
                    Ok(_) => {
                        log::debug!("Found registry key (HKLM): {}", path);
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
                            PCSTR::from_raw(CString::new(format!("{value_name}")).unwrap().as_bytes_with_nul().as_ptr()),
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

// Try known CLSIDs
// Create RequestProcessor2 COM object directly from ProgID
let prog_id = HSTRING::from("QBXMLRP2.RequestProcessor.2");
let clsid = unsafe { CLSIDFromProgID(&prog_id)? };

// Create the COM object with specific server type
let dispatch: IDispatch = unsafe {
    CoCreateInstance(
        &clsid,
        None,
        CLSCTX_LOCAL_SERVER  // Use local server like SDKTestPlus3
    )?
};
        log::debug!("Successfully created COM instance");

        // Get all method IDs upfront
        let open_connection_id = Self::get_method_id(&dispatch, "OpenConnection")?;
        let open_connection2_id = Self::get_method_id(&dispatch, "OpenConnection2")?;
        let begin_session_id = Self::get_method_id(&dispatch, "BeginSession")?;
        let end_session_id = Self::get_method_id(&dispatch, "EndSession")?;
        let close_connection_id = Self::get_method_id(&dispatch, "CloseConnection")?;
        let process_request_id = Self::get_method_id(&dispatch, "ProcessRequest")?;

        Ok(Self {
            inner: dispatch,
            open_connection_id,
            open_connection2_id,
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
            )?
        }
        Ok(dispid)
    }

pub fn open_connection2(&self, app_id: &str, app_name: &str, conn_type: QBXMLRPConnectionType) -> windows::core::Result<()> {
        let app_id_var = create_bstr_variant(app_id);
        let app_name_var = create_bstr_variant(app_name);
        let conn_type_var = VARIANT::default(); // TODO: Create integer variant
        let args = [app_id_var, app_name_var, conn_type_var];

        let mut params = create_dispparams(&args);
        let mut result = VARIANT::default();
        let mut exc_info = EXCEPINFO::default();
        let mut arg_err = 0u32;
        
        unsafe {
            self.inner.Invoke(
                self.open_connection2_id,
                &GUID::zeroed(),
                0x0409,  // LOCALE_SYSTEM_DEFAULT
                DISPATCH_FLAGS(DISPATCH_METHOD.0 as u16),
                &mut params,
                Some(&mut result),
                Some(&mut exc_info),
                Some(&mut arg_err)
            )
        }
    }

pub fn open_connection(&self, app_id: &str, app_name: &str) -> windows::core::Result<()> {
        // Set auth preferences first
        unsafe {
            let mut auth_prefs_result = VARIANT::default();
            let _ = self.inner.Invoke(
                Self::get_method_id(&self.inner, "AuthPreferences")?,
                &GUID::zeroed(),
                0x0409,  // LOCALE_SYSTEM_DEFAULT
                DISPATCH_FLAGS(DISPATCH_METHOD.0 as u16),
                &mut create_empty_dispparams(),
                Some(&mut auth_prefs_result),
                None,
                None,
            )?;

            if auth_prefs_result.Anonymous.Anonymous.vt == VT_DISPATCH {
                let auth_prefs: IDispatch = (*auth_prefs_result.Anonymous.Anonymous.Anonymous.pdispVal).clone();
                
                // Set enterprise flag (0x8)
                let auth_flags = create_variant_i4(0x8);
                let mut params = create_dispparams(&[auth_flags]);
                
                let _ = auth_prefs.Invoke(
                    Self::get_method_id(&auth_prefs, "PutAuthFlags")?,
                    &GUID::zeroed(),
                    0x0409,  // LOCALE_SYSTEM_DEFAULT
                    DISPATCH_FLAGS(DISPATCH_METHOD.0 as u16),
                    &mut params,
                    None,
                    None,
                    None,
                )?;
            }
        }
        let app_id_var = create_bstr_variant(app_id);
        let app_name_var = create_bstr_variant(app_name);
        let args = [app_id_var, app_name_var];

        let mut params = create_dispparams(&args);
        let mut result = VARIANT::default();
        let mut exc_info = EXCEPINFO::default();
        let mut arg_err = 0u32;
        
        unsafe {
            self.inner.Invoke(
                self.open_connection_id,
                &GUID::zeroed(),
                0x0409,  // LOCALE_SYSTEM_DEFAULT
                DISPATCH_FLAGS(DISPATCH_METHOD.0 as u16),
                &mut params,
                Some(&mut result),
                Some(&mut exc_info),
                Some(&mut arg_err)
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
        let mut exc_info = EXCEPINFO::default();
        let mut arg_err = 0u32;

        unsafe {
            self.inner.Invoke(
                self.begin_session_id,
                &GUID::zeroed(),
                0x0409,  // LOCALE_SYSTEM_DEFAULT
                DISPATCH_FLAGS(DISPATCH_METHOD.0 as u16),
                &mut params,
                Some(&mut result),
                Some(&mut exc_info),
                Some(&mut arg_err)
            )?;
            
            variant_to_string(&result)
        }
    }

pub fn end_session(&self, ticket: &str) -> windows::core::Result<()> {
        let ticket_var = create_bstr_variant(ticket);
        let args = [ticket_var];
        let mut params = create_dispparams(&args);
        let mut result = VARIANT::default();
        let mut exc_info = EXCEPINFO::default();
        let mut arg_err = 0u32;

        unsafe {
            self.inner.Invoke(
                self.end_session_id,
                &GUID::zeroed(),
                0x0409,  // LOCALE_SYSTEM_DEFAULT
                DISPATCH_FLAGS(DISPATCH_METHOD.0 as u16),
                &mut params,
                Some(&mut result),
                Some(&mut exc_info),
                Some(&mut arg_err)
            )
        }
    }

pub fn close_connection(&self) -> windows::core::Result<()> {
        let mut params = create_empty_dispparams();
        let mut result = VARIANT::default();
        let mut exc_info = EXCEPINFO::default();
        let mut arg_err = 0u32;
        
        unsafe {
            self.inner.Invoke(
                self.close_connection_id,
                &GUID::zeroed(),
                0x0409,  // LOCALE_SYSTEM_DEFAULT
                DISPATCH_FLAGS(DISPATCH_METHOD.0 as u16),
                &mut params,
                Some(&mut result),
                Some(&mut exc_info),
                Some(&mut arg_err)
            )
        }
    }

pub fn process_request(&self, ticket: &str, request: &str) -> windows::core::Result<String> {
        let ticket_var = create_bstr_variant(ticket);
        let request_var = create_bstr_variant(request);
        let args = [ticket_var, request_var];
        let mut params = create_dispparams(&args);
        let mut result = VARIANT::default();
        let mut exc_info = EXCEPINFO::default();
        let mut arg_err = 0u32;

        unsafe {
            self.inner.Invoke(
                self.process_request_id,
                &GUID::zeroed(),
                0x0409,  // LOCALE_SYSTEM_DEFAULT
                DISPATCH_FLAGS(DISPATCH_METHOD.0 as u16),
                &mut params,
                Some(&mut result),
                Some(&mut exc_info),
                Some(&mut arg_err)
            )?;
            
            variant_to_string(&result)
        }
    }

pub fn get_current_company_file_name(&self, ticket: &str) -> windows::core::Result<String> {
        let ticket_var = create_bstr_variant(ticket);
        let args = [ticket_var];
        let mut params = create_dispparams(&args);
        let mut result = VARIANT::default();
        let mut exc_info = EXCEPINFO::default();
        let mut arg_err = 0u32;

        unsafe {
            self.inner.Invoke(
                3,  // DISPID for GetCurrentCompanyFileName
                &GUID::zeroed(),
                0x0409,  // LOCALE_SYSTEM_DEFAULT
                DISPATCH_FLAGS(DISPATCH_METHOD.0 as u16),
                &mut params,
                Some(&mut result),
                Some(&mut exc_info),
                Some(&mut arg_err)
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
