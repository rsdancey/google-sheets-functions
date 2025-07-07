use anyhow::Result;
use log::{info, error};

#[cfg(windows)]
use windows::{
    core::{HSTRING, VARIANT, BSTR, PCWSTR},
    Win32::System::Com::{
        CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_INPROC_SERVER,
        COINIT_APARTMENTTHREADED, IDispatch, CLSIDFromProgID,
        DISPATCH_METHOD, DISPPARAMS, EXCEPINFO, DISPATCH_FLAGS
    },
};

use crate::config::QuickBooksConfig;

#[derive(Debug)]
pub struct QuickBooksClient {
    config: QuickBooksConfig,
}

#[cfg(windows)]
struct ComGuard;

#[cfg(windows)]
impl ComGuard {
    fn new() -> Result<Self> {
        unsafe {
            let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
            match hr.0 {
                0 => { // S_OK
                    info!("COM initialized successfully");
                    Ok(ComGuard)
                }
                1 => { // S_FALSE - already initialized with same mode
                    info!("COM already initialized in compatible mode");
                    Ok(ComGuard)
                }
                code if code == 0x80010106u32 as i32 => { // RPC_E_CHANGED_MODE
                    info!("COM already initialized in different mode");
                    Ok(ComGuard)
                }
                code => {
                    error!("Failed to initialize COM: HRESULT 0x{:08x}", code as u32);
                    Err(anyhow::anyhow!("COM initialization failed: HRESULT 0x{:08x}", code as u32))
                }
            }
        }
    }
}

#[cfg(windows)]
impl Drop for ComGuard {
    fn drop(&mut self) {
        unsafe {
            CoUninitialize();
            info!("COM uninitialized");
        }
    }
}

impl QuickBooksClient {
    pub fn new(config: &QuickBooksConfig) -> Result<Self> {
        Ok(Self {
            config: config.clone(),
        })
    }

    /// Test basic QuickBooks SDK availability
    pub async fn test_connection(&self) -> Result<()> {
        #[cfg(windows)]
        {
            self.test_sdk_availability().await
        }
        
        #[cfg(not(windows))]
        {
            Err(anyhow::anyhow!("QuickBooks Desktop is only supported on Windows"))
        }
    }

    /// Attempt to register with QuickBooks and establish a connection
    pub async fn register_with_quickbooks(&self) -> Result<()> {
        #[cfg(windows)]
        {
            self.perform_registration().await
        }
        
        #[cfg(not(windows))]
        {
            Err(anyhow::anyhow!("QuickBooks Desktop is only supported on Windows"))
        }
    }

    #[cfg(windows)]
    async fn test_sdk_availability(&self) -> Result<()> {
        info!("Testing QuickBooks SDK availability...");
        
        // Initialize COM with guard
        let _com = ComGuard::new()?;

        let result = self.create_session_manager();

        // COM will be uninitialized when _com guard is dropped
        match result {
            Ok(_) => {
                info!("✅ QuickBooks SDK is available");
                Ok(())
            }
            Err(e) => {
                error!("❌ QuickBooks SDK not available: {}", e);
                Err(e)
            }
        }
    }

    #[cfg(windows)]
    async fn perform_registration(&self) -> Result<()> {
        info!("Starting QuickBooks registration process...");
        
        // Initialize COM with guard
        let _com = ComGuard::new()?;

        info!("About to call register_application...");
        let result = self.register_application();
        info!("register_application returned: {:?}", result.is_ok());

        // COM will be uninitialized when _com guard is dropped
        result
    }

    #[cfg(windows)]
    fn create_session_manager(&self) -> Result<IDispatch> {
        // Try QuickBooks session manager classes in order of preference
        let session_manager_classes = [
            "QBFC16.QBSessionManager", // QB Enterprise v24+
            "QBFC15.QBSessionManager", // QB Enterprise v23
            "QBFC14.QBSessionManager", // QB Enterprise v22
            "QBFC13.QBSessionManager", // QB Enterprise v21
        ];
        
        for class_name in &session_manager_classes {
            info!("Trying to create session manager: {}", class_name);
            
            let prog_id = HSTRING::from(*class_name);
            
            if let Ok(clsid) = unsafe { CLSIDFromProgID(&prog_id) } {
                if let Ok(session_manager) = unsafe { 
                    CoCreateInstance(&clsid, None, CLSCTX_INPROC_SERVER) 
                } {
                    info!("✅ Successfully created session manager: {}", class_name);
                    
                    // Skip preliminary functionality test - will test during actual connection
                    info!("✅ Session manager created successfully, will test during connection");
                    return Ok(session_manager);
                }
            }
        }
        
        Err(anyhow::anyhow!("Failed to create QuickBooks session manager. Ensure QuickBooks Desktop is installed."))
    }

    #[cfg(windows)]
    fn register_application(&self) -> Result<()> {
        info!("Creating QuickBooks session manager...");
        
        let session_manager = match self.create_session_manager() {
            Ok(sm) => {
                info!("Session manager created successfully");
                sm
            }
            Err(e) => {
                error!("Failed to create session manager: {}", e);
                return Err(e);
            }
        };
        
        info!("Attempting to register application with QuickBooks...");
        
        // Application details for registration
        let app_id = "QuickBooks-Sheets-Sync-v1";
        let app_name = "QuickBooks Sheets Sync";
        
        info!("About to call perform_complete_session_test...");
        
        // Follow the C++ SDK pattern: OpenConnection2 -> BeginSession -> EndSession -> CloseConnection
        match self.perform_complete_session_test(&session_manager, app_id, app_name) {
            Ok(()) => {
                info!("✅ Successfully registered with QuickBooks!");
                Ok(())
            }
            Err(e) => {
                error!("❌ Registration failed: {}", e);
                info!("💡 Error details: {:?}", e);
                
                // Try to get more specific error information
                if let Some(source) = e.source() {
                    error!("Error source: {}", source);
                }
                
                Err(e)
            }
        }
    }

    #[cfg(windows)]
    fn perform_complete_session_test(&self, session_manager: &IDispatch, app_id: &str, app_name: &str) -> Result<()> {
        info!("Starting complete session test following C++ SDK pattern...");
        
        // Step 1: OpenConnection
        info!("Step 1: Attempting OpenConnection...");
        match self.try_open_connection(session_manager, app_id, app_name) {
            Ok(()) => {
                info!("✅ OpenConnection successful");
            }
            Err(e) => {
                error!("❌ OpenConnection failed: {}", e);
                return Err(anyhow::anyhow!("OpenConnection failed: {}", e));
            }
        }
        
        // Step 2: BeginSession (with current company file)
        info!("Step 2: Attempting BeginSession...");
        let ticket = match self.begin_session(session_manager, "") {
            Ok(t) => {
                info!("✅ BeginSession successful, ticket: {:?}", t);
                t
            }
            Err(e) => {
                error!("❌ BeginSession failed: {}", e);
                return Err(anyhow::anyhow!("BeginSession failed: {}", e));
            }
        };
        
        // Step 3: Test getting company file name
        info!("Step 3: Testing GetCurrentCompanyFileName...");
        match self.get_current_company_file_name(session_manager, &ticket) {
            Ok(company_name) => {
                info!("✅ Current company file: {}", company_name);
            }
            Err(e) => {
                info!("⚠️ Could not get company file name: {}", e);
            }
        }
        
        // Step 4: Test version information
        info!("Step 4: Testing version information...");
        match self.test_version_info(session_manager, &ticket) {
            Ok(_) => {
                info!("✅ Version information test successful");
            }
            Err(e) => {
                info!("⚠️ Version information test failed: {}", e);
            }
        }
        
        // Step 5: EndSession
        info!("Step 5: Attempting EndSession...");
        match self.end_session(session_manager, &ticket) {
            Ok(()) => {
                info!("✅ EndSession successful");
            }
            Err(e) => {
                error!("❌ EndSession failed: {}", e);
                // Don't return error here - try to close connection anyway
            }
        }
        
        // Step 6: CloseConnection
        info!("Step 6: Attempting CloseConnection...");
        match self.close_connection(session_manager) {
            Ok(()) => {
                info!("✅ CloseConnection successful");
            }
            Err(e) => {
                error!("❌ CloseConnection failed: {}", e);
                // Don't return error here - connection might already be closed
            }
        }
        
        Ok(())
    }

    #[cfg(windows)]
    fn try_open_connection(&self, session_manager: &IDispatch, app_id: &str, app_name: &str) -> Result<()> {
        info!("Opening Connection to QuickBooks");
        info!("  AppID: '{}'", app_id);
        info!("  AppName: '{}'", app_name);
        
        let app_id_variant = self.create_string_variant(app_id)?;
        let app_name_variant = self.create_string_variant(app_name)?;
        let conn_type_variant = self.create_int_variant(1)?; // 1 = localQBD mode
        
        // Use OpenConnection2 with 3 parameters: app_id, app_name, and connection_type=1 (localQBD)
        info!("💡 If this is your first time, QuickBooks will show an authorization dialog");
        info!("Attempting to open connection with localQBD mode");
        match self.invoke_method(
            session_manager,
            "OpenConnection2",
            &[&app_id_variant, &app_name_variant, &conn_type_variant]
        ) {
            Ok(_) => {
                info!("✅ Successfully opened connection to QuickBooks");
                Ok(())
            }
            Err(e) => {
                error!("❌ Failed to connect to QuickBooks");
                error!("Make sure:");
                error!("  1. QuickBooks is running and a company file is open");
                error!("  2. No other applications are currently using QuickBooks");
                error!("  3. You have administrator privileges");
                error!("Error details: {}", e);
                Err(anyhow::anyhow!("Failed to open QuickBooks connection: {}", e))
            }
        }
    }

    #[cfg(windows)]
    fn begin_session(&self, session_manager: &IDispatch, company_file: &str) -> Result<String> {
        // First try to get current company file if no file specified
        let effective_company_file = if company_file.is_empty() {
            // Try to get the company file name through the Company property
            match self.get_property(session_manager, "Company") {
                Ok(company) => {
                    info!("Got company object");
                    let company_dispatch = self.variant_to_dispatch(&company)?;
                    match self.get_property(&company_dispatch, "FileName") {
                        Ok(current_file) => {
                            match self.variant_to_string(&current_file) {
                                Ok(file_path) => {
                                    info!("Found currently open company file: {}", file_path);
                                    file_path
                                }
                                Err(e) => {
                                    info!("Failed to convert company file path to string: {}", e);
                                    return Err(anyhow::anyhow!("Failed to get current company file path: {}", e));
                                }
                            }
                        }
                        Err(e) => {
                            info!("Failed to get company file name: {}", e);
                            return Err(anyhow::anyhow!("Failed to get company file name: {}", e));
                        }
                    }
                }
                Err(e) => {
                    info!("Failed to get Company object: {}", e);
                    return Err(anyhow::anyhow!("Failed to get Company object: {}", e));
                }
            }
        } else {
            company_file.to_string()
        };

        info!("Starting BeginSession with company file: '{}'", effective_company_file);
        
        let company_file_variant = self.create_string_variant(&effective_company_file)?;
        
        // Try different file open modes for multi-user QuickBooks
        let file_open_modes = [
            (1, "MultiUser"),      // 1 = MultiUser mode
            (0, "DoNotCare"),      // 0 = DoNotCare 
            (2, "SingleUser"),     // 2 = SingleUser (might still work in multi-user)
        ];
        
        for &(mode, mode_name) in &file_open_modes {
            info!("  Trying BeginSession with mode: {} ({})", mode, mode_name);
            
            let file_open_mode = self.create_int_variant(mode)?;
            
            // BeginSession returns a ticket (session ID)
            match self.invoke_method(
                session_manager, 
                "BeginSession", 
                &[&company_file_variant, &file_open_mode]
            ) {
                Ok(result) => {
                    let ticket = self.variant_to_string(&result)?;
                    info!("✅ BeginSession successful with mode {} ({}), ticket: {}", mode, mode_name, ticket);
                    return Ok(ticket);
                }
                Err(e) => {
                    info!("  ❌ BeginSession failed with mode {} ({}): {}", mode, mode_name, e);
                    continue;
                }
            }
        }
        
        Err(anyhow::anyhow!("BeginSession failed with all file open modes"))
    }

    #[cfg(windows)]
    fn end_session(&self, session_manager: &IDispatch, ticket: &str) -> Result<()> {
        info!("Ending session with ticket: {}", ticket);
        
        let ticket_variant = self.create_string_variant(ticket)?;
        
        self.invoke_method(session_manager, "EndSession", &[&ticket_variant])
            .map(|_| ())
    }

    #[cfg(windows)]
    fn get_current_company_file_name(&self, session_manager: &IDispatch, ticket: &str) -> Result<String> {
        let ticket_variant = self.create_string_variant(ticket)?;
        
        let result = self.invoke_method(
            session_manager, 
            "GetCurrentCompanyFileName", 
            &[&ticket_variant]
        )?;
        
        self.variant_to_string(&result)
    }

    #[cfg(windows)]
    fn test_version_info(&self, session_manager: &IDispatch, ticket: &str) -> Result<()> {
        info!("Testing version information...");
        
        // Test basic version properties (similar to C++ SDK)
        if let Ok(result) = self.get_property(session_manager, "MajorVersion") {
            if let Ok(major) = self.variant_to_string(&result) {
                info!("  Major Version: {}", major);
            }
        }
        
        if let Ok(result) = self.get_property(session_manager, "MinorVersion") {
            if let Ok(minor) = self.variant_to_string(&result) {
                info!("  Minor Version: {}", minor);
            }
        }
        
        if let Ok(result) = self.get_property(session_manager, "ReleaseLevel") {
            if let Ok(release) = self.variant_to_string(&result) {
                info!("  Release Level: {}", release);
            }
        }
        
        if let Ok(result) = self.get_property(session_manager, "ReleaseNumber") {
            if let Ok(release_num) = self.variant_to_string(&result) {
                info!("  Release Number: {}", release_num);
            }
        }
        
        // Test QBXMLVersionsForSession (similar to C++ SDK)
        let ticket_variant = self.create_string_variant(ticket)?;
        if let Ok(result) = self.invoke_method(
            session_manager, 
            "QBXMLVersionsForSession", 
            &[&ticket_variant]
        ) {
            info!("  ✅ QBXMLVersionsForSession accessible");
            // Note: Parsing SAFEARRAY in Rust is complex, so we just verify it's callable
        }
        
        Ok(())
    }

    #[cfg(windows)]
    fn call_open_connection(&self, session_manager: &IDispatch, app_id: &str) -> Result<()> {
        info!("Trying OpenConnection (1 parameter)...");
        
        let app_id_variant = self.create_string_variant(app_id)?;
        
        self.invoke_method(session_manager, "OpenConnection", &[&app_id_variant])
            .map(|_| ())
    }

    #[cfg(windows)]
    fn close_connection(&self, session_manager: &IDispatch) -> Result<()> {
        info!("Closing QuickBooks connection...");
        
        self.invoke_method(session_manager, "CloseConnection", &[])
            .map(|_| ())
    }

    #[cfg(windows)]
    fn test_session_manager(&self, session_manager: &IDispatch) -> Result<()> {
        info!("🔍 Testing session manager functionality...");
        
        let mut tests_passed = 0;
        let mut tests_failed = 0;
        
        // Test 1: Check if we can get basic object information
        info!("  Test 1: Checking object type information...");
        match self.get_property(session_manager, "TypeInfo") {
            Ok(_) => {
                info!("    ✅ TypeInfo property accessible");
                tests_passed += 1;
            }
            Err(e) => {
                info!("    ❌ TypeInfo property failed: {}", e);
                tests_failed += 1;
            }
        }
        
        // Test 2: Try to get SDK version (this might work on some versions)
        info!("  Test 2: Checking SDK version...");
        match self.get_property(session_manager, "Version") {
            Ok(version_variant) => {
                if let Ok(version_str) = self.variant_to_string(&version_variant) {
                    info!("    ✅ QuickBooks SDK Version: {}", version_str);
                    tests_passed += 1;
                } else {
                    info!("    ⚠️ Version property exists but couldn't parse as string");
                    tests_passed += 1; // Still counts as working
                }
            }
            Err(e) => {
                info!("    ❌ Version property failed: {}", e);
                tests_failed += 1;
            }
        }
        
        // Test 3: Try a method that should be available even without a connection
        info!("  Test 3: Testing basic method availability...");
        match self.invoke_method(session_manager, "GetCurrentCompanyFileName", &[]) {
            Ok(_) => {
                info!("    ✅ GetCurrentCompanyFileName method is callable");
                tests_passed += 1;
            }
            Err(e) => {
                let error_msg = e.to_string();
                if error_msg.contains("No company file is open") || 
                   error_msg.contains("company file") ||
                   error_msg.contains("not connected") {
                    info!("    ✅ Method callable but no company file open: {}", error_msg);
                    tests_passed += 1;
                } else if error_msg.contains("0x80020009") {
                    info!("    ❌ COM exception - object may not be functional: {}", error_msg);
                    tests_failed += 1;
                } else if error_msg.contains("0x80020006") {
                    info!("    ❌ Method not found - wrong object type?: {}", error_msg);
                    tests_failed += 1;
                } else {
                    info!("    ❌ Unexpected error: {}", error_msg);
                    tests_failed += 1;
                }
            }
        }
        
        // Test 4: Try another basic method
        info!("  Test 4: Testing another basic method...");
        match self.invoke_method(session_manager, "GetIsConnected", &[]) {
            Ok(result) => {
                info!("    ✅ GetIsConnected method callable");
                // Try to interpret the result
                if let Ok(connected_str) = self.variant_to_string(&result) {
                    info!("    📋 Connection status: {}", connected_str);
                }
                tests_passed += 1;
            }
            Err(e) => {
                let error_msg = e.to_string();
                if error_msg.contains("0x80020006") {
                    info!("    ⚠️ GetIsConnected method not available on this version");
                    // Don't count as failure - method might not exist on all versions
                } else {
                    info!("    ❌ GetIsConnected failed: {}", error_msg);
                    tests_failed += 1;
                }
            }
        }
        
        // Summary
        info!("🔍 Session manager test results: {} passed, {} failed", tests_passed, tests_failed);
        
        if tests_passed > 0 {
            info!("✅ Session manager has some functionality - proceeding");
            Ok(())
        } else {
            Err(anyhow::anyhow!("Session manager failed all tests - object is not functional. Make sure QuickBooks Desktop is running with a company file open."))
        }
    }

    #[cfg(windows)]
    fn get_property(&self, dispatch: &IDispatch, property_name: &str) -> Result<VARIANT> {
        let property_name_bstr = BSTR::from(property_name);
        let names = [PCWSTR::from_raw(property_name_bstr.as_ptr())];
        let mut dispid = -1i32;
        
        unsafe {
            dispatch.GetIDsOfNames(
                &windows::core::GUID::zeroed(),
                names.as_ptr(),
                1,
                0x0409,
                &mut dispid,
            )?;
        }
        
        let dispparams = DISPPARAMS {
            rgvarg: std::ptr::null_mut(),
            cArgs: 0,
            rgdispidNamedArgs: std::ptr::null_mut(),
            cNamedArgs: 0,
        };
        
        let mut result = VARIANT::default();
        let mut excepinfo = EXCEPINFO::default();
        let mut arg_err = 0u32;
        
        unsafe {
            dispatch.Invoke(
                dispid,
                &windows::core::GUID::zeroed(),
                0x0409,
                DISPATCH_FLAGS(1), // DISPATCH_PROPERTYGET
                &dispparams,
                Some(&mut result),
                Some(&mut excepinfo),
                Some(&mut arg_err),
            )?;
        }
        
        Ok(result)
    }

    #[cfg(windows)]
    fn variant_to_dispatch(&self, variant: &VARIANT) -> Result<IDispatch> {
        unsafe {
            if let Some(dispatch) = variant.try_as_dispatch()? {
                Ok(dispatch)
            } else {
                Err(anyhow::anyhow!("Failed to convert VARIANT to IDispatch"))
            }
        }
    }

    fn variant_to_string(&self, variant: &VARIANT) -> Result<String> {
        match BSTR::try_from(variant) {
            Ok(bstr) => Ok(bstr.to_string()),
            Err(_) => Err(anyhow::anyhow!("Could not convert VARIANT to string"))
        }
    }

    // Helper methods for COM operations

    #[cfg(windows)]
    fn create_string_variant(&self, value: &str) -> Result<VARIANT> {
        let bstr = BSTR::from(value);
        Ok(VARIANT::from(bstr))
    }

    #[cfg(windows)]
    fn create_bool_variant(&self, value: bool) -> Result<VARIANT> {
        Ok(VARIANT::from(value as i16))
    }

    fn create_int_variant(&self, value: i32) -> Result<VARIANT> {
        Ok(VARIANT::from(value))
    }

    #[cfg(windows)]
    fn invoke_method(&self, dispatch: &IDispatch, method_name: &str, params: &[&VARIANT]) -> Result<VARIANT> {
        info!("Invoking method: {} with {} parameters", method_name, params.len());
        
        // Get method DISPID
        let method_name_bstr = BSTR::from(method_name);
        let names = [PCWSTR::from_raw(method_name_bstr.as_ptr())];
        let mut dispid = -1i32;
        
        info!("Getting DISPID for method: {}", method_name);
        unsafe {
            match dispatch.GetIDsOfNames(
                &windows::core::GUID::zeroed(),
                names.as_ptr(),
                1,
                0x0409, // LCID_ENGLISH_US
                &mut dispid,
            ) {
                Ok(()) => {
                    info!("Got DISPID {} for method {}", dispid, method_name);
                }
                Err(e) => {
                    error!("Failed to get DISPID for method {}: {:?}", method_name, e);
                    return Err(anyhow::anyhow!("Failed to get DISPID for method {}: {:?}", method_name, e));
                }
            }
        }
        
        // Prepare parameters (in reverse order for IDispatch)
        let mut param_array: Vec<VARIANT> = params.iter().rev().map(|&p| p.clone()).collect();
        let dispparams = DISPPARAMS {
            rgvarg: if param_array.is_empty() { std::ptr::null_mut() } else { param_array.as_mut_ptr() },
            cArgs: param_array.len() as u32,  // Use param_array.len() instead of params.len()
            rgdispidNamedArgs: std::ptr::null_mut(),
            cNamedArgs: 0,
        };
        
        // Keep param_array alive until after Invoke
        let _param_guard = param_array;
        
        info!("About to invoke method {} (DISPID: {})", method_name, dispid);
        
        // Invoke the method
        let mut result = VARIANT::default();
        let mut excepinfo = EXCEPINFO::default();
        let mut arg_err = 0u32;
        
            unsafe {
                let invoke_result = dispatch.Invoke(
                    dispid,
                    &windows::core::GUID::zeroed(),
                    0x0409,
                    DISPATCH_FLAGS(DISPATCH_METHOD.0 as u16),
                    &dispparams,
                    Some(&mut result),
                    Some(&mut excepinfo),
                    Some(&mut arg_err),
                );
                
                match invoke_result {
                Ok(()) => {
                    info!("Method {} invoked successfully", method_name);
                    Ok(result)
                }
                Err(e) => {
                    error!("Method {} invocation failed: {:?}", method_name, e);
                    error!("Exception info: SCODE={:?}, arg_err={}", excepinfo.scode, arg_err);
                    
                    // Try to get more detailed error information
                    if !excepinfo.bstrDescription.is_empty() {
                        let desc = excepinfo.bstrDescription.to_string();
                        error!("Exception description: {}", desc);
                    }
                    
                    Err(anyhow::anyhow!("Method {} invocation failed: {:?}", method_name, e))
                }
            }
        }
    }

    // New method for processing XML requests (similar to C++ SDK ProcessRequest)
    #[cfg(windows)]
    pub async fn process_xml_request(&self, xml_request: &str) -> Result<String> {
        info!("Processing XML request...");
        
        // Initialize COM with guard
        let _com = ComGuard::new()?;
        
        self.perform_xml_request(xml_request)
    }

    #[cfg(windows)]
    fn perform_xml_request(&self, xml_request: &str) -> Result<String> {
        info!("Creating session manager for XML request...");
        let session_manager = self.create_session_manager()?;
        
        let app_id = "QuickBooks-Sheets-Sync-v1";
        let app_name = "QuickBooks Sheets Sync";
        
        // Follow the C++ SDK pattern: OpenConnection -> BeginSession -> ProcessRequest -> EndSession -> CloseConnection
        
        // Step 1: OpenConnection
        self.try_open_connection(&session_manager, app_id, app_name)?;
        info!("✅ OpenConnection successful");
        
        // Step 2: BeginSession
        let ticket = self.begin_session(&session_manager, "")?; // Empty string means current company file
        info!("✅ BeginSession successful");
        
        // Step 3: ProcessRequest (the main XML processing step)
        let xml_response = self.process_request(&session_manager, &ticket, xml_request)?;
        info!("✅ ProcessRequest successful");
        
        // Step 4: EndSession
        self.end_session(&session_manager, &ticket)?;
        info!("✅ EndSession successful");
        
        // Step 5: CloseConnection
        self.close_connection(&session_manager)?;
        info!("✅ CloseConnection successful");
        
        Ok(xml_response)
    }

    #[cfg(windows)]
    fn process_request(&self, session_manager: &IDispatch, ticket: &str, xml_request: &str) -> Result<String> {
        info!("Processing QBXML request...");
        
        let ticket_variant = self.create_string_variant(ticket)?;
        let xml_request_variant = self.create_string_variant(xml_request)?;
        
        let result = self.invoke_method(
            session_manager, 
            "ProcessRequest", 
            &[&ticket_variant, &xml_request_variant]
        )?;
        
        let xml_response = self.variant_to_string(&result)?;
        info!("Received XML response ({} characters)", xml_response.len());
        
        Ok(xml_response)
    }

    // Method to create a simple account query XML request
    pub fn create_account_query_xml(&self) -> String {
        r#"<?xml version="1.0" ?>
<?qbxml version="8.0"?>
<QBXML>
   <QBXMLMsgsRq onError="stopOnError">
      <AccountQueryRq requestID="1">
      </AccountQueryRq>
   </QBXMLMsgsRq>
</QBXML>"#.to_string()
    }

    // Convenience method to test account data retrieval
    #[cfg(windows)]
    pub async fn test_account_data_retrieval(&self) -> Result<String> {
        let xml_request = self.create_account_query_xml();
        info!("Testing account data retrieval with XML request...");
        
        self.process_xml_request(&xml_request).await
    }

    #[cfg(not(windows))]
    pub async fn test_account_data_retrieval(&self) -> Result<String> {
        Err(anyhow::anyhow!("Account data retrieval is only supported on Windows"))
    }
}
