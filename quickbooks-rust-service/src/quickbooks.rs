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

#[derive(Debug, Clone)]
pub struct QuickBooksClient {
    config: QuickBooksConfig,
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
        
        unsafe {
            let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
            if hr.is_err() {
                // COM may already be initialized, which is fine
            }
        }

        let result = self.create_session_manager();

        unsafe {
            CoUninitialize();
        }

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
        
        unsafe {
            let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
            if hr.is_err() {
                // COM may already be initialized, which is fine
            }
        }

        let result = self.register_application();

        unsafe {
            CoUninitialize();
        }

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
                    
                    // Test the session manager functionality
                    match self.test_session_manager(&session_manager) {
                        Ok(_) => {
                            info!("✅ Session manager functionality test passed!");
                            return Ok(session_manager);
                        }
                        Err(e) => {
                            error!("❌ Session manager functionality test failed: {}", e);
                            info!("💡 This indicates the session manager object was created but is not functional");
                            info!("💡 Common causes:");
                            info!("   - QuickBooks Desktop is not running");
                            info!("   - No company file is open in QuickBooks");
                            info!("   - QuickBooks is not in a ready state");
                            info!("   - Permission/security issues");
                            continue; // Try next class
                        }
                    }
                }
            }
        }
        
        Err(anyhow::anyhow!("Failed to create QuickBooks session manager. Ensure QuickBooks Desktop is installed."))
    }

    #[cfg(windows)]
    fn register_application(&self) -> Result<()> {
        info!("Creating QuickBooks session manager...");
        let session_manager = self.create_session_manager()?;
        
        info!("Attempting to register application with QuickBooks...");
        
        // Application details for registration
        let app_id = "QuickBooks-Sheets-Sync-v1";
        let app_name = "QuickBooks Sheets Sync";
        
        // Try different OpenConnection methods
        self.try_open_connection(&session_manager, app_id, app_name)?;
        
        info!("✅ Successfully registered with QuickBooks!");
        
        // Clean up the connection
        self.close_connection(&session_manager)?;
        
        Ok(())
    }

    #[cfg(windows)]
    fn try_open_connection(&self, session_manager: &IDispatch, app_id: &str, app_name: &str) -> Result<()> {
        info!("Attempting OpenConnection with AppID: '{}', AppName: '{}'", app_id, app_name);
        
        // Method 1: Try OpenConnection (traditional single parameter method)
        if let Ok(_) = self.call_open_connection(session_manager, app_id) {
            info!("✅ OpenConnection successful (traditional method)");
            return Ok(());
        }
        
        // Method 2: Try OpenConnection2 with 2 parameters
        if let Ok(_) = self.call_open_connection2_basic(session_manager, app_id, app_name) {
            info!("✅ OpenConnection2 successful (basic method)");
            return Ok(());
        }
        
        // Method 3: Try OpenConnection2 with connection type parameter
        if let Ok(_) = self.call_open_connection2_with_type(session_manager, app_id, app_name) {
            info!("✅ OpenConnection2 successful (with connection type)");
            return Ok(());
        }
        
        Err(anyhow::anyhow!("All OpenConnection methods failed. QuickBooks may need to authorize this application."))
    }

    #[cfg(windows)]
    fn call_open_connection(&self, session_manager: &IDispatch, app_id: &str) -> Result<()> {
        info!("Trying OpenConnection (1 parameter)...");
        
        let app_id_variant = self.create_string_variant(app_id)?;
        
        self.invoke_method(session_manager, "OpenConnection", &[&app_id_variant])
            .map(|_| ())
    }

    #[cfg(windows)]
    fn call_open_connection2_basic(&self, session_manager: &IDispatch, app_id: &str, app_name: &str) -> Result<()> {
        info!("Trying OpenConnection2 (2 parameters)...");
        
        let app_id_variant = self.create_string_variant(app_id)?;
        let app_name_variant = self.create_string_variant(app_name)?;
        
        self.invoke_method(session_manager, "OpenConnection2", &[&app_id_variant, &app_name_variant])
            .map(|_| ())
    }

    #[cfg(windows)]
    fn call_open_connection2_with_type(&self, session_manager: &IDispatch, app_id: &str, app_name: &str) -> Result<()> {
        info!("Trying OpenConnection2 (3 parameters with connection types)...");
        
        let app_id_variant = self.create_string_variant(app_id)?;
        let app_name_variant = self.create_string_variant(app_name)?;
        
        // Try different connection type values that are commonly used for local QuickBooks Desktop
        let connection_types = [1, 0]; // 1 = localQBD (most common), 0 = default
        
        for &conn_type in &connection_types {
            info!("  Trying connection type: {}", conn_type);
            let conn_type_variant = self.create_int_variant(conn_type)?;
            
            if let Ok(_) = self.invoke_method(
                session_manager, 
                "OpenConnection2", 
                &[&app_id_variant, &app_name_variant, &conn_type_variant]
            ) {
                info!("  ✅ Success with connection type: {}", conn_type);
                return Ok(());
            }
        }
        
        Err(anyhow::anyhow!("OpenConnection2 failed with all connection types"))
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
    fn create_int_variant(&self, value: i32) -> Result<VARIANT> {
        Ok(VARIANT::from(value))
    }

    #[cfg(windows)]
    fn invoke_method(&self, dispatch: &IDispatch, method_name: &str, params: &[&VARIANT]) -> Result<VARIANT> {
        // Get method DISPID
        let method_name_bstr = BSTR::from(method_name);
        let names = [PCWSTR::from_raw(method_name_bstr.as_ptr())];
        let mut dispid = -1i32;
        
        unsafe {
            dispatch.GetIDsOfNames(
                &windows::core::GUID::zeroed(),
                names.as_ptr(),
                1,
                0x0409, // LCID_ENGLISH_US
                &mut dispid,
            )?;
        }
        
        // Prepare parameters (in reverse order for IDispatch)
        let mut param_array: Vec<VARIANT> = params.iter().rev().map(|&p| p.clone()).collect();
        let dispparams = DISPPARAMS {
            rgvarg: if param_array.is_empty() { std::ptr::null_mut() } else { param_array.as_mut_ptr() },
            cArgs: params.len() as u32,
            rgdispidNamedArgs: std::ptr::null_mut(),
            cNamedArgs: 0,
        };
        
        // Invoke the method
        let mut result = VARIANT::default();
        let mut excepinfo = EXCEPINFO::default();
        let mut arg_err = 0u32;
        
        unsafe {
            dispatch.Invoke(
                dispid,
                &windows::core::GUID::zeroed(),
                0x0409,
                DISPATCH_FLAGS(DISPATCH_METHOD.0 as u16),
                &dispparams,
                Some(&mut result),
                Some(&mut excepinfo),
                Some(&mut arg_err),
            )?;
        }
        
        Ok(result)
    }
}
