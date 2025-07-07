use anyhow::Result;
use log::{info, warn, error};

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
            CoInitializeEx(None, COINIT_APARTMENTTHREADED)?;
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
            CoInitializeEx(None, COINIT_APARTMENTTHREADED)?;
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
                    CoCreateInstance::<IDispatch>(&clsid, None, CLSCTX_INPROC_SERVER) 
                } {
                    info!("✅ Successfully created session manager: {}", class_name);
                    return Ok(session_manager);
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
        let mut dispparams = DISPPARAMS {
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
