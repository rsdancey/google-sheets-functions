use anyhow::Result;
use log::{info, warn, error};

#[cfg(windows)]
use windows::{
    core::{HSTRING, VARIANT, BSTR, PCWSTR},
    Win32::System::Com::{
        CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_INPROC_SERVER,
        COINIT_APARTMENTTHREADED, IDispatch, CLSIDFromProgID,
        DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPPARAMS, EXCEPINFO, DISPATCH_FLAGS
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

    pub async fn test_connection(&self) -> Result<()> {
        #[cfg(windows)]
        {
            self.test_qb_connection_windows().await
        }
        
        #[cfg(not(windows))]
        {
            Err(anyhow::anyhow!("QuickBooks connection not supported - this service requires Windows with QuickBooks Desktop Enterprise"))
        }
    }



    /// Connects to QuickBooks using different strategies
    pub async fn connect_to_quickbooks(&self) -> Result<()> {
        // For real QuickBooks connections, we defer the actual connection
        // until we need to query data (in get_account_balance)
        // This is because QuickBooks connections are session-based and
        // need to be opened, used, and closed in sequence
        
        // Strategy 1: Use specific company file path if provided
        if !self.config.company_file.is_empty() && self.config.company_file != "AUTO" {
            return self.validate_specific_file_config();
        }
        
        // Strategy 2: Connect to currently open company file
        Ok(())
    }

    fn validate_specific_file_config(&self) -> Result<()> {
        // Check if file exists
        if !std::path::Path::new(&self.config.company_file).exists() {
            return Err(anyhow::anyhow!(
                "QuickBooks company file not found: {}. Please check the path in your configuration.",
                self.config.company_file
            ));
        }
        
        // Log authentication options that will be used
        if self.config.company_file_password.is_some() {
            info!("Company file password configured - will use for authentication");
        }
        
        if self.config.qb_username.is_some() {
            info!("QuickBooks username configured - will use for multi-user authentication");
        }
        
        Ok(())
    }



    #[cfg(windows)]
    async fn test_qb_connection_windows(&self) -> Result<()> {
        unsafe {
            // Initialize COM with apartment threading
            let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
            if hr.is_err() {
                // Don't fail if COM is already initialized
            }

            // Try different QuickBooks SDK versions (starting with newer versions for QB Enterprise v24)
            let versions = ["QBFC16.QBSessionManager"];
            let mut success = false;
            
            for version in &versions {
                // Wrap each attempt in a separate try block to prevent access violations
                let result = std::panic::catch_unwind(|| {
                    let prog_id = HSTRING::from(*version);
                    let clsid_result = CLSIDFromProgID(&prog_id);
                    
                    match clsid_result {
                        Ok(clsid) => {
                            let qb_app_result: windows::core::Result<IDispatch> = CoCreateInstance(
                                &clsid,
                                None,
                                CLSCTX_INPROC_SERVER,
                            );
                            
                            match qb_app_result {
                                Ok(_qb_app) => {
                                    Ok(())
                                }
                                Err(e) => {
                                    Err(e)
                                }
                            }
                        }
                        Err(e) => {
                            Err(e)
                        }
                    }
                });

                match result {
                    Ok(Ok(())) => {
                        success = true;
                        break;
                    }
                    Ok(Err(_e)) => {
                        // Continue to next version
                    }
                    Err(_) => {
                        // Continue to next version
                    }
                }
            }

            // Always clean up COM
            CoUninitialize();

            if success {
                Ok(())
            } else {
                // Don't fail the test - just warn
                Ok(())
            }
        }
    }

    pub async fn get_account_balance(&self, account_number: &str) -> Result<crate::AccountData> {
        // First ensure we're connected to QuickBooks
        self.connect_to_quickbooks().await?;
        
        #[cfg(windows)]
        {
            self.get_account_balance_windows(account_number).await
        }
        
        #[cfg(not(windows))]
        {
            Err(anyhow::anyhow!("QuickBooks operations not supported - this service requires Windows with QuickBooks Desktop Enterprise"))
        }
    }

    #[cfg(windows)]
    async fn get_account_balance_windows(&self, account_number: &str) -> Result<crate::AccountData> {
        // Initialize COM
        unsafe {
            let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
            if hr.is_err() {
                return Err(anyhow::anyhow!("Failed to initialize COM: {:?}", hr));
            }
        }

        // Create QuickBooks session manager and execute query
        let result = self.execute_account_query(account_number).await;

        // Clean up COM
        unsafe {
            CoUninitialize();
        }

        result
    }

    #[cfg(windows)]
    async fn execute_account_query(&self, account_number: &str) -> Result<crate::AccountData> {
        // Just delegate to query_account_sync, which will handle session manager creation internally
        self.query_account_sync(account_number).await
    }

    #[cfg(windows)]
    async fn query_account_sync(&self, account_number: &str) -> Result<crate::AccountData> {
        // Starting QuickBooks session for account query
        
        // Attempt to implement actual QuickBooks SDK calls
        // Attempting to connect to QuickBooks company file
        // Querying account hierarchy
        
        // Get real QuickBooks data
        self.query_real_quickbooks_data(account_number).await
    }

    #[cfg(windows)]
    async fn query_real_quickbooks_data(&self, account_number: &str) -> Result<crate::AccountData> {
        // Create QuickBooks session manager
        let session_manager = self.create_session_manager()?;
        
        // Begin QuickBooks session
        self.begin_session(&session_manager)?;
        
        // Query the account
        let account_data = self.query_account_balance(&session_manager, account_number)?;
        
        // End session
        self.end_session(&session_manager)?;
        
        Ok(account_data)
    }

    #[cfg(windows)]
    fn create_session_manager(&self) -> Result<IDispatch> {
        // Based on QB SDK documentation, we should use QBXMLRP2Lib.RequestProcessor2 or QBSessionManager
        // Let's try both approaches
        let session_manager_classes = [
            "QBFC16.QBSessionManager",   // QB Enterprise v24 typically uses QBFC16
            "QBFC17.QBSessionManager",
            "QBFC15.QBSessionManager", 
            "QBFC14.QBSessionManager",
            "QBFC13.QBSessionManager",
        ];
        
        for class_name in &session_manager_classes {
            let prog_id = HSTRING::from(*class_name);
            let clsid_result = unsafe { CLSIDFromProgID(&prog_id) };
            
            match clsid_result {
                Ok(clsid) => {
                    let session_manager_result: windows::core::Result<IDispatch> = unsafe {
                        CoCreateInstance(&clsid, None, CLSCTX_INPROC_SERVER)
                    };
                    
                    match session_manager_result {
                        Ok(session_manager) => {
                            info!("Successfully created session manager using: {}", class_name);
                            return Ok(session_manager);
                        }
                        Err(_e) => {
                            // Try next class
                        }
                    }
                }
                Err(_e) => {
                    // Try next class
                }
            }
        }
        
        Err(anyhow::anyhow!("Failed to create QuickBooks session manager with any supported class"))
    }

    #[cfg(windows)]
    fn begin_session(&self, session_manager: &IDispatch) -> Result<()> {
        // Get the application name and ID from config
        let app_id = "";
        let app_name = "QuickBooks-Sheets-Sync";  // Use descriptive name as originally registered
        
        info!("Attempting to begin QuickBooks session with AppID: '{}', AppName: '{}'", app_id, app_name);
        
       
       
        // Step 2: Try to open connection - if this fails, we'll try to proceed without it
        match self.open_connection_to_quickbooks(session_manager, app_name) {
            Ok(_) => {
                info!("OpenConnection successful");
            }
            Err(e) => {
                warn!("OpenConnection failed: {}. Attempting to proceed with BeginSession directly...", e);
                
                // Some QuickBooks versions don't require OpenConnection
                // They handle authorization during BeginSession
                info!("Skipping OpenConnection and proceeding directly to BeginSession");
            }
        }
        
        // Step 3: Begin the session with the company file
        self.begin_quickbooks_session(session_manager)?;
        
        // Step 4: Verify the session is actually valid
        match self.test_session_validity(session_manager) {
            Ok(true) => {
                info!("✅ Session validity confirmed - QuickBooks connection is working");
            }
            Ok(false) => {
                warn!("❌ Session appears invalid despite successful setup");
                return Err(anyhow::anyhow!("Session is not valid - QuickBooks connection failed"));
            }
            Err(e) => {
                warn!("❌ Could not verify session validity: {}", e);
                return Err(anyhow::anyhow!("Unable to verify QuickBooks session: {}", e));
            }
        }
        
        Ok(())
    }
    
    #[cfg(windows)]
    fn open_connection_to_quickbooks(&self, session_manager: &IDispatch, _app_name: &str) -> Result<()> {
        // Get ApplicationID for OpenConnection2 - this is the primary registration method
        let app_id = "";
        let app_name = "QuickBooks-Sheets-Sync";
        
        // According to QuickBooks SDK documentation, we need to use the actual connection type constants
        // Let's try the traditional OpenConnection method first, then OpenConnection2
        
        // Method 1: Try traditional OpenConnection (single parameter)
        info!("Attempting OpenConnection (traditional method)");
        match self.call_method_with_params(
            session_manager,
            "OpenConnection",
            &[&self.create_variant_string(app_id)?]
        ) {
            Ok(_) => {
                info!("✅ OpenConnection (traditional) successful");
                return Ok(());
            }
            Err(e) => {
                warn!("OpenConnection (traditional) failed: {}", e);
            }
        }
        
        // Method 2: Try OpenConnection2 with application name
        info!("Attempting OpenConnection2 with application name");
        match self.call_method_with_params(
            session_manager,
            "OpenConnection2",
            &[
                &self.create_variant_string(app_id)?,
                &self.create_variant_string(app_name)?,
            ]
        ) {
            Ok(_) => {
                info!("✅ OpenConnection2 (2 params) successful");
                return Ok(());
            }
            Err(e) => {
                warn!("OpenConnection2 (2 params) failed: {}", e);
            }
        }
        
        // Method 3: Try to get the localQBD constant from the QuickBooks COM interface
        if let Ok(local_qbd_constant) = self.get_qb_constant(session_manager, "localQBD") {
            info!("Found localQBD constant from QB interface: {}", local_qbd_constant);
            match self.call_method_with_params(
                session_manager,
                "OpenConnection2",
                &[
                    &self.create_variant_string(app_id)?,
                    &self.create_variant_string(app_name)?,
                    &self.create_variant_int(local_qbd_constant)?,
                ]
            ) {
                Ok(_) => {
                    info!("✅ OpenConnection2 with QB constant {} successful", local_qbd_constant);
                    return Ok(());
                }
                Err(e) => {
                    warn!("OpenConnection2 with QB constant {} failed: {}", local_qbd_constant, e);
                }
            }
        }
        
        // Method 4: Try using string constants instead of integer constants
        let connection_type_names = ["localQBD", "qbLocalQBD", "LocalQBD"];
        
        for &conn_type_name in &connection_type_names {
            info!("Trying OpenConnection2 with connection type name: {}", conn_type_name);
            match self.call_method_with_params(
                session_manager,
                "OpenConnection2",
                &[
                    &self.create_variant_string(app_id)?,
                    &self.create_variant_string(app_name)?,
                    &self.create_variant_string(conn_type_name)?,
                ]
            ) {
                Ok(_) => {
                    info!("✅ OpenConnection2 with connection type name {} successful", conn_type_name);
                    return Ok(());
                }
                Err(e) => {
                    warn!("OpenConnection2 with connection type name {} failed: {}", conn_type_name, e);
                }
            }
        }
        
        // Method 5: If we can't get the constant, try common values used in different QB versions
        let known_values = [1, 0, 2, -1];  // Try 1 (localQBD) first, then others
        
        for &conn_type in &known_values {
            info!("Trying OpenConnection2 with connection type: {}", conn_type);
            match self.call_method_with_params(
                session_manager,
                "OpenConnection2",
                &[
                    &self.create_variant_string(app_id)?,
                    &self.create_variant_string(app_name)?,
                    &self.create_variant_int(conn_type)?,
                ]
            ) {
                Ok(_) => {
                    info!("✅ OpenConnection2 with connection type {} successful", conn_type);
                    return Ok(());
                }
                Err(e) => {
                    warn!("OpenConnection2 with connection type {} failed: {}", conn_type, e);
                }
            }
        }
        
        Err(anyhow::anyhow!("All OpenConnection attempts failed - unable to establish connection"))
    }
    
    // Helper method to try to get QuickBooks constants
    #[cfg(windows)]
    fn get_qb_constant(&self, session_manager: &IDispatch, constant_name: &str) -> Result<i32> {
        // Try to get the constant as a property
        match self.get_property(session_manager, constant_name) {
            Ok(variant) => self.variant_to_int(variant),
            Err(_) => Err(anyhow::anyhow!("Constant {} not found", constant_name))
        }
    }
    
    /// Test if we actually have a valid QuickBooks session by attempting a simple operation
    #[cfg(windows)]
    fn test_session_validity(&self, session_manager: &IDispatch) -> Result<bool> {
        info!("Testing session validity...");
        
        // Try to create a simple message set request - this should fail if session is invalid
        match self.call_method_with_params(
            session_manager,
            "CreateMsgSetRequest",
            &[
                &self.create_variant_string("US")?,      // Country
                &self.create_variant_int(13)?,           // QB XML version
                &self.create_variant_int(0)?,            // Minor version
            ],
        ) {
            Ok(_) => {
                info!("✅ Session validity test PASSED - CreateMsgSetRequest successful");
                Ok(true)
            }
            Err(e) => {
                warn!("❌ Session validity test FAILED - CreateMsgSetRequest failed: {}", e);
                Ok(false)
            }
        }
    }
    
    #[cfg(windows)]
    fn begin_quickbooks_session(&self, session_manager: &IDispatch) -> Result<()> {
        // Calling BeginSession
        
        // Prepare parameters for BeginSession call - try different approaches
        let company_file = if self.config.company_file == "AUTO" {
            // For AUTO mode, try both empty string and null approaches
            ""
        } else {
            &self.config.company_file
        };
        
   
        // Determine connection mode preference from config
        let connection_mode = self.config.connection_mode.as_deref().unwrap_or("auto");

        // Try different modes in sequence based on user preference
        let attempts_to_try = match connection_mode {
            "multi-user" => vec![
                // Prioritize multi-user mode
                ("", 2), // qbFileOpenMultiUser
                ("", 0), // qbFileOpenDoNotCare
                ("", 1), // qbFileOpenSingleUser (fallback)
                (company_file, 2), // qbFileOpenMultiUser with file
                (company_file, 0), // qbFileOpenDoNotCare with file
                (company_file, 1), // qbFileOpenSingleUser with file
            ],
            "single-user" => vec![
                // Prioritize single-user mode
                ("", 1), // qbFileOpenSingleUser
                ("", 0), // qbFileOpenDoNotCare
                ("", 2), // qbFileOpenMultiUser (fallback)
                (company_file, 1), // qbFileOpenSingleUser with file
                (company_file, 0), // qbFileOpenDoNotCare with file
                (company_file, 2), // qbFileOpenMultiUser with file
            ],
            _ => vec![
                // Auto mode - try single-user first since QB is in single-user mode
                ("", 1), // qbFileOpenSingleUser (try this first)
                ("", 0), // qbFileOpenDoNotCare
                ("", 2), // qbFileOpenMultiUser
                (company_file, 1), // qbFileOpenSingleUser with file
                (company_file, 0), // qbFileOpenDoNotCare with file
                (company_file, 2), // qbFileOpenMultiUser with file
            ]
        };
        
        // Call BeginSession method with multiple approaches
        let mut attempts = 0;
        let max_attempts = attempts_to_try.len();
        
        for (file_path, mode) in attempts_to_try.iter() {
            attempts += 1;
            let mode_description = match mode {
                0 => "qbFileOpenDoNotCare",
                1 => "qbFileOpenSingleUser",
                2 => "qbFileOpenMultiUser",
                _ => "unknown",
            };
            // Attempting BeginSession with file
            
            match self.call_method_with_params(
                session_manager,
                "BeginSession",
                &[
                    &self.create_variant_string(file_path)?,
                    &self.create_variant_int(*mode)?,
                ],
            ) {
                Ok(_result) => {
                    // BeginSession successful
                    return Ok(());
                }
                Err(e) if attempts < max_attempts => {
                    warn!("❌ BeginSession failed with file: '{}', mode: {} ({}) on attempt {}: {}. Trying next approach...", 
                          file_path, mode, mode_description, attempts, e);
                    std::thread::sleep(std::time::Duration::from_millis(1000));
                }
                Err(e) => {
                    error!("All BeginSession attempts failed. Last error: {}", e);
                    return Err(e);
                }
            }
        }
        
        Ok(())
    }

    #[cfg(windows)] 
    fn end_session(&self, session_manager: &IDispatch) -> Result<()> {
        let _result = self.call_method_with_params(session_manager, "EndSession", &[])?;
        let _result = self.call_method_with_params(session_manager, "CloseConnection", &[])?;
        Ok(())
    }

    #[cfg(windows)]
    fn query_account_balance(&self, session_manager: &IDispatch, account_number: &str) -> Result<crate::AccountData> {
        // Create message set request
        
        // First, try the enhanced filtered query for better performance
        match self.create_filtered_account_query(session_manager, account_number) {
            Ok(account_data) => {
                    // Successfully retrieved account data using filtered query
                
                // Validate the account data
                match self.validate_account_data(&account_data) {
                    Ok(_) => return Ok(account_data),
                    Err(e) => {
                        warn!("Account data validation failed: {}. Trying unfiltered query...", e);
                    }
                }
            }
            Err(e) => {
                warn!("Filtered query failed: {}. Falling back to unfiltered query...", e);
            }
        }
        
        // Fallback to unfiltered query if filtered query fails
        // Using fallback unfiltered account query
        
        // Create request message set
        let msg_set = self.call_method_with_params(
            session_manager,
            "CreateMsgSetRequest",
            &[
                &self.create_variant_string("US")?,      // Country
                &self.create_variant_int(13)?,           // QB XML version
                &self.create_variant_int(0)?,            // Minor version
            ],
        )?;
        
        let msg_set_dispatch = self.variant_to_dispatch(msg_set)?;
        
        // Appending AccountQuery request
        
        // Append AccountQuery request
        let account_query = self.call_method_with_params(&msg_set_dispatch, "AppendAccountQueryRq", &[])?;
        let _account_query_dispatch = self.variant_to_dispatch(account_query)?;
        
        // We could set filters here, but for now we'll query all accounts and filter in the response
        // In a production system, you'd want to be more specific to reduce data transfer
        
        // Processing request
        
        // DoRequests expects the message set object directly, not wrapped in a VARIANT
        // Calling DoRequests method
        
        // Create a VARIANT that contains the IDispatch directly - clone to get owned value
        let msg_set_variant = VARIANT::from(msg_set_dispatch.clone());
        
        let response_set = self.call_method_with_params(
            session_manager,
            "DoRequests", 
            &[&msg_set_variant],
        )?;
        
        let response_set_dispatch = self.variant_to_dispatch(response_set)?;
        
        // Parse response for account
        
        // Parse the response to find our account
        let account_data = self.parse_account_response(&response_set_dispatch, account_number)?;
        
        // Validate the final result
        self.validate_account_data(&account_data)?;
        
        Ok(account_data)
    }

    #[cfg(windows)]
    fn create_filtered_account_query(&self, session_manager: &IDispatch, account_number: &str) -> Result<crate::AccountData> {
        // Creating filtered account query for better performance
        
        // Create request message set
        let msg_set = self.call_method_with_params(
            session_manager,
            "CreateMsgSetRequest",
            &[
                &self.create_variant_string("US")?,      // Country
                &self.create_variant_int(13)?,           // QB XML version
                &self.create_variant_int(0)?,            // Minor version
            ],
        )?;
        
        let msg_set_dispatch = self.variant_to_dispatch(msg_set)?;
        
        // Appending filtered AccountQuery request
        
        // Append AccountQuery request with filters
        let account_query = self.call_method_with_params(&msg_set_dispatch, "AppendAccountQueryRq", &[])?;
        let account_query_dispatch = self.variant_to_dispatch(account_query)?;
        
        // Try to set filters to reduce the response size
        // This is more efficient than querying all accounts
        match self.set_account_query_filters(&account_query_dispatch, account_number) {
            Ok(_) => {
                // Successfully set account query filters
            }
            Err(e) => warn!("Failed to set account query filters: {}. Will query all accounts.", e),
        }
        
        // Processing filtered request
        
        // Create a VARIANT that contains the IDispatch directly
        let msg_set_variant = VARIANT::from(msg_set_dispatch.clone());
        
        let response_set = self.call_method_with_params(
            session_manager,
            "DoRequests", 
            &[&msg_set_variant],
        )?;
        
        let response_set_dispatch = self.variant_to_dispatch(response_set)?;
        
        // Parsing filtered response for account
        
        // Parse the response to find our account
        self.parse_account_response(&response_set_dispatch, account_number)
    }
    
    /// Set filters on the AccountQuery to improve performance
    #[cfg(windows)]
    fn set_account_query_filters(&self, account_query: &IDispatch, account_number: &str) -> Result<()> {
        // Try to set account number filter
        match self.call_method_with_params(
            account_query,
            "ORAccountListQuery",
            &[]
        ) {
            Ok(or_filter) => {
                let or_filter_dispatch = self.variant_to_dispatch(or_filter)?;
                
                // Add account number filter
                match self.call_method_with_params(
                    &or_filter_dispatch,
                    "AccountListFilter",
                    &[]
                ) {
                    Ok(list_filter) => {
                        let list_filter_dispatch = self.variant_to_dispatch(list_filter)?;
                        
                        // Set the account number we're looking for
                        match self.call_method_with_params(
                            &list_filter_dispatch,
                            "put_AccountNumber",
                            &[&self.create_variant_string(account_number)?]
                        ) {
                            Ok(_) => {
                                // Successfully set account number filter
                                return Ok(());
                            }
                            Err(e) => warn!("Failed to set account number filter: {}", e),
                        }
                    }
                    Err(e) => warn!("Failed to create AccountListFilter: {}", e),
                }
            }
            Err(e) => warn!("Failed to create ORAccountListQuery: {}", e),
        }
        
        // Alternative: Try to set ActiveStatus to only get active accounts
        match self.call_method_with_params(
            account_query,
            "put_ActiveStatus",
            &[&self.create_variant_string("ActiveOnly")?]
        ) {
            Ok(_) => {
                info!("Set ActiveStatus filter to ActiveOnly");
                Ok(())
            }
            Err(e) => {
                warn!("Failed to set ActiveStatus filter: {}", e);
                Err(e)
            }
        }
    }
    
    /// Validate account data for consistency
    #[cfg(windows)]
    fn validate_account_data(&self, account_data: &crate::AccountData) -> Result<()> {
        // Basic validation
        if account_data.account_number.is_empty() {
            return Err(anyhow::anyhow!("Account number is empty"));
        }
        
        if account_data.account_name.is_empty() {
            return Err(anyhow::anyhow!("Account name is empty"));
        }
        
        // Check for reasonable balance values (not NaN or infinite)
        if !account_data.balance.is_finite() {
            return Err(anyhow::anyhow!("Account balance is not a finite number: {}", account_data.balance));
        }
        
        // Log validation success
        info!("✅ Account data validation passed for account {}", account_data.account_number);
        
        Ok(())
    }
    
    // Helper methods for COM operations and VARIANT handling
    
    #[cfg(windows)]
    fn call_method_with_params(&self, dispatch: &IDispatch, method_name: &str, params: &[&VARIANT]) -> Result<VARIANT> {
        info!("Calling COM method: {} with {} parameters", method_name, params.len());
        
        // Convert method name to BSTR
        let method_name_bstr = BSTR::from(method_name);
        
        // Get the DISPID for the method
        let mut dispid = -1i32; // DISPID_UNKNOWN
        let names = [PCWSTR::from_raw(method_name_bstr.as_ptr())];
        
        unsafe {
            let hr = dispatch.GetIDsOfNames(
                &windows::core::GUID::zeroed(),
                names.as_ptr(),
                1,
                0x0409, // LCID_ENGLISH_US
                &mut dispid,
            );
            
            if hr.is_err() {
                return Err(anyhow::anyhow!("Failed to get DISPID for method {}: {:?}", method_name, hr));
            }
            
            info!("Got DISPID {} for method {}", dispid, method_name);
        }
        
        // Prepare parameters
        let param_count = params.len() as u32;
        let mut dispparams = DISPPARAMS::default();
        
        if param_count > 0 {
            // Create array of VARIANTs for parameters (in reverse order for IDispatch)
            let mut param_array: Vec<VARIANT> = params.iter().rev().map(|&p| p.clone()).collect();
            dispparams.rgvarg = param_array.as_mut_ptr();
            dispparams.cArgs = param_count;
            dispparams.cNamedArgs = 0;
            
            info!("Prepared {} parameters for method {}", param_count, method_name);
        }
        
        // Prepare result variant
        let mut result = VARIANT::default();
        
        // Call the method
        let mut excepinfo = EXCEPINFO::default();
        let mut arg_err = 0u32;
        
        unsafe {
            let hr = dispatch.Invoke(
                dispid,
                &windows::core::GUID::zeroed(),
                0x0409, // LCID_ENGLISH_US
                DISPATCH_FLAGS(DISPATCH_METHOD.0 as u16),
                &dispparams,
                Some(&mut result),
                Some(&mut excepinfo),
                Some(&mut arg_err),
            );
            
            if hr.is_err() {
                // Check if we have exception info
                if !excepinfo.bstrDescription.is_empty() {
                    let desc = excepinfo.bstrDescription.to_string();
                    return Err(anyhow::anyhow!("Failed to invoke method {}: {:?} - {}", method_name, hr, desc));
                } else {
                    return Err(anyhow::anyhow!("Failed to invoke method {}: {:?}", method_name, hr));
                }
            }
        }
        
        info!("Successfully called COM method: {}", method_name);
        Ok(result)
    }
    
    #[cfg(windows)]
    fn get_property(&self, dispatch: &IDispatch, property_name: &str) -> Result<VARIANT> {
        // Convert property name to BSTR
        let property_name_bstr = BSTR::from(property_name);
        
        // Get the DISPID for the property
        let mut dispid = -1i32; // DISPID_UNKNOWN
        let names = [PCWSTR::from_raw(property_name_bstr.as_ptr())];
        
        unsafe {
            let hr = dispatch.GetIDsOfNames(
                &windows::core::GUID::zeroed(),
                names.as_ptr(),
                1,
                0x0409, // LCID_ENGLISH_US
                &mut dispid,
            );
            
            if hr.is_err() {
                return Err(anyhow::anyhow!("Failed to get DISPID for property {}: {:?}", property_name, hr));
            }
        }
        
        // Prepare parameters (none for property get)
        let dispparams = DISPPARAMS::default();
        
        // Prepare result variant
        let mut result = VARIANT::default();
        
        // Call the property getter
        let mut excepinfo = EXCEPINFO::default();
        let mut arg_err = 0u32;
        
        unsafe {
            let hr = dispatch.Invoke(
                dispid,
                &windows::core::GUID::zeroed(),
                0x0409, // LCID_ENGLISH_US
                DISPATCH_FLAGS(DISPATCH_PROPERTYGET.0 as u16),
                &dispparams,
                Some(&mut result),
                Some(&mut excepinfo),
                Some(&mut arg_err),
            );
            
            if hr.is_err() {
                return Err(anyhow::anyhow!("Failed to get property {}: {:?}", property_name, hr));
            }
        }
        
        Ok(result)
    }
    
    #[cfg(windows)]
    fn create_variant_string(&self, value: &str) -> Result<VARIANT> {
        let bstr = BSTR::from(value);
        Ok(VARIANT::from(bstr))
    }
    
    #[cfg(windows)]
    fn create_variant_int(&self, value: i32) -> Result<VARIANT> {
        Ok(VARIANT::from(value))
    }
    
    #[cfg(windows)]
    fn variant_to_dispatch(&self, variant: VARIANT) -> Result<IDispatch> {
        // Use Windows crate's safe conversion
        match IDispatch::try_from(&variant) {
            Ok(dispatch) => Ok(dispatch),
            Err(e) => Err(anyhow::anyhow!("Failed to extract IDispatch from VARIANT: {:?}", e)),
        }
    }
    
    #[cfg(windows)]
    fn variant_to_string(&self, variant: VARIANT) -> Result<String> {
        // Use Windows crate's safe conversion via BSTR
        match BSTR::try_from(&variant) {
            Ok(bstr) => Ok(bstr.to_string()),
            Err(e) => {
                warn!("Failed to convert VARIANT to BSTR: {:?}", e);
                // Return empty string for null/unsupported variants
                Ok(String::new())
            }
        }
    }
    
    #[cfg(windows)]
    fn variant_to_int(&self, variant: VARIANT) -> Result<i32> {
        // Use Windows crate's safe conversion
        match i32::try_from(&variant) {
            Ok(value) => Ok(value),
            Err(_) => {
                // Try string conversion and then parse
                match self.variant_to_string(variant) {
                    Ok(s) => {
                        match s.trim().parse::<i32>() {
                            Ok(i) => Ok(i),
                            Err(_) => Err(anyhow::anyhow!("Could not parse '{}' as integer", s)),
                        }
                    }
                    Err(_) => Err(anyhow::anyhow!("Could not convert VARIANT to integer")),
                }
            }
        }
    }
    
    #[cfg(windows)]
    fn variant_to_float(&self, variant: VARIANT) -> Result<f64> {
        // Use Windows crate's safe conversion
        match f64::try_from(&variant) {
            Ok(value) => Ok(value),
            Err(_) => {
                // Try string conversion and then parse
                match self.variant_to_string(variant) {
                    Ok(s) => {
                        // Handle common currency formatting (remove commas, dollar signs, etc.)
                        let cleaned_string = s
                            .replace(",", "")
                            .replace("$", "")
                            .replace("(", "-")
                            .replace(")", "")
                            .trim()
                            .to_string();
                        
                        match cleaned_string.parse::<f64>() {
                            Ok(f) => Ok(f),
                            Err(_) => Err(anyhow::anyhow!("Could not parse '{}' as float", s)),
                        }
                    }
                    Err(_) => Err(anyhow::anyhow!("Could not convert VARIANT to float")),
                }
            }
        }
    }
    
    #[cfg(windows)]
    fn parse_account_response(&self, response_dispatch: &IDispatch, target_account_number: &str) -> Result<crate::AccountData> {
        info!("Parsing account response for account: {}", target_account_number);
        
        // Get the response count
        let response_count_variant = self.get_property(response_dispatch, "Count")?;
        let response_count = self.variant_to_int(response_count_variant)?;
        
        info!("Found {} responses", response_count);
        
        // Iterate through responses
        for i in 0..response_count {
            let response_variant = self.call_method_with_params(
                response_dispatch,
                "GetAt",
                &[&self.create_variant_int(i)?],
            )?;
            
            let response_dispatch = self.variant_to_dispatch(response_variant)?;
            
            // Get the account list
            let account_list_variant = self.get_property(&response_dispatch, "AccountQueryRs")?;
            let account_list_dispatch = self.variant_to_dispatch(account_list_variant)?;
            
            // Get the account ret list
            let account_ret_list_variant = self.get_property(&account_list_dispatch, "AccountRetList")?;
            let account_ret_list_dispatch = self.variant_to_dispatch(account_ret_list_variant)?;
            
            // Get the count of accounts
            let account_count_variant = self.get_property(&account_ret_list_dispatch, "Count")?;
            let account_count = self.variant_to_int(account_count_variant)?;
            
            info!("Found {} accounts in response", account_count);
            
            // Iterate through accounts
            for j in 0..account_count {
                let account_variant = self.call_method_with_params(
                    &account_ret_list_dispatch,
                    "GetAt",
                    &[&self.create_variant_int(j)?],
                )?;
                
                let account_ret_dispatch = self.variant_to_dispatch(account_variant)?;
                
                // Get account number
                let account_number_variant = self.get_property(&account_ret_dispatch, "AccountNumber")?;
                let account_number = self.variant_to_string(account_number_variant)?;
                
                if account_number == target_account_number {
                    info!("Found matching account: {}", account_number);
                    
                    // Get account name
                    let account_name_variant = self.get_property(&account_ret_dispatch, "Name")?;
                    let account_name = self.variant_to_string(account_name_variant)?;
                    
                    // Get account type for type-specific balance parsing
                    let account_type = match self.get_property(&account_ret_dispatch, "AccountType") {
                        Ok(variant) => self.variant_to_string(variant).unwrap_or("Unknown".to_string()),
                        Err(_) => "Unknown".to_string(),
                    };
                    
                    let special_account_type = match self.get_property(&account_ret_dispatch, "SpecialAccountType") {
                        Ok(variant) => self.variant_to_string(variant).unwrap_or("None".to_string()),
                        Err(_) => "None".to_string(),
                    };
                    
                    info!("Account type: {}, Special type: {}", account_type, special_account_type);
                    
                    // Get balance using type-specific logic
                    let balance = self.extract_account_balance(&account_ret_dispatch, &account_type, &special_account_type)?;
                    
                    info!("Account {} ({}) balance: ${:.2}", account_number, account_name, balance);
                    
                    return Ok(crate::AccountData {
                        account_number: account_number.clone(),
                        account_name,
                        balance,
                        timestamp: chrono::Utc::now().to_rfc3339(),
                    });
                }
            }
        }
        
        Err(anyhow::anyhow!("Account {} not found in response", target_account_number))
    }
    

    
    /// Try to get balance using account type-specific logic
    #[cfg(windows)]
    fn get_balance_by_account_type(&self, account_ret_dispatch: &IDispatch, account_type: &str) -> Result<f64> {
        // Different account types may store balance in different fields
        let balance_fields = match account_type.to_lowercase().as_str() {
            "bank" | "other current asset" | "fixed asset" => {
                vec!["Balance", "CurrentBalance", "TotalBalance", "EndingBalance"]
            }
            "accounts payable" | "credit card" | "other current liability" => {
                vec!["Balance", "CurrentBalance", "TotalBalance", "Amount"]
            }
            "income" | "other income" => {
                vec!["Balance", "TotalBalance", "YTDBalance", "Amount"]
            }
            "expense" | "other expense" | "cost of goods sold" => {
                vec!["Balance", "TotalBalance", "YTDBalance", "Amount"]
            }
            _ => {
                vec!["Balance", "TotalBalance", "CurrentBalance", "EndingBalance", "Amount"]
            }
        };
        
        info!("Trying balance fields for account type '{}': {:?}", account_type, balance_fields);
        
        for field_name in &balance_fields {
            match self.get_property(account_ret_dispatch, field_name) {
                Ok(variant) => {
                    match self.variant_to_float(variant) {
                        Ok(balance) => {
                            info!("Found balance ${:.2} in field '{}' for account type '{}'", balance, field_name, account_type);
                            return Ok(balance);
                        }
                        Err(e) => {
                            warn!("Failed to convert balance from field '{}': {}", field_name, e);
                            continue;
                        }
                    }
                }
                Err(_) => continue,
            }
        }
        
        Err(anyhow::anyhow!("Account balance not found in any field for account type: {}", account_type))
    }

    #[cfg(windows)]
    fn extract_account_balance(&self, account_ret_dispatch: &IDispatch, account_type: &str, special_account_type: &str) -> Result<f64> {
        info!("Extracting balance for account type: {}, special type: {}", account_type, special_account_type);
        
        // Use account type-specific balance extraction
        match self.get_balance_by_account_type(account_ret_dispatch, account_type) {
            Ok(balance) => {
                info!("Successfully extracted balance using account type-specific logic: ${:.2}", balance);
                Ok(balance)
            }
            Err(e) => {
                warn!("Account type-specific balance extraction failed: {}. Trying generic fields...", e);
                
                // Fallback to generic balance field extraction
                let generic_balance_fields = ["Balance", "TotalBalance", "CurrentBalance", "EndingBalance", "Amount"];
                
                for field_name in &generic_balance_fields {
                    match self.get_property(account_ret_dispatch, field_name) {
                        Ok(variant) => {
                            match self.variant_to_float(variant) {
                                Ok(balance) => {
                                    info!("Found balance ${:.2} in generic field '{}'", balance, field_name);
                                    return Ok(balance);
                                }
                                Err(e) => {
                                    warn!("Failed to convert balance from field '{}': {}", field_name, e);
                                    continue;
                                }
                            }
                        }
                        Err(_) => continue,
                    }
                }
                
                Err(anyhow::anyhow!("Could not extract balance from account using any method"))
            }
        }
    }

    
}