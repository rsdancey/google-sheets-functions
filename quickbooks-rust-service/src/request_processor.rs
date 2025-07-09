use winapi::shared::guiddef::{CLSID, IID_NULL};
use winapi::um::oaidl::{IDispatch, VARIANT, EXCEPINFO};
use crate::safe_variant::SafeVariant;
use crate::FileMode;

const DISPATCH_METHOD: u16 = 1;

#[allow(non_upper_case_globals)]
pub const IID_IDispatch: winapi::shared::guiddef::GUID = winapi::shared::guiddef::GUID {
    Data1: 0x00020400,
    Data2: 0x0000,
    Data3: 0x0000,
    Data4: [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46],
};

#[derive(Debug, Clone)]
pub struct AccountInfo {
    pub name: String,
    pub number: String,
    pub account_type: String,
    pub balance: f64,
}

/// Type-safe wrapper for QBFC SessionManager 
/// This uses the QBFC API (QBFC16.QBSessionManager) instead of QBXML API (QBXMLRP2.RequestProcessor)
/// The QBFC API is more reliable for COM interop and uses different parameter types
pub struct RequestProcessor2 {
    inner: *mut IDispatch,
}

impl RequestProcessor2 {
    pub fn new() -> Result<Self, anyhow::Error> {
        // Try QBFC ProgIDs - use the working QBFC API instead of QBXML
        let prog_ids_to_try = [
            "QBFC16.QBSessionManager",         // QB 2024/2023 - most likely
            "QBFC15.QBSessionManager",         // QB 2022
            "QBFC14.QBSessionManager",         // QB 2021
            "QBFC13.QBSessionManager",         // QB 2020 - fallback
        ];

        for prog_id_str in prog_ids_to_try.iter() {
            log::info!("Trying QBFC ProgID: {}", prog_id_str);
            let prog_id_wide = widestring::U16CString::from_str(*prog_id_str).unwrap();
            let mut clsid: CLSID = unsafe { std::mem::zeroed() };
            let hr = unsafe {
                winapi::um::combaseapi::CLSIDFromProgID(
                    prog_id_wide.as_ptr(),
                    &mut clsid as *mut CLSID
                )
            };
            if hr < 0 {
                log::warn!("ProgID {} not found or CLSIDFromProgID failed: HRESULT=0x{:08X}", prog_id_str, hr as u32);
                continue;
            }
            let mut dispatch_ptr: *mut IDispatch = std::ptr::null_mut();
            let hr = unsafe {
                winapi::um::combaseapi::CoCreateInstance(
                    &clsid,
                    std::ptr::null_mut(),
                    winapi::shared::wtypesbase::CLSCTX_INPROC_SERVER,
                    &IID_IDispatch,
                    &mut dispatch_ptr as *mut _ as *mut _
                )
            };
            if hr >= 0 && !dispatch_ptr.is_null() {
                log::info!("✅ Successfully created QBFC COM instance with ProgID: {}", prog_id_str);
                let instance = Self {
                    inner: dispatch_ptr,
                };
                log::info!("RequestProcessor2::new: instance address = {:p}, COM inner = {:p}", &instance, instance.inner);
                return Ok(instance);
            } else {
                log::warn!("Failed to create COM instance for {}: HRESULT=0x{:08X}", prog_id_str, hr as u32);
            }
        }
        Err(anyhow::anyhow!("Failed to create QBFC COM instance for all ProgIDs"))
    }

    fn invoke_method(&self, method_name: &str, params: &[SafeVariant]) -> Result<SafeVariant, anyhow::Error> {
        // Log parameter types and values for debugging
        // Log BSTR details for each parameter before COM call
        for (i, param) in params.iter().enumerate() {
            let vt = unsafe { param.as_variant().n1.n2().vt };
            if vt == winapi::shared::wtypes::VT_BSTR as u16 {
                let bstr = unsafe { *param.as_variant().n1.n2().n3.bstrVal() };
                if !bstr.is_null() {
                    let len = unsafe { winapi::um::oleauto::SysStringLen(bstr) } as usize;
                    let slice = unsafe { std::slice::from_raw_parts(bstr, len) };
                    println!("[invoke_method] param[{}] BSTR ptr={:p} len={} utf16={:?}", i, bstr, len, &slice[..std::cmp::min(len, 8)]);
                } else {
                    println!("[invoke_method] param[{}] BSTR ptr=NULL", i);
                }
            }
            println!("[invoke_method] param[{}] vt={}", i, vt);
        }
        let mut dispid = 0i32;
        let method_name_wide = widestring::U16CString::from_str(method_name).unwrap();
        let names = [method_name_wide.as_ptr()];
        let hr = unsafe {
            ((*(*self.inner).lpVtbl).GetIDsOfNames)(
                self.inner,
                &IID_NULL,
                names.as_ptr() as *mut _,
                1,
                0x0409,
                &mut dispid
            )
        };
        if hr < 0 {
            return Err(anyhow::anyhow!("GetIDsOfNames failed: HRESULT=0x{:08X}", hr));
        }
        // --- FIX: Ensure VARIANTs outlive the COM call ---
        let mut variants: Vec<winapi::um::oaidl::VARIANT> = params.iter().map(|v| v.to_winvariant()).collect();
        let mut dispparams = winapi::um::oaidl::DISPPARAMS {
            rgvarg: if variants.is_empty() { std::ptr::null_mut() } else { variants.as_mut_ptr() },
            rgdispidNamedArgs: std::ptr::null_mut(),
            cArgs: variants.len() as u32,
            cNamedArgs: 0,
        };
        let mut result: VARIANT = unsafe { std::mem::zeroed() };
        let mut excepinfo: EXCEPINFO = unsafe { std::mem::zeroed() };
        let mut arg_err = 0u32;
        let hr = unsafe {
            ((*(*self.inner).lpVtbl).Invoke)(
                self.inner,
                dispid,
                &IID_NULL,
                0x0409,
                DISPATCH_METHOD,
                &mut dispparams,
                &mut result,
                &mut excepinfo,
                &mut arg_err
            )
        };
        if hr < 0 {
            // Extract EXCEPINFO details for diagnostics
            let bstr_to_string = |bstr: *mut u16| {
                if bstr.is_null() { return String::new(); }
                unsafe {
                    let len = (0..).take_while(|&i| *bstr.offset(i) != 0).count();
                    let slice = std::slice::from_raw_parts(bstr, len);
                    String::from_utf16_lossy(slice)
                }
            };
            let description = bstr_to_string(excepinfo.bstrDescription);
            let source = bstr_to_string(excepinfo.bstrSource);
            let helpfile = bstr_to_string(excepinfo.bstrHelpFile);
            log::error!(
                "COM Invoke failed: method={method_name}, HRESULT=0x{hr:08X}, arg_err={},\n  EXCEPINFO: code={}, wCode={}, source='{}', description='{}', helpfile='{}', helpctx={}, scode=0x{:08X}",
                arg_err,
                excepinfo.wCode,
                excepinfo.wCode,
                source,
                description,
                helpfile,
                excepinfo.dwHelpContext,
                excepinfo.scode as u32
            );
            return Err(anyhow::anyhow!(
                "Invoke failed: method={method}, HRESULT=0x{hr:08X}, description='{description}', source='{source}', helpfile='{helpfile}', scode=0x{scode:08X}",
                method=method_name,
                hr=hr,
                description=description,
                source=source,
                helpfile=helpfile,
                scode=excepinfo.scode as u32
            ));
        }
        Ok(SafeVariant::from_winvariant(&result))
    }

    pub fn open_connection(&self, _app_id: &str, app_name: &str) -> Result<(), anyhow::Error> {
        log::info!("open_connection: self address = {:p}, COM inner = {:p}", self, self.inner);
        // Always pass empty string for AppID to avoid accidental registration
        let app_id_var = SafeVariant::from_string("");
        let app_name_var = SafeVariant::from_string(app_name);
        // we call the parameters in the reverse order from the IDL because it works
        match self.invoke_method("OpenConnection", &[app_name_var, app_id_var]) {
            Ok(_) => {
                log::info!("✅ OpenConnection successful (signature: AppID, AppName)");
                Ok(())
            },
            Err(e) => {
                log::error!("OpenConnection failed: {:#}", e);
                for cause in e.chain().skip(1) {
                    log::error!("Caused by: {:#}", cause);
                }
                Err(anyhow::anyhow!("Failed to open QuickBooks connection. See error logs above for HRESULT, EXCEPINFO, and details."))
            }
        }
    }

    pub fn begin_session(&self, company_file: &str, file_mode: FileMode) -> Result<*mut IDispatch, anyhow::Error> {
        log::info!("begin_session: self address = {:p}, COM inner = {:p}", self, self.inner);
        log::info!("Attempting to begin QuickBooks session...");
        let file_var = SafeVariant::from_string(company_file);
        let mode_int = match file_mode {
            FileMode::SingleUser => 1,
            FileMode::MultiUser => 2,
            FileMode::DoNotCare => 2, // per IDL: omDontCare = 2
            FileMode::Online => 3,
        };
        let mode_var = SafeVariant::from_i32(mode_int);
        // Correct COM parameter order: [mode_var, file_var]
        // DO NOT REVERSE THE ORDER
        let result = self.invoke_method("BeginSession", &[mode_var, file_var])?;
        // Log the VARIANT type and pointer before attempting as_dispatch
        let vt = unsafe { result.as_variant().n1.n2().vt };
        let dispatch_ptr = if vt == winapi::shared::wtypes::VT_DISPATCH as u16 {
            unsafe { *result.as_variant().n1.n2().n3.pdispVal() }
        } else {
            std::ptr::null_mut()
        };
        log::info!("BeginSession returned VARIANT vt={} (expected {}), dispatch_ptr={:p}", vt, winapi::shared::wtypes::VT_DISPATCH, dispatch_ptr);
        // Per IDL: returns ISession** (VT_DISPATCH)
        let session_dispatch = result.as_dispatch()?;
        log::info!("✅ BeginSession successful, got session IDispatch pointer: {:p}", session_dispatch);
        // --- Extra validation: try calling a harmless method on the session object ---
        // We'll attempt to call 'EndSession' (should succeed if session is valid)
        // If this fails, log a warning but do not return error here
        if session_dispatch.is_null() {
            log::warn!("BeginSession returned a null session pointer!");
        } else {
            // Try a harmless method call to validate the session object
            let test_result = self.invoke_method_on_dispatch(session_dispatch, "get_Class", &[]);
            match test_result {
                Ok(val) => log::info!("Session object responded to get_Class: {}", val.to_string().unwrap_or_else(|| "<non-string>".to_string())),
                Err(e) => log::warn!("Session object did not respond to get_Class: {:#}", e),
            }
        }
        Ok(session_dispatch)
    }

    pub fn end_session(&self) -> Result<(), anyhow::Error> {
        log::info!("end_session: self address = {:p}, COM inner = {:p}", self, self.inner);
        self.invoke_method("EndSession", &[])?;
        Ok(())
    }

    pub fn close_connection(&self) -> Result<(), anyhow::Error> {
        log::info!("close_connection: self address = {:p}, COM inner = {:p}", self, self.inner);
        self.invoke_method("CloseConnection", &[])?;
        Ok(())
    }

    pub fn process_request(&self, ticket: &str, request: &str) -> Result<String, anyhow::Error> {
        log::info!("process_request: self address = {:p}, COM inner = {:p}", self, self.inner);
        let ticket_var = SafeVariant::from_string(ticket);
        let request_var = SafeVariant::from_string(request);
        let result = self.invoke_method("DoRequests", &[ticket_var, request_var])?;
        Ok(result.to_string().unwrap_or_default())
    }

    pub fn get_current_company_file_name(&self) -> Result<String, anyhow::Error> {
        log::info!("get_current_company_file_name: self address = {:p}, COM inner = {:p}", self, self.inner);
        let result = self.invoke_method("GetCurrentCompanyFileName", &[])?;
        Ok(result.to_string().unwrap_or_default())
    }

    pub fn query_account_by_number(&self, session: *mut IDispatch, account_number: &str) -> Result<Option<AccountInfo>, anyhow::Error> {
        log::info!("query_account_by_number: self address = {:p}, COM inner = {:p}", self, self.inner);
        log::info!("Querying account by number using QBFC API: {}", account_number);
        // Step 1: Create AccountQuery object using QBFC API
        let query_result = self.invoke_method("CreateAccountQuery", &[SafeVariant::from_string(account_number)])?;
        let query_dispatch = query_result.to_dispatch().ok_or_else(|| anyhow::anyhow!("CreateAccountQuery did not return a dispatch pointer"))?;
        let query_var = SafeVariant::from_dispatch(Some(query_dispatch));
        let response_result = self.invoke_method("GetAccountResponse", &[query_var])?;
        let response_dispatch = response_result.to_dispatch().ok_or_else(|| anyhow::anyhow!("GetAccountResponse did not return a dispatch pointer"))?;
        log::debug!("✅ Created AccountQuery object");
        // Step 2: Set account number filter on the query
        let account_number_var = SafeVariant::from_string(account_number);
        self.invoke_method_on_dispatch(query_dispatch, "put_AccountNumber", &[account_number_var])?;
        log::debug!("✅ Set account number filter: {}", account_number);
        // Step 3: Execute the query
        let query_var = SafeVariant::from_dispatch(Some(query_dispatch));
        let response_result = self.invoke_method("GetAccountResponse", &[query_var])?;
        let response_dispatch = response_result.to_dispatch().ok_or_else(|| anyhow::anyhow!("GetAccountResponse did not return a dispatch pointer"))?;
        log::debug!("✅ Executed account query");
        // Step 4: Parse the response to extract account information
        self.parse_account_response(response_dispatch)
    }

    /// Helper method to invoke methods on IDispatch objects
    fn invoke_method_on_dispatch(&self, dispatch: *mut IDispatch, method_name: &str, params: &[SafeVariant]) -> Result<SafeVariant, anyhow::Error> {
        let mut dispid = 0i32;
        let method_name_wide = widestring::U16CString::from_str(method_name).unwrap();
        let names = [method_name_wide.as_ptr()];
        let hr = unsafe {
            ((*(*dispatch).lpVtbl).GetIDsOfNames)(
                dispatch,
                &IID_NULL,
                names.as_ptr() as *mut _,
                1,
                0,
                &mut dispid
            )
        };
        if hr < 0 {
            return Err(anyhow::anyhow!("GetIDsOfNames on dispatch failed: HRESULT=0x{:08X}", hr));
        }
        let mut dispparams = crate::safe_variant::create_dispparams_safe(params);
        let mut result: VARIANT = unsafe { std::mem::zeroed() };
        let mut excepinfo: EXCEPINFO = unsafe { std::mem::zeroed() };
        let hr = unsafe {
            ((*(*dispatch).lpVtbl).Invoke)(
                dispatch,
                dispid,
                &IID_NULL,
                0,
                DISPATCH_METHOD,
                &mut dispparams,
                &mut result,
                &mut excepinfo,
                std::ptr::null_mut()
            )
        };
        if hr < 0 {
            return Err(anyhow::anyhow!("Invoke on dispatch failed: HRESULT=0x{:08X}", hr));
        }
        Ok(SafeVariant::from_winvariant(&result))
    }
    
    /// Parse account information from QBFC response
    fn parse_account_response(&self, response_dispatch: *mut IDispatch) -> Result<Option<AccountInfo>, anyhow::Error> {
        log::info!("Parsing QBFC account response...");
        // Get response list (accounts)
        let response_list_result = self.invoke_method_on_dispatch(response_dispatch, "get_ResponseList", &[])?;
        let response_list = response_list_result.to_dispatch().ok_or_else(|| anyhow::anyhow!("get_ResponseList did not return a dispatch pointer"))?;
        // Get count of accounts returned
        let count_result = self.invoke_method_on_dispatch(response_list, "get_Count", &[])?;
        let count = count_result.to_i32().unwrap_or(0);
        log::debug!("Found {} account(s) in response", count);
        if count == 0 {
            log::warn!("No accounts found with the specified criteria");
            return Ok(None);
        }
        // Get first account (should be only one since we filtered by account number)
        let index_var = SafeVariant::from_i32(0); // 0-based index
        let account_result = self.invoke_method_on_dispatch(response_list, "GetAt", &[index_var])?;
        let account_dispatch = account_result.to_dispatch().ok_or_else(|| anyhow::anyhow!("GetAt did not return a dispatch pointer"))?;
        // Extract account details using QBFC API
        let name_result = self.invoke_method_on_dispatch(account_dispatch, "get_Name", &[])?;
        let name = name_result.to_string().unwrap_or_else(|| "Unknown".to_string());
        let number_result = self.invoke_method_on_dispatch(account_dispatch, "get_AccountNumber", &[])?;
        let number = number_result.to_string().unwrap_or_else(|| "Unknown".to_string());
        let type_result = self.invoke_method_on_dispatch(account_dispatch, "get_AccountType", &[])?;
        let account_type_enum = type_result.to_i32().unwrap_or(0);
        let account_type = self.qbfc_account_type_to_string(account_type_enum);
        let balance_result = self.invoke_method_on_dispatch(account_dispatch, "get_Balance", &[])?;
        let balance = balance_result.to_f64().unwrap_or(0.0);
        log::info!("✅ Successfully extracted account: {} ({}), Type: {}, Balance: \\${:.2}", name, number, account_type, balance);
        Ok(Some(AccountInfo {
            name,
            number,
            account_type,
            balance,
        }))
    }
    
    /// Convert QBFC account type enum to string
    fn qbfc_account_type_to_string(&self, account_type: i32) -> String {
        match account_type {
            0 => "AccountsPayable".to_string(),
            1 => "AccountsReceivable".to_string(),
            2 => "Bank".to_string(),
            3 => "CostOfGoodsSold".to_string(),
            4 => "CreditCard".to_string(),
            5 => "Equity".to_string(),
            6 => "Expense".to_string(),
            7 => "FixedAsset".to_string(),
            8 => "Income".to_string(),
            9 => "LongTermLiability".to_string(),
            10 => "OtherAsset".to_string(),
            11 => "OtherCurrentAsset".to_string(),
            12 => "OtherCurrentLiability".to_string(),
            13 => "OtherExpense".to_string(),
            14 => "OtherIncome".to_string(),
            _ => format!("Unknown({})", account_type),
        }
    }

}

impl Drop for RequestProcessor2 {
    fn drop(&mut self) {
        log::warn!("RequestProcessor2::drop called! self address = {:p}", self);
        // Attempt to close any open QuickBooks session and connection on drop
        // This is best-effort: log errors but do not panic
        log::info!("RequestProcessor2::drop: Attempting to clean up QuickBooks connection...");
        // Try to call CloseConnection (safe even if not open)
        let _ = self.invoke_method("CloseConnection", &[]).map_err(|e| {
            log::warn!("Drop: CloseConnection failed: {:#}", e);
        });
    }
}
