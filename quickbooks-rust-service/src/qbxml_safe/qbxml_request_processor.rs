// Type-safe wrapper for QBXMLRP2.RequestProcessor COM API
// Mirrors the structure of request_processor.rs but uses tickets (strings) instead of pointers




use winapi::shared::guiddef::{CLSID, IID_NULL};
use winapi::um::oaidl::{IDispatch, VARIANT, EXCEPINFO};
use crate::qbxml_safe::qbxml_safe_variant::SafeVariant;
use crate::file_mode::FileMode;

const DISPATCH_METHOD: u16 = 1;

pub struct QbxmlRequestProcessor {
    inner: *mut IDispatch,
}

// Locally define IID_IDispatch for use in CoCreateInstance
#[allow(non_upper_case_globals)]
pub const IID_IDispatch: winapi::shared::guiddef::GUID = winapi::shared::guiddef::GUID {
    Data1: 0x00020400,
    Data2: 0x0000,
    Data3: 0x0000,
    Data4: [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46],
};

impl QbxmlRequestProcessor {
    pub fn new() -> Result<Self, anyhow::Error> {
        // Use the single QBXML ProgID for RequestProcessor
        let prog_id = "QBXMLRP2.RequestProcessor";
        log::info!("Trying QBXML ProgID: {}", prog_id);
        let prog_id_wide = widestring::U16CString::from_str(prog_id).unwrap();
        let mut clsid: CLSID = unsafe { std::mem::zeroed() };
        let hr = unsafe {
            winapi::um::combaseapi::CLSIDFromProgID(
                prog_id_wide.as_ptr(),
                &mut clsid as *mut CLSID
            )
        };
        if hr < 0 {
            log::error!("ProgID {} not found or CLSIDFromProgID failed: HRESULT=0x{:08X}", prog_id, hr as u32);
            return Err(anyhow::anyhow!("Failed to find QBXML COM ProgID: {} (HRESULT=0x{:08X})", prog_id, hr as u32));
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
            log::info!("✅ Successfully created QBXML COM instance with ProgID: {}", prog_id);
            let instance = Self {
                inner: dispatch_ptr,
            };
            log::info!("RequestProcessor2::new: instance address = {:p}, COM inner = {:p}", &instance, instance.inner);
            Ok(instance)
        } else {
            log::error!("Failed to create COM instance for {}: HRESULT=0x{:08X}", prog_id, hr as u32);
            Err(anyhow::anyhow!("Failed to create QBXML COM instance for ProgID: {} (HRESULT=0x{:08X})", prog_id, hr as u32))
        }
    }

    pub fn open_connection(&self, _app_id: &str, app_name: &str) -> Result<(), anyhow::Error> {
        log::info!("open_connection: self address = {:p}, COM inner = {:p}", self, self.inner);
        // Always pass empty string for AppID to avoid accidental registration (QBXML does not use AppID)
        let app_id_var = SafeVariant::from_string("");
        let app_name_var = SafeVariant::from_string(app_name);
        // Parameter order matches QBFC for consistency
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

    pub fn open_connection2(&self, app_id: &str, app_name: &str, _conn_type: i32) -> Result<(), anyhow::Error> {
        let app_id_var = SafeVariant::from_string(app_id);
        let app_name_var = SafeVariant::from_string(app_name);
        // Call OpenConnection with the parameters intentionally reversed which mirrors the code path in qbfc_request_processor.rs
        self.invoke_method("OpenConnection", &[app_name_var, app_id_var])?;
        Ok(())
    }

    pub fn begin_session(&self, company_file: &str, file_mode: FileMode) -> Result<String, anyhow::Error> {
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
        let result = self.invoke_method("BeginSession", &[mode_var, file_var])?;
        let vt = unsafe { result.as_variant().n1.n2().vt };
        let ticket = result.to_string().unwrap_or_default();
        log::info!("BeginSession returned VARIANT vt={} (expected {}), ticket='{}'", vt, winapi::shared::wtypes::VT_BSTR, ticket);
        if ticket.is_empty() {
            log::warn!("BeginSession returned an empty ticket string!");
        }
        Ok(ticket)
    }

    pub fn process_request(&self, ticket: &str, request: &str) -> Result<String, anyhow::Error> {
        let ticket_var = SafeVariant::from_string(ticket);
        let request_var = SafeVariant::from_string(request);
        let result = self.invoke_method("ProcessRequest", &[ticket_var, request_var])?;
        result.to_string().ok_or_else(|| anyhow::anyhow!("ProcessRequest did not return a string"))
    }

    pub fn end_session(&self, ticket: &str) -> Result<(), anyhow::Error> {
        log::info!("end_session: self address = {:p}, COM inner = {:p}, ticket = '{}'", self, self.inner, ticket);
        let ticket_var = SafeVariant::from_string(ticket);
        self.invoke_method("EndSession", &[ticket_var])?;
        log::info!("✅ EndSession successful for ticket '{}'", ticket);
        Ok(())
    }

    pub fn close_connection(&self) -> Result<(), anyhow::Error> {
        log::info!("close_connection: self address = {:p}, COM inner = {:p}", self, self.inner);
        self.invoke_method("CloseConnection", &[])?;
        log::info!("✅ CloseConnection successful");
        Ok(())
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
}
