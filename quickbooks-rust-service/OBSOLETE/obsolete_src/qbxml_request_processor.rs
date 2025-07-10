use winapi::shared::guiddef::{CLSID, IID_NULL};
use winapi::um::oaidl::{IDispatch, VARIANT, EXCEPINFO};
use crate::qbxml_safe::safe_variant::SafeVariant;

const DISPATCH_METHOD: u16 = 1;

#[allow(non_upper_case_globals)]
pub const IID_IDispatch: winapi::shared::guiddef::GUID = winapi::shared::guiddef::GUID {
    Data1: 0x00020400,
    Data2: 0x0000,
    Data3: 0x0000,
    Data4: [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46],
};

/// Type-safe wrapper for QBXML RequestProcessor
/// This uses the QBXML API (QBXMLRP2.RequestProcessor) and ticket-based session management
pub struct QbxmlRequestProcessor {
    inner: *mut IDispatch,
}

impl QbxmlRequestProcessor {
    pub fn new() -> Result<Self, anyhow::Error> {
        let prog_id_str = "QBXMLRP2.RequestProcessor";
        let prog_id_wide = widestring::U16CString::from_str(prog_id_str).unwrap();
        let mut clsid: CLSID = unsafe { std::mem::zeroed() };
        let hr = unsafe {
            winapi::um::combaseapi::CLSIDFromProgID(
                prog_id_wide.as_ptr(),
                &mut clsid as *mut CLSID
            )
        };
        if hr < 0 {
            return Err(anyhow::anyhow!("ProgID {} not found or CLSIDFromProgID failed: HRESULT=0x{:08X}", prog_id_str, hr as u32));
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
            Ok(Self { inner: dispatch_ptr })
        } else {
            Err(anyhow::anyhow!("Failed to create QBXML COM instance: HRESULT=0x{:08X}", hr as u32))
        }
    }

    fn invoke_method(&self, method_name: &str, params: &[SafeVariant]) -> Result<SafeVariant, anyhow::Error> {
        // Log parameter types and values for debugging
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
        let mut variants: Vec<winapi::um::oaidl::VARIANT> = params.iter().map(|v| v.as_variant().clone()).collect();
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
            println!(
                "COM Invoke failed: method={}, HRESULT=0x{:08X}, arg_err={},\n  EXCEPINFO: code={}, wCode={}, source='{}', description='{}', helpfile='{}', helpctx={}, scode=0x{:08X}",
                method_name,
                hr,
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
                "Invoke failed: method={}, HRESULT=0x{:08X}, description='{}', source='{}', helpfile='{}', scode=0x{:08X}",
                method_name,
                hr,
                description,
                source,
                helpfile,
                excepinfo.scode as u32
            ));
        }
        Ok(SafeVariant(result))
    }

    pub fn open_connection2(&self, app_id: &str, app_name: &str, conn_type: i32) -> Result<(), anyhow::Error> {
        let app_id_var = SafeVariant::from_string(app_id);
        let app_name_var = SafeVariant::from_string(app_name);
        let conn_type_var = SafeVariant::from_i32(conn_type);
        self.invoke_method("OpenConnection2", &[app_id_var, app_name_var, conn_type_var])?;
        Ok(())
    }

    pub fn begin_session(&self, company_file: &str, mode: i32) -> Result<String, anyhow::Error> {
        let file_var = SafeVariant::from_string(company_file);
        let mode_var = SafeVariant::from_i32(mode);
        let result = self.invoke_method("BeginSession", &[file_var, mode_var])?;
        result.to_string().ok_or_else(|| anyhow::anyhow!("BeginSession did not return a ticket string"))
    }

    pub fn end_session(&self, ticket: &str) -> Result<(), anyhow::Error> {
        let ticket_var = SafeVariant::from_string(ticket);
        self.invoke_method("EndSession", &[ticket_var])?;
        Ok(())
    }

    pub fn close_connection(&self) -> Result<(), anyhow::Error> {
        self.invoke_method("CloseConnection", &[])?;
        Ok(())
    }

    pub fn process_request(&self, ticket: &str, request: &str) -> Result<String, anyhow::Error> {
        let ticket_var = SafeVariant::from_string(ticket);
        let request_var = SafeVariant::from_string(request);
        let result = self.invoke_method("ProcessRequest", &[ticket_var, request_var])?;
        result.to_string().ok_or_else(|| anyhow::anyhow!("ProcessRequest did not return a response string"))
    }
}

impl Drop for QbxmlRequestProcessor {
    fn drop(&mut self) {
        let _ = self.invoke_method("CloseConnection", &[]);
    }
}
