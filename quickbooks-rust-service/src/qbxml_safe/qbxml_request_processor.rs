// Type-safe wrapper for QBXMLRP2.RequestProcessor COM API
// Mirrors the structure of request_processor.rs but uses tickets (strings) instead of pointers

use winapi::shared::guiddef::{CLSID, IID_NULL};
use winapi::um::oaidl::{IDispatch, VARIANT, EXCEPINFO};
use crate::qbxml_safe::safe_variant::SafeVariant;
use quickbooks_sheets_sync::request_processor;

pub struct QbxmlRequestProcessor {
    inner: *mut IDispatch,
}

impl QbxmlRequestProcessor {
    pub fn new() -> Result<Self, anyhow::Error> {
        // Use QBXMLRP2.RequestProcessor ProgID
        let prog_id = "QBXMLRP2.RequestProcessor";
        let prog_id_wide = widestring::U16CString::from_str(prog_id).unwrap();
        let mut clsid: CLSID = unsafe { std::mem::zeroed() };
        let hr = unsafe {
            winapi::um::combaseapi::CLSIDFromProgID(
                prog_id_wide.as_ptr(),
                &mut clsid as *mut CLSID
            )
        };
        if hr < 0 {
            return Err(anyhow::anyhow!("CLSIDFromProgID failed: HRESULT=0x{:08X}", hr));
        }
        let mut dispatch_ptr: *mut IDispatch = std::ptr::null_mut();
        let hr = unsafe {
            winapi::um::combaseapi::CoCreateInstance(
                &clsid,
                std::ptr::null_mut(),
                winapi::shared::wtypesbase::CLSCTX_INPROC_SERVER,
                &request_processor::IID_IDispatch,
                &mut dispatch_ptr as *mut _ as *mut _
            )
        };
        if hr >= 0 && !dispatch_ptr.is_null() {
            Ok(Self { inner: dispatch_ptr })
        } else {
            Err(anyhow::anyhow!("CoCreateInstance failed: HRESULT=0x{:08X}", hr))
        }
    }

    pub fn open_connection2(&self, app_id: &str, app_name: &str, conn_type: i32) -> Result<(), anyhow::Error> {
        let app_id_var = SafeVariant::from_string(app_id);
        let app_name_var = SafeVariant::from_string(app_name);
        let conn_type_var = SafeVariant::from_i32(conn_type); // 1 = local QB
        self.invoke_method("OpenConnection2", &[app_id_var, app_name_var, conn_type_var])?;
        Ok(())
    }

    pub fn begin_session(&self, company_file: &str, mode: i32) -> Result<String, anyhow::Error> {
        let file_var = SafeVariant::from_string(company_file);
        let mode_var = SafeVariant::from_i32(mode);
        let result = self.invoke_method("BeginSession", &[file_var, mode_var])?;
        result.to_string().ok_or_else(|| anyhow::anyhow!("BeginSession did not return a string ticket"))
    }

    pub fn process_request(&self, ticket: &str, request: &str) -> Result<String, anyhow::Error> {
        let ticket_var = SafeVariant::from_string(ticket);
        let request_var = SafeVariant::from_string(request);
        let result = self.invoke_method("ProcessRequest", &[ticket_var, request_var])?;
        result.to_string().ok_or_else(|| anyhow::anyhow!("ProcessRequest did not return a string"))
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

    fn invoke_method(&self, method_name: &str, params: &[SafeVariant]) -> Result<SafeVariant, anyhow::Error> {
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
        let mut variants: Vec<VARIANT> = params.iter().map(|v| v.0).collect();
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
                1,
                &mut dispparams,
                &mut result,
                &mut excepinfo,
                &mut arg_err
            )
        };
        if hr < 0 {
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
        Ok(SafeVariant(result))
    }
}
