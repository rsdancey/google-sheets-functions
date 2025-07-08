use windows::core::{HSTRING, IUnknown, PCWSTR};
use windows::Win32::System::Com::{
    CLSIDFromProgID, CoCreateInstance, CLSCTX_LOCAL_SERVER,
    IDispatch, DISPATCH_METHOD, EXCEPINFO,
};
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
    pub fn new() -> windows::core::Result<Self> {
        let prog_id = HSTRING::from("QBXMLRP2.RequestProcessor.2");
        let clsid = unsafe { CLSIDFromProgID(&prog_id)? };
        let dispatch: IDispatch = unsafe {
            CoCreateInstance::<Option<&IUnknown>, IDispatch>(
                &clsid,
                None,
                CLSCTX_LOCAL_SERVER
            )?
        };

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
