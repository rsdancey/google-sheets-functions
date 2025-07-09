// SafeVariant for QBXML API, similar to the one used for QBFC
// Provides type-safe wrappers for VARIANTs for COM interop

use winapi::um::oaidl::VARIANT;
use winapi::shared::wtypes::{VT_BSTR, VT_I4};
use widestring::U16CString;
use winapi::um::oleauto::{SysAllocStringLen, SysFreeString, SysStringLen};

pub struct SafeVariant(pub VARIANT);

impl SafeVariant {
    pub fn from_string(s: &str) -> Self {
        let wide = U16CString::from_str(s).unwrap();
        let bstr = unsafe { SysAllocStringLen(wide.as_ptr(), wide.len() as u32) };
        let mut var: VARIANT = unsafe { std::mem::zeroed() };
        unsafe {
            *var.n1.n2_mut().n3.bstrVal_mut() = bstr;
            var.n1.n2_mut().vt = VT_BSTR as u16;
        }
        SafeVariant(var)
    }
    pub fn from_i32(i: i32) -> Self {
        let mut var: VARIANT = unsafe { std::mem::zeroed() };
        unsafe {
            *var.n1.n2_mut().n3.lVal_mut() = i;
            var.n1.n2_mut().vt = VT_I4 as u16;
        }
        SafeVariant(var)
    }
    pub fn as_variant(&self) -> &VARIANT {
        &self.0
    }
    pub fn to_string(&self) -> Option<String> {
        let vt = unsafe { self.0.n1.n2().vt };
        if vt == VT_BSTR as u16 {
            let bstr = unsafe { *self.0.n1.n2().n3.bstrVal() };
            if bstr.is_null() {
                return None;
            }
            let len = unsafe { SysStringLen(bstr) } as usize;
            let slice = unsafe { std::slice::from_raw_parts(bstr, len) };
            Some(String::from_utf16_lossy(slice))
        } else {
            None
        }
    }
}

impl Drop for SafeVariant {
    fn drop(&mut self) {
        let vt = unsafe { self.0.n1.n2().vt };
        if vt == VT_BSTR as u16 {
            let bstr = unsafe { *self.0.n1.n2().n3.bstrVal() };
            if !bstr.is_null() {
                unsafe { SysFreeString(bstr) };
            }
        }
    }
}

// Add more helpers as needed for QBXML
