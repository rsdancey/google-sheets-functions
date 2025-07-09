use std::ptr;
use winapi::um::oaidl::{IDispatch, VARIANT, DISPPARAMS};
use winapi::um::oleauto::{VariantInit, SysAllocString, SysStringLen};
use winapi::shared::wtypes::{BSTR, VT_I4, VT_R8, VT_BSTR, VT_DISPATCH, VT_EMPTY};
use widestring::U16CString;
use std::mem::zeroed;

#[derive(Clone)]
pub struct SafeVariant {
    inner: VARIANT,
}

impl SafeVariant {
    pub fn new() -> Self {
        let mut var: VARIANT = unsafe { zeroed() };
        unsafe { VariantInit(&mut var) };
        Self { inner: var }
    }
    pub fn from_i32(value: i32) -> Self {
        let mut var: VARIANT = unsafe { zeroed() };
        unsafe {
            VariantInit(&mut var);
            *var.n1.n2_mut().n3.lVal_mut() = value;
            var.n1.n2_mut().vt = VT_I4 as u16;
        }
        Self { inner: var }
    }
    pub fn from_f64(value: f64) -> Self {
        let mut var: VARIANT = unsafe { zeroed() };
        unsafe {
            VariantInit(&mut var);
            *var.n1.n2_mut().n3.dblVal_mut() = value;
            var.n1.n2_mut().vt = VT_R8 as u16;
        }
        Self { inner: var }
    }
    pub fn from_string(s: &str) -> Self {
        let mut var: VARIANT = unsafe { zeroed() };
        unsafe {
            VariantInit(&mut var);
            let wide = U16CString::from_str(s).unwrap();
            let bstr: BSTR = SysAllocString(wide.as_ptr());
            *var.n1.n2_mut().n3.bstrVal_mut() = bstr;
            var.n1.n2_mut().vt = VT_BSTR as u16;
            // Log the BSTR pointer, VT, and the string value for verification
            println!("[SafeVariant::from_string] s='{}' (len={}) bstr={:p} vt={}", s, s.len(), bstr, var.n1.n2().vt);
            // Also, print the first few UTF-16 code units for extra verification
            let utf16: Vec<u16> = s.encode_utf16().collect();
            println!("[SafeVariant::from_string] utf16={:?}", &utf16[..std::cmp::min(utf16.len(), 8)]);
        }
        Self { inner: var }
    }
    pub fn from_dispatch(dispatch: Option<*mut IDispatch>) -> Self {
        let mut var: VARIANT = unsafe { zeroed() };
        unsafe {
            VariantInit(&mut var);
            if let Some(ptr) = dispatch {
                *var.n1.n2_mut().n3.pdispVal_mut() = ptr;
                var.n1.n2_mut().vt = VT_DISPATCH as u16;
            } else {
                var.n1.n2_mut().vt = VT_EMPTY as u16;
            }
        }
        Self { inner: var }
    }
    pub fn to_i32(&self) -> Option<i32> {
        unsafe {
            if self.inner.n1.n2().vt as u32 == VT_I4 {
                Some(*self.inner.n1.n2().n3.lVal())
            } else {
                None
            }
        }
    }
    pub fn to_f64(&self) -> Option<f64> {
        unsafe {
            if self.inner.n1.n2().vt as u32 == VT_R8 {
                Some(*self.inner.n1.n2().n3.dblVal())
            } else {
                None
            }
        }
    }
    pub fn to_string(&self) -> Option<String> {
        unsafe {
            if self.inner.n1.n2().vt as u32 == VT_BSTR {
                let bstr = *self.inner.n1.n2().n3.bstrVal();
                if bstr.is_null() {
                    return None;
                }
                let len = SysStringLen(bstr) as usize;
                let slice = std::slice::from_raw_parts(bstr as *const u16, len);
                Some(String::from_utf16_lossy(slice))
            } else {
                None
            }
        }
    }
    pub fn to_dispatch(&self) -> Option<*mut IDispatch> {
        unsafe {
            if self.inner.n1.n2().vt as u32 == VT_DISPATCH {
                let ptr = *self.inner.n1.n2().n3.pdispVal();
                if ptr.is_null() {
                    None
                } else {
                    Some(ptr)
                }
            } else {
                None
            }
        }
    }
    pub fn as_variant(&self) -> &VARIANT {
        &self.inner
    }
    pub fn to_winvariant(&self) -> VARIANT {
        self.inner.clone()
    }
    pub fn from_winvariant(win: &VARIANT) -> Self {
        Self { inner: win.clone() }
    }
    pub fn as_dispatch(&self) -> Result<*mut IDispatch, anyhow::Error> {
        unsafe {
            if self.inner.n1.n2().vt as u32 == VT_DISPATCH {
                let ptr = *self.inner.n1.n2().n3.pdispVal();
                if !ptr.is_null() {
                    Ok(ptr)
                } else {
                    log::error!("SafeVariant: pdispVal is null (VARIANT vt={})", self.inner.n1.n2().vt);
                    Err(anyhow::anyhow!("SafeVariant: pdispVal is null (VARIANT vt={})", self.inner.n1.n2().vt))
                }
            } else {
                log::error!("SafeVariant: not a VT_DISPATCH (vt={})", self.inner.n1.n2().vt);
                Err(anyhow::anyhow!("SafeVariant: not a VT_DISPATCH (vt={})", self.inner.n1.n2().vt))
            }
        }
    }
}

pub fn create_dispparams_safe(args: &[SafeVariant]) -> DISPPARAMS {
    let mut variants: Vec<VARIANT> = args.iter().map(|v| v.to_winvariant()).collect();
    DISPPARAMS {
        rgvarg: if variants.is_empty() { ptr::null_mut() } else { variants.as_mut_ptr() },
        rgdispidNamedArgs: ptr::null_mut(),
        cArgs: variants.len() as u32,
        cNamedArgs: 0,
    }
}
