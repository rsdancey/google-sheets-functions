use std::ptr;
use windows::core::{BSTR, Result as WindowsResult};
use windows::Win32::System::Com::{DISPPARAMS, IDispatch};
use windows::Win32::System::Variant::{
    VARIANT, VariantChangeType, VariantClear,
    VAR_CHANGE_FLAGS, VT_BSTR, VT_I4, VT_DISPATCH, VT_R8
};

/// Safe VARIANT wrapper that handles the ManuallyDrop union issues
pub struct SafeVariant {
    inner: VARIANT,
}

impl Clone for SafeVariant {
    fn clone(&self) -> Self {
        unsafe {
            // Create a new variant and copy the contents
            let mut new_variant = VARIANT::default();
            windows::Win32::System::Variant::VariantCopy(&mut new_variant, &self.inner).unwrap();
            Self { inner: new_variant }
        }
    }
}

impl Drop for SafeVariant {
    fn drop(&mut self) {
        unsafe {
            let _ = VariantClear(&mut self.inner);
        }
    }
}

impl SafeVariant {
    /// Create a new empty VARIANT
    pub fn new() -> Self {
        Self {
            inner: VARIANT::default(),
        }
    }
    
    /// Create a VARIANT containing an i32 value
    pub fn from_i32(value: i32) -> Self {
        let mut variant = Self::new();
        unsafe {
            variant.inner.Anonymous.Anonymous.vt = VT_I4;
            variant.inner.Anonymous.Anonymous.Anonymous.lVal = value;
        }
        variant
    }
    
    /// Create a VARIANT containing a BSTR value
    pub fn from_string(s: &str) -> Self {
        let mut variant = Self::new();
        let bstr = BSTR::from(s);
        unsafe {
            variant.inner.Anonymous.Anonymous.vt = VT_BSTR;
            variant.inner.Anonymous.Anonymous.Anonymous.bstrVal = std::mem::ManuallyDrop::new(bstr);
        }
        variant
    }
    
    /// Create a VARIANT containing an IDispatch pointer
    pub fn from_dispatch(dispatch: Option<IDispatch>) -> Self {
        let mut variant = Self::new();
        unsafe {
            variant.inner.Anonymous.Anonymous.vt = VT_DISPATCH;
            variant.inner.Anonymous.Anonymous.Anonymous.pdispVal = std::mem::ManuallyDrop::new(dispatch);
        }
        variant
    }
    
    /// Extract i32 value from VARIANT
    pub fn to_i32(&self) -> WindowsResult<Option<i32>> {
        unsafe {
            let vt = self.inner.Anonymous.Anonymous.vt;
            if vt == VT_I4 {
                Ok(Some(self.inner.Anonymous.Anonymous.Anonymous.lVal))
            } else {
                // Try to convert to i32
                let mut result = VARIANT::default();
                match VariantChangeType(&mut result, &self.inner, VAR_CHANGE_FLAGS(0), VT_I4) {
                    Ok(_) => {
                        let value = result.Anonymous.Anonymous.Anonymous.lVal;
                        VariantClear(&mut result);
                        Ok(Some(value))
                    }
                    Err(_) => Ok(None),
                }
            }
        }
    }
    
    /// Extract f64 value from VARIANT
    pub fn to_f64(&self) -> WindowsResult<Option<f64>> {
        unsafe {
            let vt = self.inner.Anonymous.Anonymous.vt;
            if vt == VT_R8 {
                Ok(Some(self.inner.Anonymous.Anonymous.Anonymous.dblVal))
            } else {
                // Try to convert to f64
                let mut result = VARIANT::default();
                match VariantChangeType(&mut result, &self.inner, VAR_CHANGE_FLAGS(0), VT_R8) {
                    Ok(_) => {
                        let value = result.Anonymous.Anonymous.Anonymous.dblVal;
                        VariantClear(&mut result);
                        Ok(Some(value))
                    }
                    Err(_) => Ok(None),
                }
            }
        }
    }
    
    /// Extract string value from VARIANT
    pub fn to_string(&self) -> WindowsResult<Option<String>> {
        unsafe {
            let vt = self.inner.Anonymous.Anonymous.vt;
            if vt == VT_BSTR {
                let bstr_val = &self.inner.Anonymous.Anonymous.Anonymous.bstrVal;
                let result = if let Some(bstr) = bstr_val.as_ref() {
                    bstr.to_string()
                } else {
                    String::new()
                };
                Ok(Some(result))
            } else {
                // Try to convert to BSTR
                let mut result = VARIANT::default();
                match VariantChangeType(&mut result, &self.inner, VAR_CHANGE_FLAGS(0), VT_BSTR) {
                    Ok(_) => {
                        let bstr_val = &result.Anonymous.Anonymous.Anonymous.bstrVal;
                        let str_result = if let Some(bstr) = bstr_val.as_ref() {
                            bstr.to_string()
                        } else {
                            String::new()
                        };
                        VariantClear(&mut result);
                        Ok(Some(str_result))
                    }
                    Err(_) => Ok(None),
                }
            }
        }
    }
    
    /// Extract IDispatch from VARIANT
    pub fn to_dispatch(&self) -> WindowsResult<Option<IDispatch>> {
        unsafe {
            let vt = self.inner.Anonymous.Anonymous.vt;
            if vt == VT_DISPATCH {
                let dispatch_val = &self.inner.Anonymous.Anonymous.Anonymous.pdispVal;
                if let Some(dispatch) = dispatch_val.as_ref() {
                    Ok(Some(dispatch.clone()))
                } else {
                    Ok(None)
                }
            } else {
                Ok(None)
            }
        }
    }
    
    /// Get the underlying VARIANT for COM calls
    pub fn as_variant(&self) -> &VARIANT {
        &self.inner
    }
    
    /// Get a mutable reference to the underlying VARIANT
    pub fn as_mut_variant(&mut self) -> &mut VARIANT {
        &mut self.inner
    }
}

impl From<VARIANT> for SafeVariant {
    fn from(variant: VARIANT) -> Self {
        Self { inner: variant }
    }
}

impl From<&VARIANT> for SafeVariant {
    fn from(variant: &VARIANT) -> Self {
        unsafe {
            let mut new_variant = VARIANT::default();
            windows::Win32::System::Variant::VariantCopy(&mut new_variant, variant).unwrap();
            Self { inner: new_variant }
        }
    }
}

/// Create DISPPARAMS from SafeVariant slice
pub fn create_dispparams_from_slice(args: &[SafeVariant]) -> DISPPARAMS {
    let raw_variants: Vec<VARIANT> = args.iter().map(|v| unsafe { 
        let mut clone = VARIANT::default();
        windows::Win32::System::Variant::VariantCopy(&mut clone, &v.inner).unwrap();
        clone
    }).collect();
    
    DISPPARAMS {
        rgvarg: raw_variants.as_ptr() as *mut _,
        rgdispidNamedArgs: ptr::null_mut(),
        cArgs: raw_variants.len() as u32,
        cNamedArgs: 0,
    }
}

/// Create DISPPARAMS from SafeVariant slice (alias for backward compatibility)
pub fn create_dispparams_safe(args: &[SafeVariant]) -> DISPPARAMS {
    create_dispparams_from_slice(args)
}

/// Helper functions for backward compatibility
pub fn variant_to_string(variant: &VARIANT) -> WindowsResult<String> {
    let safe_variant = SafeVariant::from(variant);
    safe_variant.to_string().map(|opt| opt.unwrap_or_default())
}

pub fn variant_to_i32(variant: &VARIANT) -> WindowsResult<i32> {
    let safe_variant = SafeVariant::from(variant);
    safe_variant.to_i32().map(|opt| opt.unwrap_or_default())
}

pub fn variant_to_f64(variant: &VARIANT) -> WindowsResult<f64> {
    let safe_variant = SafeVariant::from(variant);
    safe_variant.to_f64().map(|opt| opt.unwrap_or_default())
}

pub fn variant_to_dispatch(variant: &VARIANT) -> WindowsResult<Option<IDispatch>> {
    let safe_variant = SafeVariant::from(variant);
    safe_variant.to_dispatch()
}

pub fn create_variant_i4(value: i32) -> VARIANT {
    let safe = SafeVariant::from_i32(value);
    unsafe {
        let mut result = VARIANT::default();
        windows::Win32::System::Variant::VariantCopy(&mut result, &safe.inner).unwrap();
        result
    }
}

pub fn create_bstr_variant(s: &str) -> VARIANT {
    let safe = SafeVariant::from_string(s);
    unsafe {
        let mut result = VARIANT::default();
        windows::Win32::System::Variant::VariantCopy(&mut result, &safe.inner).unwrap();
        result
    }
}

pub fn create_dispatch_variant(dispatch: &IDispatch) -> VARIANT {
    let safe = SafeVariant::from_dispatch(Some(dispatch.clone()));
    unsafe {
        let mut result = VARIANT::default();
        windows::Win32::System::Variant::VariantCopy(&mut result, &safe.inner).unwrap();
        result
    }
}
