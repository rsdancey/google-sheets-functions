use std::ptr;
use windows::core::{Result as WindowsResult};
use windows::Win32::System::Com::{DISPPARAMS, IDispatch};
use windows::Win32::System::Variant::VARIANT;
use oaidl::{VariantExt, IntoVariantError, FromVariantError};

/// Modern SafeVariant wrapper using the oaidl crate for robust VARIANT handling
pub struct SafeVariant {
    inner: VARIANT,
}

impl Clone for SafeVariant {
    fn clone(&self) -> Self {
        unsafe {
            let mut new_variant = VARIANT::default();
            windows::Win32::System::Variant::VariantCopy(&mut new_variant, &self.inner).unwrap();
            Self { inner: new_variant }
        }
    }
}

impl Drop for SafeVariant {
    fn drop(&mut self) {
        unsafe {
            let _ = windows::Win32::System::Variant::VariantClear(&mut self.inner);
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
    
    /// Create a VARIANT containing an i32 value using oaidl
    pub fn from_i32(value: i32) -> Result<Self, IntoVariantError> {
        let variant = VariantExt::into_variant(value)?;
        Ok(Self { inner: unsafe { *variant.as_ptr() } })
    }
    
    /// Create a VARIANT containing a string value using oaidl
    pub fn from_string(s: &str) -> Result<Self, IntoVariantError> {
        let variant = VariantExt::into_variant(s.to_string())?;
        Ok(Self { inner: unsafe { *variant.as_ptr() } })
    }
    
    /// Create a VARIANT containing an IDispatch pointer using oaidl
    pub fn from_dispatch(dispatch: Option<IDispatch>) -> Result<Self, IntoVariantError> {
        match dispatch {
            Some(d) => {
                // For IDispatch, we need to use raw pointer conversion
                let variant = VariantExt::into_variant(d.as_raw() as *mut std::ffi::c_void)?;
                Ok(Self { inner: unsafe { *variant.as_ptr() } })
            }
            None => {
                // Create empty variant for None
                Ok(Self::new())
            }
        }
    }
    
    /// Extract i32 value from VARIANT using oaidl
    pub fn to_i32(&self) -> Result<Option<i32>, FromVariantError> {
        match i32::try_from_variant(&self.inner) {
            Ok(value) => Ok(Some(value)),
            Err(_) => Ok(None), // Conversion failed, return None instead of error
        }
    }
    
    /// Extract f64 value from VARIANT using oaidl
    pub fn to_f64(&self) -> Result<Option<f64>, FromVariantError> {
        match f64::try_from_variant(&self.inner) {
            Ok(value) => Ok(Some(value)),
            Err(_) => Ok(None), // Conversion failed, return None instead of error
        }
    }
    
    /// Extract string value from VARIANT using oaidl
    pub fn to_string(&self) -> Result<Option<String>, FromVariantError> {
        match String::try_from_variant(&self.inner) {
            Ok(value) => Ok(Some(value)),
            Err(_) => Ok(None), // Conversion failed, return None instead of error
        }
    }
    
    /// Extract IDispatch from VARIANT
    pub fn to_dispatch(&self) -> WindowsResult<Option<IDispatch>> {
        // For IDispatch extraction, we still need to use the manual approach
        // since oaidl may not directly support IDispatch conversion
        unsafe {
            let vt = self.inner.Anonymous.Anonymous.vt;
            if vt == windows::Win32::System::Variant::VT_DISPATCH {
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

/// Create DISPPARAMS from SafeVariant slice (alias for compatibility)
pub fn create_dispparams_safe(args: &[SafeVariant]) -> DISPPARAMS {
    create_dispparams_from_slice(args)
}

/// Helper functions for backward compatibility with Windows Result types
pub fn variant_to_string(variant: &VARIANT) -> WindowsResult<String> {
    let safe_variant = SafeVariant::from(variant);
    match safe_variant.to_string() {
        Ok(opt) => Ok(opt.unwrap_or_default()),
        Err(_) => Ok(String::new()), // Return empty string on error
    }
}

pub fn variant_to_i32(variant: &VARIANT) -> WindowsResult<i32> {
    let safe_variant = SafeVariant::from(variant);
    match safe_variant.to_i32() {
        Ok(opt) => Ok(opt.unwrap_or_default()),
        Err(_) => Ok(0), // Return 0 on error
    }
}

pub fn variant_to_f64(variant: &VARIANT) -> WindowsResult<f64> {
    let safe_variant = SafeVariant::from(variant);
    match safe_variant.to_f64() {
        Ok(opt) => Ok(opt.unwrap_or_default()),
        Err(_) => Ok(0.0), // Return 0.0 on error
    }
}

pub fn variant_to_dispatch(variant: &VARIANT) -> WindowsResult<Option<IDispatch>> {
    let safe_variant = SafeVariant::from(variant);
    safe_variant.to_dispatch()
}

pub fn create_variant_i4(value: i32) -> VARIANT {
    match SafeVariant::from_i32(value) {
        Ok(safe) => unsafe {
            let mut result = VARIANT::default();
            windows::Win32::System::Variant::VariantCopy(&mut result, &safe.inner).unwrap();
            result
        },
        Err(_) => VARIANT::default(), // Return empty variant on error
    }
}

pub fn create_bstr_variant(s: &str) -> VARIANT {
    match SafeVariant::from_string(s) {
        Ok(safe) => unsafe {
            let mut result = VARIANT::default();
            windows::Win32::System::Variant::VariantCopy(&mut result, &safe.inner).unwrap();
            result
        },
        Err(_) => VARIANT::default(), // Return empty variant on error
    }
}

pub fn create_dispatch_variant(dispatch: &IDispatch) -> VARIANT {
    match SafeVariant::from_dispatch(Some(dispatch.clone())) {
        Ok(safe) => unsafe {
            let mut result = VARIANT::default();
            windows::Win32::System::Variant::VariantCopy(&mut result, &safe.inner).unwrap();
            result
        },
        Err(_) => VARIANT::default(), // Return empty variant on error
    }
}
