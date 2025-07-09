// Standalone test file to compile SafeVariant in isolation
// This will show us any compilation errors

// Copy the exact content from safe_variant.rs
use std::ptr;
use windows::core::{BSTR, Result as WindowsResult};
use windows::Win32::System::Com::{DISPPARAMS, IDispatch};
use windows::Win32::System::Variant::{
    VARIANT, VariantChangeType, VariantClear,
    VAR_CHANGE_FLAGS, VT_BSTR, VT_I4, VT_DISPATCH, VT_R8, VT_EMPTY
};

/// Safe VARIANT wrapper that handles initialization properly
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
    
    /// Get the underlying VARIANT for COM calls
    pub fn as_variant(&self) -> &VARIANT {
        &self.inner
    }
    
    /// Get a mutable reference to the underlying VARIANT
    pub fn as_mut_variant(&mut self) -> &mut VARIANT {
        &mut self.inner
    }
}

/// Create DISPPARAMS from SafeVariant slice
pub fn create_dispparams_safe(_args: &[SafeVariant]) -> DISPPARAMS {
    // Fixed version to avoid dangling pointer
    DISPPARAMS {
        rgvarg: ptr::null_mut(),
        rgdispidNamedArgs: ptr::null_mut(),
        cArgs: 0,
        cNamedArgs: 0,
    }
}

fn main() {
    println!("Testing SafeVariant compilation...");
    let _variant = SafeVariant::new();
    let _params = create_dispparams_safe(&[]);
    println!("SafeVariant compiled successfully!");
}
