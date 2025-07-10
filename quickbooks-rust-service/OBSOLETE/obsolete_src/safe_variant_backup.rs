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

impl SafeVariant {
    /// Create a new empty VARIANT
    pub fn new() -> Self {
        Self {
            inner: VARIANT::default(),
        }
    }
    
    /// Get the underlying VARIANT for COM calls
    pub fn as_variant(&self) -> &VARIANT {
        &self.inner
    }
}

/// Create DISPPARAMS from SafeVariant slice
pub fn create_dispparams_safe(_args: &[SafeVariant]) -> DISPPARAMS {
    DISPPARAMS {
        rgvarg: ptr::null_mut(),
        rgdispidNamedArgs: ptr::null_mut(),
        cArgs: 0,
        cNamedArgs: 0,
    }
}
