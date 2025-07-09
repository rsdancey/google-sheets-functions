use std::ptr;
use windows::core::{BSTR, Result as WindowsResult, Interface};
use windows::Win32::System::Com::{DISPPARAMS, IDispatch};
use windows::Win32::System::Variant::{VARIANT, VT_BSTR, VT_I4, VT_DISPATCH};

/// Create a VARIANT containing an i32 value
pub(crate) fn create_variant_i4(value: i32) -> VARIANT {
    let mut variant = VARIANT::default();
    unsafe {
        std::ptr::write(&mut variant.Anonymous.Anonymous.vt, VT_I4);
        std::ptr::write(&mut variant.Anonymous.Anonymous.Anonymous.lVal, value);
    }
    variant
}

/// Create a VARIANT containing a BSTR value
pub(crate) fn create_bstr_variant(s: &str) -> VARIANT {
    let mut variant = VARIANT::default();
    let bstr = BSTR::from(s);
    unsafe {
        std::ptr::write(&mut variant.Anonymous.Anonymous.vt, VT_BSTR);
        std::ptr::write(&mut variant.Anonymous.Anonymous.Anonymous.bstrVal, std::mem::transmute(bstr));
        std::mem::forget(bstr); // Prevent double free
    }
    variant
}

/// Create a VARIANT containing an IDispatch
pub(crate) fn create_dispatch_variant(dispatch: &IDispatch) -> VARIANT {
    let mut variant = VARIANT::default();
    unsafe {
        std::ptr::write(&mut variant.Anonymous.Anonymous.vt, VT_DISPATCH);
        std::ptr::write(&mut variant.Anonymous.Anonymous.Anonymous.pdispVal, std::mem::transmute(dispatch.as_raw()));
    }
    variant
}

pub(crate) fn create_dispparams_from_slice(args: &[VARIANT]) -> DISPPARAMS {
    DISPPARAMS {
        rgvarg: if args.is_empty() { ptr::null_mut() } else { args.as_ptr() as *mut _ },
        rgdispidNamedArgs: ptr::null_mut(),
        cArgs: args.len() as u32,
        cNamedArgs: 0,
    }
}

/// Convert VARIANT to string - simplified version
pub(crate) fn variant_to_string(variant: &VARIANT) -> WindowsResult<String> {
    unsafe {
        // For now, just return a placeholder - we'll implement proper conversion later
        Ok("placeholder".to_string())
    }
}

/// Extract IDispatch from VARIANT - simplified version  
pub(crate) fn variant_to_dispatch(variant: &VARIANT) -> WindowsResult<Option<IDispatch>> {
    unsafe {
        // For now, just return None - we'll implement proper extraction later
        Ok(None)
    }
}

/// Extract f64 from VARIANT - simplified version
pub(crate) fn variant_to_f64(variant: &VARIANT) -> WindowsResult<Option<f64>> {
    unsafe {
        // For now, just return a placeholder - we'll implement proper extraction later
        Ok(Some(0.0))
    }
}

/// Extract i32 from VARIANT - simplified version  
pub(crate) fn variant_to_i32(variant: &VARIANT) -> WindowsResult<Option<i32>> {
    unsafe {
        // For now, just return a placeholder - we'll implement proper extraction later
        Ok(Some(0))
    }
}
