use std::ptr;
use windows::core::{BSTR, Result as WindowsResult, Interface};
use windows::Win32::System::Com::{DISPPARAMS, IDispatch};
use windows::Win32::System::Variant::{
    VARIANT, VariantChangeType, VariantClear, 
    VAR_CHANGE_FLAGS, VT_BSTR, VT_I4, VT_DISPATCH, VT_R8, VARENUM
};

/// Create a VARIANT containing an i32 value
pub(crate) fn create_variant_i4(value: i32) -> VARIANT {
    unsafe {
        let mut variant = VARIANT::default();
        variant.Anonymous.Anonymous.vt = VT_I4;
        variant.Anonymous.Anonymous.Anonymous.lVal = value;
        variant
    }
}

/// Create a VARIANT containing a BSTR value
pub(crate) fn create_bstr_variant(s: &str) -> VARIANT {
    unsafe {
        let mut variant = VARIANT::default();
        let bstr = BSTR::from(s);
        variant.Anonymous.Anonymous.vt = VT_BSTR;
        variant.Anonymous.Anonymous.Anonymous.bstrVal = std::mem::transmute(bstr.as_ptr());
        std::mem::forget(bstr); // Prevent double free
        variant
    }
}

/// Create a VARIANT containing an IDispatch
pub(crate) fn create_dispatch_variant(dispatch: &IDispatch) -> VARIANT {
    unsafe {
        let mut variant = VARIANT::default();
        variant.Anonymous.Anonymous.vt = VT_DISPATCH;
        variant.Anonymous.Anonymous.Anonymous.pdispVal = std::mem::transmute(dispatch.as_raw());
        variant
    }
}

pub(crate) fn create_dispparams(args: &[VARIANT]) -> DISPPARAMS {
    DISPPARAMS {
        rgvarg: if args.is_empty() { ptr::null_mut() } else { args.as_ptr() as *mut _ },
        rgdispidNamedArgs: ptr::null_mut(),
        cArgs: args.len() as u32,
        cNamedArgs: 0,
    }
}

pub(crate) fn create_dispparams_from_slice(args: &[VARIANT]) -> DISPPARAMS {
    DISPPARAMS {
        rgvarg: if args.is_empty() { ptr::null_mut() } else { args.as_ptr() as *mut _ },
        rgdispidNamedArgs: ptr::null_mut(),
        cArgs: args.len() as u32,
        cNamedArgs: 0,
    }
}

pub(crate) fn create_empty_dispparams() -> DISPPARAMS {
    DISPPARAMS {
        rgvarg: ptr::null_mut(),
        rgdispidNamedArgs: ptr::null_mut(),
        cArgs: 0,
        cNamedArgs: 0,
    }
}

/// Convert VARIANT to string using Windows API
pub(crate) fn variant_to_string(variant: &VARIANT) -> WindowsResult<String> {
    unsafe {
        let mut dest = VARIANT::default();
        match VariantChangeType(&mut dest, variant, VAR_CHANGE_FLAGS(0), VT_BSTR) {
            Ok(()) => {
                let bstr_ptr = dest.Anonymous.Anonymous.Anonymous.bstrVal;
                let result = if !bstr_ptr.is_null() {
                    BSTR::from_raw(bstr_ptr).to_string()
                } else {
                    String::new()
                };
                VariantClear(&mut dest);
                Ok(result)
            }
            Err(e) => {
                VariantClear(&mut dest);
                Err(e)
            }
        }
    }
}

/// Extract IDispatch from VARIANT
pub(crate) fn variant_to_dispatch(variant: &VARIANT) -> WindowsResult<Option<IDispatch>> {
    unsafe {
        if variant.Anonymous.Anonymous.vt == VT_DISPATCH {
            let dispatch_ptr = variant.Anonymous.Anonymous.Anonymous.pdispVal;
            if !dispatch_ptr.is_null() {
                Ok(Some(IDispatch::from_raw(dispatch_ptr)))
            } else {
                Ok(None)
            }
        } else {
            Ok(None)
        }
    }
}

/// Extract f64 from VARIANT
pub(crate) fn variant_to_f64(variant: &VARIANT) -> WindowsResult<Option<f64>> {
    unsafe {
        match variant.Anonymous.Anonymous.vt {
            VT_R8 => Ok(Some(variant.Anonymous.Anonymous.Anonymous.dblVal)),
            _ => {
                let mut dest = VARIANT::default();
                match VariantChangeType(&mut dest, variant, VAR_CHANGE_FLAGS(0), VT_R8) {
                    Ok(()) => {
                        let result = dest.Anonymous.Anonymous.Anonymous.dblVal;
                        VariantClear(&mut dest);
                        Ok(Some(result))
                    }
                    Err(_) => Ok(None),
                }
            }
        }
    }
}

/// Extract i32 from VARIANT
pub(crate) fn variant_to_i32(variant: &VARIANT) -> WindowsResult<Option<i32>> {
    unsafe {
        match variant.Anonymous.Anonymous.vt {
            VT_I4 => Ok(Some(variant.Anonymous.Anonymous.Anonymous.lVal)),
            _ => {
                let mut dest = VARIANT::default();
                match VariantChangeType(&mut dest, variant, VAR_CHANGE_FLAGS(0), VT_I4) {
                    Ok(()) => {
                        let result = dest.Anonymous.Anonymous.Anonymous.lVal;
                        VariantClear(&mut dest);
                        Ok(Some(result))
                    }
                    Err(_) => Ok(None),
                }
            }
        }
    }
}
