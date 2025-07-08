use std::ptr;
use windows::core::{BSTR, Result as WindowsResult};
use windows::Win32::System::Com::DISPPARAMS;
use windows::Win32::System::Variant::{VARIANT, VariantChangeType, VAR_CHANGE_FLAGS, VT_BSTR};

pub(crate) fn create_bstr_variant(s: &str) -> VARIANT {
    unsafe {
        let mut variant = VARIANT::default();
        let bstr = BSTR::from(s);
        (*variant.Anonymous.Anonymous).vt = VT_BSTR;
        let ptr = &mut (*variant.Anonymous.Anonymous).Anonymous as *mut _ as *mut BSTR;
        *ptr = bstr;
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

pub(crate) fn create_empty_dispparams() -> DISPPARAMS {
    DISPPARAMS {
        rgvarg: ptr::null_mut(),
        rgdispidNamedArgs: ptr::null_mut(),
        cArgs: 0,
        cNamedArgs: 0,
    }
}

pub(crate) fn variant_to_string(variant: &VARIANT) -> WindowsResult<String> {
    unsafe {
        // Check if already a BSTR
        if variant.Anonymous.Anonymous.vt == VT_BSTR {
            let bstr = &variant.Anonymous.Anonymous.Anonymous.bstrVal;
            return Ok(bstr.to_string());
        }
        
        // Try to convert to BSTR
        let mut dest = VARIANT::default();
        VariantChangeType(&mut dest, variant, VAR_CHANGE_FLAGS(0), VT_BSTR)?;
        
        let bstr = &dest.Anonymous.Anonymous.Anonymous.bstrVal;
        Ok(bstr.to_string())
    }
}
