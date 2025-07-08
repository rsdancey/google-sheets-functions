use std::mem::ManuallyDrop;
use windows::core::BSTR;
use windows::Win32::System::Com::DISPPARAMS;
use windows::Win32::System::Variant::{VARIANT, VARENUM, VT_BSTR};

pub(crate) fn create_bstr_variant(s: &str) -> VARIANT {
    let bstr = ManuallyDrop::new(BSTR::from(s));
    unsafe {
        let mut variant = VARIANT::default();
        let variant_anon = &mut variant.Anonymous;
        let variant_union = &mut variant_anon.Anonymous;
        
        variant_union.vt = VARENUM(VT_BSTR.0);
        variant_union.Anonymous.bstrVal = bstr;
        
        variant
    }
}

pub(crate) fn create_dispparams(args: &[VARIANT]) -> DISPPARAMS {
    DISPPARAMS {
        rgvarg: if args.is_empty() { std::ptr::null_mut() } else { args.as_ptr() as *mut _ },
        rgdispidNamedArgs: std::ptr::null_mut(),
        cArgs: args.len() as u32,
        cNamedArgs: 0,
    }
}

pub(crate) fn create_empty_dispparams() -> DISPPARAMS {
    DISPPARAMS {
        rgvarg: std::ptr::null_mut(),
        rgdispidNamedArgs: std::ptr::null_mut(),
        cArgs: 0,
        cNamedArgs: 0,
    }
}

pub(crate) fn variant_to_string(variant: &VARIANT) -> windows::core::Result<String> {
    unsafe {
        let variant_anon = &variant.Anonymous;
        let variant_union = &variant_anon.Anonymous;

        if variant_union.vt == VARENUM(VT_BSTR.0) {
            let bstr = &variant_union.Anonymous.bstrVal;
            return Ok(bstr.to_string());
        }
        Err(windows::core::Error::new(
            windows::core::HRESULT(-1),
            "VARIANT is not a BSTR".into()
        ))
    }
}
