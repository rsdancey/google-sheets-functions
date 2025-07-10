use windows::Win32::System::Variant::VARIANT;

pub struct SafeVariant {
    inner: VARIANT,
}

impl SafeVariant {
    pub fn new() -> Self {
        Self {
            inner: VARIANT::default(),
        }
    }
}

pub fn create_dispparams_safe(_args: &[SafeVariant]) -> windows::Win32::System::Com::DISPPARAMS {
    windows::Win32::System::Com::DISPPARAMS {
        rgvarg: std::ptr::null_mut(),
        rgdispidNamedArgs: std::ptr::null_mut(),
        cArgs: 0,
        cNamedArgs: 0,
    }
}
