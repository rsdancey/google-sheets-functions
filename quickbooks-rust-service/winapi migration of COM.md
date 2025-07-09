# Migration Plan: COM Interop to winapi, Keep windows for Non-COM

## 1. Dependencies
- In `Cargo.toml`:
  - Keep the `windows` crate for non-COM APIs (e.g., registry, file, UI, etc.).
  - Ensure `winapi` is included with all necessary COM features:
    ```toml
    winapi = { version = "0.3", features = ["oleauto", "wtypes", "oaidl", "combaseapi"] }
    ```
  - Keep `variant-rs` for safe VARIANT handling.

## 2. Type Replacements
- In all COM-related modules (`request_processor.rs`, `account_service.rs`, `safe_variant.rs`):
  - Replace all `windows::Win32::System::Com::{IDispatch, VARIANT, ...}` with `winapi::um::oaidl::{IDispatch, VARIANT, ...}`.
  - Update all COM interface pointers, method calls, and type signatures to use `winapi` types.

## 3. COM Initialization and Object Creation
- Use `winapi` APIs for:
  - `CoInitializeEx` (`winapi::um::objbase::CoInitializeEx`)
  - `CoCreateInstance` (`winapi::um::combaseapi::CoCreateInstance`)
- Remove all `windows` crate COM object creation.

## 4. SafeVariant and VARIANT Handling
- Refactor `SafeVariant` to use only `variant-rs` and `winapi` types.
- All conversions, method calls, and memory management should use `variant-rs` and `winapi`.

## 5. IDispatch and Method Invocation
- Use `winapi`'s `IDispatch` interface and method signatures.
- Use `variant-rs` macros and helpers for property/method calls if desired.

## 6. Error Handling
- Replace `windows::core::Result` and error types with idiomatic Rust or `winapi`-style error handling (typically `HRESULT`).

## 7. Testing
- Thoroughly test all COM interop, especially object creation, method calls, and VARIANT conversions.

---

**As you migrate, keep all non-COM usage of the `windows` crate unchanged.**
