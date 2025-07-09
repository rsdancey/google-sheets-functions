# variant-rs Implementation Plan

## 1. Add Dependency
- Add `variant-rs` to your `Cargo.toml`.

## 2. Refactor SafeVariant
- Replace your current `SafeVariant` implementation with a thin wrapper around `variant_rs::Variant`.
- Implement constructors and extractors for all types you use (`i32`, `f64`, `String`, `Option<IDispatch>`, etc.) using `variant-rs` idioms.

## 3. Conversion Glue: Variant <-> VARIANT
- For every COM call that requires a `windows::Win32::System::Com::VARIANT`:
  - Convert `variant_rs::Variant` to `winapi::VARIANT` using `variant-rs`'s `TryInto`.
  - Then, unsafely transmute or manually copy the `winapi::VARIANT` to a `windows::Win32::System::Com::VARIANT` (they are layout-compatible, but not type-compatible).
  - For return values, convert back from `windows::Win32::System::Com::VARIANT` to `winapi::VARIANT`, then to `variant_rs::Variant`.

## 4. Update All Usage Sites
- Update all code that creates, passes, or extracts VARIANTs to use the new `SafeVariant` and conversion glue.
- This includes all method calls in `request_processor.rs`, `account_service.rs`, and any other COM interop code.

## 5. IDispatch and BSTR Handling
- Ensure that `Option<IDispatch>` and `BSTR` are handled using the types expected by `variant-rs`.
- Add conversion glue if your code uses `windows` crate's `IDispatch` or `BSTR` types.

## 6. Testing
- Write or update tests to ensure that all conversions and COM calls work as expected.
- Pay special attention to memory safety and ownership (especially for BSTR and IDispatch).

## 7. Documentation and Comments
- Document all conversion glue code, explaining why the conversion is safe and how it works.
- Add comments to all places where unsafe code is required for type conversion.

---

## Summary Table of Required Changes

| Step | File(s) | Description |
|------|---------|-------------|
| 1    | Cargo.toml | Add `variant-rs` dependency |
| 2    | safe_variant.rs | Refactor to use `variant_rs::Variant` |
| 3    | safe_variant.rs, request_processor.rs, account_service.rs | Implement conversion glue for every COM call |
| 4    | All usage sites | Update to use new SafeVariant and conversions |
| 5    | safe_variant.rs, request_processor.rs | Ensure BSTR/IDispatch compatibility |
| 6    | tests/ | Add/modify tests for new logic |
| 7    | All | Add documentation and comments for glue code |
