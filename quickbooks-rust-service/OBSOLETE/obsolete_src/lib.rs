// QuickBooks Desktop Sync Service Library
// Using SafeVariant wrappers for robust VARIANT/COM handling

pub mod config;
pub mod high_level_client;
pub mod safe_variant;

// COM-related modules now use SafeVariant for robust VARIANT handling
pub mod request_processor; 
pub mod account_service;

// Common types used across modules
#[derive(Debug, Clone, Copy)]
pub enum FileMode {
    DoNotCare = 0,
    SingleUser = 1,
    MultiUser = 2,
    Online = 3,
}

// Re-export commonly used types
pub use request_processor::AccountInfo;
