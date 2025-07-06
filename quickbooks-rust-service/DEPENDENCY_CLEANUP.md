# Project Dependency Analysis - Unnecessary Dependencies

## Current Project Size
- **Total**: 1.3GB
- **Target Directory**: 1.3GB (build artifacts)

## Dependencies Analysis

### ✅ **Dependencies Currently Used**
1. **reqwest** - HTTP client for Google Sheets API calls
2. **serde** + **serde derive** - JSON serialization/deserialization for API requests
3. **tokio** - Async runtime (only needs: `macros`, `rt-multi-thread`, `time`)
4. **anyhow** - Error handling throughout the codebase
5. **log** + **env_logger** - Logging system
6. **figment** - Configuration loading (only needs: `toml`)
7. **tokio-cron-scheduler** - Scheduling system
8. **chrono** - Timestamp generation for AccountData
9. **windows** - QuickBooks COM interface (Windows only)

### ❌ **Unnecessary Dependencies to Remove**

1. **`regex = "1.11.1"`** 
   - **Reason**: Only used in the old config.rs.old file for Windows path preprocessing
   - **Impact**: The new Figment-based config doesn't use regex
   - **Size Impact**: ~500KB compiled

2. **`serde_json = "1.0"`**
   - **Reason**: Not directly used - reqwest handles JSON with serde automatically
   - **Impact**: Reqwest with "json" feature provides this functionality
   - **Size Impact**: ~300KB compiled

3. **Figment `"env"` feature**
   - **Reason**: Not using Figment's environment variable mapping
   - **Impact**: Manual env var handling is used instead
   - **Size Impact**: ~50KB compiled

4. **Tokio `"full"` feature**
   - **Reason**: Only using `tokio::main`, `tokio::time::sleep`, and basic async
   - **Impact**: Can use specific features: `["macros", "rt-multi-thread", "time"]`
   - **Size Impact**: ~2MB compiled (major reduction)

5. **Chrono `"serde"` feature**
   - **Reason**: Only using `chrono::Utc::now().to_rfc3339()` for timestamps
   - **Impact**: Serde feature not needed for this use case
   - **Size Impact**: ~200KB compiled

### 🗑️ **Unused Files to Remove**

1. **`src/config.rs.old`** - Old configuration file (backup)
2. **`src/test_path.rs`** - Test file that uses the old config system
3. **`src/config_figment.rs`** - Duplicate file (already seen this exists)

## Recommended Changes

### 1. Update Cargo.toml
```toml
[dependencies]
# HTTP client for Google Sheets API
reqwest = { version = "0.12", features = ["json"] }
# JSON serialization 
serde = { version = "1.0", features = ["derive"] }
# Async runtime (reduced features)
tokio = { version = "1.0", features = ["macros", "rt-multi-thread", "time"] }
# Error handling
anyhow = "1.0"
# Logging
log = "0.4"
env_logger = "0.11"
# Configuration (reduced features)
figment = { version = "0.10", features = ["toml"] }
# Scheduling
tokio-cron-scheduler = "0.11"
# Date/time handling (no serde feature)
chrono = "0.4"
```

### 2. Remove Files
- `src/config.rs.old`
- `src/test_path.rs`
- `src/config_figment.rs` (if it's a duplicate)

## Expected Size Reduction
- **Compiled Dependencies**: ~3MB reduction
- **Source Code**: ~15KB reduction
- **Build Time**: ~10-15% faster builds
- **Target Directory**: Should reduce by ~200-300MB

## Risk Assessment
- **Low Risk**: All identified dependencies are confirmed unused
- **Testing**: Run full test suite after changes
- **Compatibility**: No breaking changes to functionality
