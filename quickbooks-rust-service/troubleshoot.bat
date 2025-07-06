@echo off
echo QuickBooks Rust Service Troubleshooting
echo =======================================

echo.
echo 1. Testing with Mock Data (no QuickBooks required)...
copy config\config.mock.toml config\config.toml
cargo run --release -- --test-connection --skip-sheets-test

echo.
echo 2. If the above worked, the issue is QuickBooks SDK installation
echo    To fix QuickBooks connection:
echo    - Install QuickBooks Desktop SDK
echo    - Run QuickBooks as Administrator  
echo    - Run this script as Administrator

echo.
echo 3. To test Google Sheets (requires proper webapp_url and api_key):
echo    - Edit config\config.toml with your Google Apps Script details
echo    - Run: cargo run --release -- --test-connection

echo.
echo 4. For production use:
echo    - Install QuickBooks SDK
echo    - Set company_file = "AUTO" or specific file path
echo    - Configure Google Sheets properly
echo    - Run: cargo run --release

pause
