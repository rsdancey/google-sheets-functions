use anyhow::Result;
use quickbooks_sheets_sync::{config::Config, high_level_client::SyncService};
use winapi::um::combaseapi::CoInitializeEx;
use winapi::um::objbase::COINIT_APARTMENTTHREADED;
use winapi::um::combaseapi::CoUninitialize;

#[cfg(feature = "qbxml")]
mod qbxml_safe;
#[cfg(feature = "qbxml")]
mod qbxml_request_processor;

fn main() -> Result<()> {
    println!("QuickBooks Account Sync Service v2");
    println!("==================================");
    println!();
    
    // Load configuration
    let config = Config::load()?;
    println!("✅ Configuration loaded from config.toml");

    // Initialize COM in STA mode for QuickBooks compatibility
    unsafe {
        let hr = CoInitializeEx(std::ptr::null_mut(), COINIT_APARTMENTTHREADED);
        if hr < 0 {
            panic!("CoInitializeEx failed: HRESULT=0x{:08X}", hr);
        }
    }

    #[cfg(feature = "QBXML")]
    {
        println!("Running in QBXML mode!");
        use qbxml_request_processor::QbxmlClient;
        let app_id = "";
        let app_name = "QuickBooks Sync Service";
        let company_file = "C:\\Path\\To\\Company.qbw";
        let mode = 1; // 1 = single user
        let request_xml = "<QBXML><QBXMLMsgsRq onError='stopOnError'><CompanyQueryRq/></QBXMLMsgsRq></QBXML>";
        match QbxmlClient::new(app_id, app_name, company_file, mode) {
            Ok(mut client) => {
                match client.process_request(request_xml) {
                    Ok(response) => println!("QBXML response: {}", response),
                    Err(e) => eprintln!("QBXML request error: {:#}", e),
                }
                let _ = client.end_session();
                let _ = client.close_connection();
            }
            Err(e) => eprintln!("QBXML init error: {:#}", e),
        }
        unsafe { CoUninitialize(); }
        return Ok(());
    }

    #[cfg(feature = "QBFC")]
    {
        println!("Running in QBXML mode!");
        use qbxml_request_processor::QbxmlClient;
        let app_id = "";
        let app_name = "Rust QBXML Test";
        let company_file = "C:\\Path\\To\\Company.qbw";
        let mode = 1; // 1 = single user
        let request_xml = "<QBXML><QBXMLMsgsRq onError='stopOnError'><CompanyQueryRq/></QBXMLMsgsRq></QBXML>";
        match QbxmlClient::new(app_id, app_name, company_file, mode) {
            Ok(mut client) => {
                match client.process_request(request_xml) {
                    Ok(response) => println!("QBXML response: {}", response),
                    Err(e) => eprintln!("QBXML request error: {:#}", e),
                }
                let _ = client.end_session();
                let _ = client.close_connection();
            }
            Err(e) => eprintln!("QBXML init error: {:#}", e),
        }
        unsafe { CoUninitialize(); }
        return Ok(());
    }


   
    // Create high-level sync service
    let sync_service = SyncService::new(config);
    
    // Perform the sync operation using clean, high-level API
    let result = sync_service.sync_account_to_sheets();
    
    // Uninitialize COM before exit
    unsafe { CoUninitialize(); }
    
    result?;
    
    Ok(())
}
