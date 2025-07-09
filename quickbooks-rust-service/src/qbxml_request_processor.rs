// High-level wrapper for QBXML API, mirrors request_processor.rs but uses tickets
// This file is the QBXML equivalent of request_processor.rs

use crate::qbxml_safe;
use quickbooks_sheets_sync::request_processor;
use qbxml_safe::qbxml_request_processor::QbxmlRequestProcessor;

pub struct QbxmlClient {
    processor: QbxmlRequestProcessor,
    ticket: Option<String>,
}

impl QbxmlClient {
    pub fn new(app_id: &str, app_name: &str, company_file: &str, mode: i32) -> Result<Self, anyhow::Error> {
        let processor = QbxmlRequestProcessor::new()?;
        processor.open_connection2(app_id, app_name, 1)?; // 1 = local QB
        let ticket = processor.begin_session(company_file, mode)?;
        Ok(Self { processor, ticket: Some(ticket) })
    }

    pub fn process_request(&self, request_xml: &str) -> Result<String, anyhow::Error> {
        let ticket = self.ticket.as_ref().ok_or_else(|| anyhow::anyhow!("No session ticket"))?;
        self.processor.process_request(ticket, request_xml)
    }

    pub fn end_session(&mut self) -> Result<(), anyhow::Error> {
        if let Some(ticket) = &self.ticket {
            self.processor.end_session(ticket)?;
            self.ticket = None;
        }
        Ok(())
    }

    pub fn close_connection(&self) -> Result<(), anyhow::Error> {
        self.processor.close_connection()
    }
}
