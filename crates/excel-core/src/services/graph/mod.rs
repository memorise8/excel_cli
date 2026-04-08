pub mod auth;

use crate::models::{ExcelError, ExcelResult};

const GRAPH_BASE_URL: &str = "https://graph.microsoft.com/v1.0";

/// Microsoft Graph API backend for Excel operations
/// Provides capabilities not available locally: formula evaluation, charts, pivots
///
/// Requires OAuth2 authentication with Microsoft Graph API.
/// Use `excel-cli auth login` to authenticate.
pub struct GraphService {
    token: Option<String>,
    client: reqwest::Client,
}

impl GraphService {
    pub fn new(token: Option<String>) -> Self {
        Self {
            token,
            client: reqwest::Client::new(),
        }
    }

    pub fn is_authenticated(&self) -> bool {
        self.token.is_some()
    }

    /// Make an authenticated request to the Microsoft Graph API.
    /// `method` is "GET", "POST", "PATCH", "DELETE", etc.
    /// `path` is the path after the base URL, e.g. "/me/drive/root".
    /// `body` is an optional JSON payload.
    pub async fn graph_request(
        &self,
        method: &str,
        path: &str,
        body: Option<serde_json::Value>,
    ) -> ExcelResult<serde_json::Value> {
        let token = self.token.as_deref().ok_or_else(|| {
            ExcelError::AuthRequired(
                "Not authenticated. Run 'excel-cli auth login' first.".to_string(),
            )
        })?;

        let url = format!("{}{}", GRAPH_BASE_URL, path);

        let mut req = self
            .client
            .request(
                method.parse().map_err(|_| {
                    ExcelError::CloudApiError(format!("Invalid HTTP method: {}", method))
                })?,
                &url,
            )
            .bearer_auth(token)
            .header("Content-Type", "application/json");

        if let Some(json_body) = body {
            req = req.json(&json_body);
        }

        let resp = req
            .send()
            .await
            .map_err(|e| ExcelError::CloudApiError(format!("Graph API request failed: {}", e)))?;

        let status = resp.status();
        let json: serde_json::Value = resp.json().await.map_err(|e| {
            ExcelError::CloudApiError(format!("Failed to parse Graph API response: {}", e))
        })?;

        if !status.is_success() {
            let message = json
                .pointer("/error/message")
                .and_then(|v| v.as_str())
                .unwrap_or("Unknown Graph API error")
                .to_string();
            return Err(ExcelError::CloudApiError(format!(
                "Graph API error {}: {}",
                status, message
            )));
        }

        Ok(json)
    }

    /// Make an authenticated request that returns raw bytes (for file downloads, PDF export).
    pub async fn graph_request_bytes(
        &self,
        method: &str,
        path: &str,
    ) -> ExcelResult<Vec<u8>> {
        let token = self.token.as_deref().ok_or_else(|| {
            ExcelError::AuthRequired("Not authenticated. Run 'excel-cli auth login' first.".to_string())
        })?;

        let url = format!("{}{}", GRAPH_BASE_URL, path);
        let resp = self
            .client
            .request(
                method.parse().map_err(|_| {
                    ExcelError::CloudApiError(format!("Invalid HTTP method: {}", method))
                })?,
                &url,
            )
            .bearer_auth(token)
            .send()
            .await
            .map_err(|e| ExcelError::CloudApiError(format!("Graph API request failed: {}", e)))?;

        let status = resp.status();
        if !status.is_success() {
            let body = resp.text().await.unwrap_or_default();
            return Err(ExcelError::CloudApiError(format!("Graph API error {}: {}", status, body)));
        }

        resp.bytes()
            .await
            .map(|b| b.to_vec())
            .map_err(|e| ExcelError::CloudApiError(format!("Failed to read response bytes: {}", e)))
    }

    /// Upload raw bytes with a specific content type.
    pub async fn graph_upload_bytes(
        &self,
        path: &str,
        bytes: &[u8],
        content_type: &str,
    ) -> ExcelResult<serde_json::Value> {
        let token = self.token.as_deref().ok_or_else(|| {
            ExcelError::AuthRequired("Not authenticated. Run 'excel-cli auth login' first.".to_string())
        })?;

        let url = format!("{}{}", GRAPH_BASE_URL, path);
        let resp = self
            .client
            .put(&url)
            .bearer_auth(token)
            .header("Content-Type", content_type)
            .body(bytes.to_vec())
            .send()
            .await
            .map_err(|e| ExcelError::CloudApiError(format!("Upload failed: {}", e)))?;

        let status = resp.status();
        let json: serde_json::Value = resp.json().await.map_err(|e| {
            ExcelError::CloudApiError(format!("Failed to parse upload response: {}", e))
        })?;

        if !status.is_success() {
            let message = json.pointer("/error/message")
                .and_then(|v| v.as_str())
                .unwrap_or("Upload failed")
                .to_string();
            return Err(ExcelError::CloudApiError(format!("Graph API error {}: {}", status, message)));
        }

        Ok(json)
    }

    // ═══════════════════════════════════════════
    // File operations
    // ═══════════════════════════════════════════

    /// Upload a local file to OneDrive root. Returns the drive item metadata including `id`.
    pub async fn file_upload(&self, local_path: &str) -> ExcelResult<serde_json::Value> {
        let path = std::path::Path::new(local_path);
        let filename = path.file_name()
            .and_then(|n| n.to_str())
            .ok_or_else(|| ExcelError::CloudApiError("Invalid file path".to_string()))?;

        let bytes = std::fs::read(local_path)
            .map_err(|e| ExcelError::Io(e))?;

        let api_path = format!("/me/drive/root:/{}:/content", filename);
        self.graph_upload_bytes(&api_path, &bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet").await
    }

    /// Download a file from OneDrive by item ID to a local path.
    pub async fn file_download(&self, item_id: &str, local_path: &str) -> ExcelResult<()> {
        let api_path = format!("/me/drive/items/{}/content", item_id);
        let bytes = self.graph_request_bytes("GET", &api_path).await?;
        std::fs::write(local_path, &bytes)
            .map_err(|e| ExcelError::Io(e))?;
        Ok(())
    }

    // ═══════════════════════════════════════════
    // Session management
    // ═══════════════════════════════════════════

    /// Create a persistent workbook session for batching operations.
    pub async fn create_session(&self, item_id: &str) -> ExcelResult<String> {
        let path = format!("/me/drive/items/{}/workbook/createSession", item_id);
        let body = serde_json::json!({"persistChanges": true});
        let result = self.graph_request("POST", &path, Some(body)).await?;
        result.get("id")
            .and_then(|v| v.as_str())
            .map(|s| s.to_string())
            .ok_or_else(|| ExcelError::CloudApiError("No session ID in response".to_string()))
    }

    /// Close a workbook session.
    pub async fn close_session(&self, item_id: &str, session_id: &str) -> ExcelResult<()> {
        let token = self.token.as_deref().ok_or_else(|| {
            ExcelError::AuthRequired("Not authenticated".to_string())
        })?;

        let url = format!("{}/me/drive/items/{}/workbook/closeSession", GRAPH_BASE_URL, item_id);
        let resp = self.client
            .post(&url)
            .bearer_auth(token)
            .header("workbook-session-id", session_id)
            .header("Content-Length", "0")
            .send()
            .await
            .map_err(|e| ExcelError::CloudApiError(format!("Close session failed: {}", e)))?;

        if !resp.status().is_success() {
            let body = resp.text().await.unwrap_or_default();
            return Err(ExcelError::CloudApiError(format!("Close session error: {}", body)));
        }
        Ok(())
    }

    // ═══════════════════════════════════════════
    // Worksheet operations
    // ═══════════════════════════════════════════

    /// List all worksheets in a workbook.
    pub async fn worksheet_list(&self, item_id: &str) -> ExcelResult<serde_json::Value> {
        let path = format!("/me/drive/items/{}/workbook/worksheets", item_id);
        self.graph_request("GET", &path, None).await
    }

    // ═══════════════════════════════════════════
    // Range operations
    // ═══════════════════════════════════════════

    /// Read range values, formulas, and number formats.
    pub async fn range_read(
        &self,
        item_id: &str,
        sheet: &str,
        range: &str,
    ) -> ExcelResult<serde_json::Value> {
        let path = format!(
            "/me/drive/items/{}/workbook/worksheets/{}/range(address='{}')",
            item_id, sheet, range
        );
        self.graph_request("GET", &path, None).await
    }

    /// Read range format (font, fill, borders, alignment, number format).
    pub async fn range_read_format(
        &self,
        item_id: &str,
        sheet: &str,
        range: &str,
    ) -> ExcelResult<serde_json::Value> {
        let path = format!(
            "/me/drive/items/{}/workbook/worksheets/{}/range(address='{}')/format",
            item_id, sheet, range
        );
        self.graph_request("GET", &path, None).await
    }

    /// Read range font properties.
    pub async fn range_read_font(
        &self,
        item_id: &str,
        sheet: &str,
        range: &str,
    ) -> ExcelResult<serde_json::Value> {
        let path = format!(
            "/me/drive/items/{}/workbook/worksheets/{}/range(address='{}')/format/font",
            item_id, sheet, range
        );
        self.graph_request("GET", &path, None).await
    }

    /// Read range fill properties.
    pub async fn range_read_fill(
        &self,
        item_id: &str,
        sheet: &str,
        range: &str,
    ) -> ExcelResult<serde_json::Value> {
        let path = format!(
            "/me/drive/items/{}/workbook/worksheets/{}/range(address='{}')/format/fill",
            item_id, sheet, range
        );
        self.graph_request("GET", &path, None).await
    }

    /// Read range border properties.
    pub async fn range_read_borders(
        &self,
        item_id: &str,
        sheet: &str,
        range: &str,
    ) -> ExcelResult<serde_json::Value> {
        let path = format!(
            "/me/drive/items/{}/workbook/worksheets/{}/range(address='{}')/format/borders",
            item_id, sheet, range
        );
        self.graph_request("GET", &path, None).await
    }

    /// Write values/formulas/numberFormat to a range.
    pub async fn range_write(
        &self,
        item_id: &str,
        sheet: &str,
        range: &str,
        body: serde_json::Value,
    ) -> ExcelResult<serde_json::Value> {
        let path = format!(
            "/me/drive/items/{}/workbook/worksheets/{}/range(address='{}')",
            item_id, sheet, range
        );
        self.graph_request("PATCH", &path, Some(body)).await
    }

    /// Write font format to a range.
    pub async fn range_write_font(
        &self,
        item_id: &str,
        sheet: &str,
        range: &str,
        font: serde_json::Value,
    ) -> ExcelResult<serde_json::Value> {
        let path = format!(
            "/me/drive/items/{}/workbook/worksheets/{}/range(address='{}')/format/font",
            item_id, sheet, range
        );
        self.graph_request("PATCH", &path, Some(font)).await
    }

    /// Write fill format to a range.
    pub async fn range_write_fill(
        &self,
        item_id: &str,
        sheet: &str,
        range: &str,
        fill: serde_json::Value,
    ) -> ExcelResult<serde_json::Value> {
        let path = format!(
            "/me/drive/items/{}/workbook/worksheets/{}/range(address='{}')/format/fill",
            item_id, sheet, range
        );
        self.graph_request("PATCH", &path, Some(fill)).await
    }

    // ═══════════════════════════════════════════
    // Calculation
    // ═══════════════════════════════════════════

    /// Trigger full workbook recalculation.
    pub async fn calc_now(&self, item_id: &str) -> ExcelResult<serde_json::Value> {
        let path = format!("/me/drive/items/{}/workbook/application/calculate", item_id);
        let body = serde_json::json!({"calculationType": "Full"});
        self.graph_request("POST", &path, Some(body)).await
    }

    // ═══════════════════════════════════════════
    // Export
    // ═══════════════════════════════════════════

    /// Export workbook as PDF.
    pub async fn export_pdf(&self, item_id: &str, output_path: &str) -> ExcelResult<()> {
        let path = format!("/me/drive/items/{}/content?format=pdf", item_id);
        let bytes = self.graph_request_bytes("GET", &path).await?;
        std::fs::write(output_path, &bytes)
            .map_err(|e| ExcelError::Io(e))?;
        Ok(())
    }

    // ═══════════════════════════════════════════
    // Chart operations
    // ═══════════════════════════════════════════

    /// List all charts in a worksheet.
    pub async fn chart_list(&self, item_id: &str, sheet: &str) -> ExcelResult<serde_json::Value> {
        let path = format!("/me/drive/items/{}/workbook/worksheets/{}/charts", item_id, sheet);
        self.graph_request("GET", &path, None).await
    }

    /// Create a chart.
    pub async fn chart_create(
        &self,
        item_id: &str,
        sheet: &str,
        chart_type: &str,
        source_data: &str,
        series_by: &str,
    ) -> ExcelResult<serde_json::Value> {
        let path = format!("/me/drive/items/{}/workbook/worksheets/{}/charts/add", item_id, sheet);
        let body = serde_json::json!({
            "type": chart_type,
            "sourceData": source_data,
            "seriesBy": series_by,
        });
        self.graph_request("POST", &path, Some(body)).await
    }

    /// Delete a chart by name.
    pub async fn chart_delete(&self, item_id: &str, sheet: &str, name: &str) -> ExcelResult<()> {
        let path = format!("/me/drive/items/{}/workbook/worksheets/{}/charts/{}", item_id, sheet, name);
        self.graph_request("DELETE", &path, None).await?;
        Ok(())
    }

    // ═══════════════════════════════════════════
    // Pivot Table operations
    // ═══════════════════════════════════════════

    /// List all pivot tables in a worksheet.
    pub async fn pivot_list(&self, item_id: &str, sheet: &str) -> ExcelResult<serde_json::Value> {
        let path = format!("/me/drive/items/{}/workbook/worksheets/{}/pivotTables", item_id, sheet);
        self.graph_request("GET", &path, None).await
    }

    /// Refresh a pivot table.
    pub async fn pivot_refresh(&self, item_id: &str, sheet: &str, name: &str) -> ExcelResult<()> {
        let path = format!(
            "/me/drive/items/{}/workbook/worksheets/{}/pivotTables/{}/refresh",
            item_id, sheet, name
        );
        self.graph_request("POST", &path, None).await?;
        Ok(())
    }

    /// Refresh all pivot tables in a workbook.
    pub async fn pivot_refresh_all(&self, item_id: &str) -> ExcelResult<()> {
        let path = format!(
            "/me/drive/items/{}/workbook/refreshAllPivotTables",
            item_id
        );
        self.graph_request("POST", &path, None).await?;
        Ok(())
    }
}
