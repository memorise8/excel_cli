use crate::models::{ExcelError, ExcelResult};
use serde::{Deserialize, Serialize};

/// OAuth2 token for Microsoft Graph API
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TokenInfo {
    pub access_token: String,
    pub refresh_token: Option<String>,
    pub expires_at: Option<chrono::DateTime<chrono::Utc>>,
    pub scopes: Vec<String>,
}

impl TokenInfo {
    pub fn is_expired(&self) -> bool {
        match self.expires_at {
            Some(expires) => chrono::Utc::now() >= expires,
            None => false,
        }
    }
}

/// Graph API authentication configuration
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AuthConfig {
    pub client_id: String,
    pub tenant_id: String,
    pub scopes: Vec<String>,
}

impl Default for AuthConfig {
    fn default() -> Self {
        Self {
            client_id: String::new(),
            tenant_id: "common".to_string(),
            scopes: vec![
                "Files.ReadWrite".to_string(),
                "Sites.ReadWrite.All".to_string(),
                "offline_access".to_string(),
            ],
        }
    }
}

/// Response from the device code endpoint
#[derive(Debug, Deserialize)]
struct DeviceCodeResponse {
    device_code: String,
    user_code: String,
    verification_uri: String,
    expires_in: u64,
    interval: u64,
    #[allow(dead_code)]
    message: Option<String>,
}

/// Response from the token poll endpoint
#[derive(Debug, Deserialize)]
struct TokenResponse {
    access_token: Option<String>,
    refresh_token: Option<String>,
    expires_in: Option<u64>,
    scope: Option<String>,
    error: Option<String>,
    #[allow(dead_code)]
    error_description: Option<String>,
}

/// Perform Microsoft Device Code OAuth2 flow.
/// Prints the user code and verification URL, then polls until token is granted.
pub async fn device_code_login(config: &AuthConfig) -> ExcelResult<TokenInfo> {
    let client = reqwest::Client::new();
    let tenant = &config.tenant_id;
    let scope = config.scopes.join(" ");

    // 1. Request device code
    let device_code_url = format!(
        "https://login.microsoftonline.com/{}/oauth2/v2.0/devicecode",
        tenant
    );

    let params = [
        ("client_id", config.client_id.as_str()),
        ("scope", scope.as_str()),
    ];

    let resp = client
        .post(&device_code_url)
        .form(&params)
        .send()
        .await
        .map_err(|e| ExcelError::CloudApiError(format!("Device code request failed: {}", e)))?;

    if !resp.status().is_success() {
        let body = resp.text().await.unwrap_or_default();
        return Err(ExcelError::CloudApiError(format!(
            "Device code endpoint error: {}",
            body
        )));
    }

    let dc: DeviceCodeResponse = resp
        .json()
        .await
        .map_err(|e| ExcelError::CloudApiError(format!("Failed to parse device code response: {}", e)))?;

    // 2. Show user instructions
    println!();
    println!("To sign in, visit: {}", dc.verification_uri);
    println!("Enter code:        {}", dc.user_code);
    println!();

    // 3. Poll for token
    let token_url = format!(
        "https://login.microsoftonline.com/{}/oauth2/v2.0/token",
        tenant
    );

    let poll_interval = std::time::Duration::from_secs(dc.interval.max(5));
    let deadline = std::time::Instant::now() + std::time::Duration::from_secs(dc.expires_in);

    loop {
        if std::time::Instant::now() >= deadline {
            return Err(ExcelError::AuthRequired(
                "Device code expired. Run 'auth login' again.".to_string(),
            ));
        }

        tokio::time::sleep(poll_interval).await;

        let poll_params = [
            ("client_id", config.client_id.as_str()),
            ("device_code", dc.device_code.as_str()),
            (
                "grant_type",
                "urn:ietf:params:oauth:grant-type:device_code",
            ),
        ];

        let poll_resp = client
            .post(&token_url)
            .form(&poll_params)
            .send()
            .await
            .map_err(|e| ExcelError::CloudApiError(format!("Token poll failed: {}", e)))?;

        let tr: TokenResponse = poll_resp
            .json()
            .await
            .map_err(|e| ExcelError::CloudApiError(format!("Failed to parse token response: {}", e)))?;

        match tr.error.as_deref() {
            None => {
                // Success
                let access_token = tr.access_token.ok_or_else(|| {
                    ExcelError::CloudApiError("No access_token in response".to_string())
                })?;

                let expires_at = tr.expires_in.map(|secs| {
                    chrono::Utc::now() + chrono::Duration::seconds(secs as i64)
                });

                let scopes = tr
                    .scope
                    .map(|s| s.split_whitespace().map(|s| s.to_string()).collect())
                    .unwrap_or_else(|| config.scopes.clone());

                return Ok(TokenInfo {
                    access_token,
                    refresh_token: tr.refresh_token,
                    expires_at,
                    scopes,
                });
            }
            Some("authorization_pending") => {
                // Still waiting — continue polling
                continue;
            }
            Some("slow_down") => {
                // Server asked us to back off
                tokio::time::sleep(poll_interval).await;
                continue;
            }
            Some(err) => {
                return Err(ExcelError::AuthRequired(format!(
                    "Authentication failed: {}",
                    err
                )));
            }
        }
    }
}

/// Refresh an expired access token using the refresh token.
pub async fn refresh_access_token(config: &AuthConfig, refresh_token: &str) -> ExcelResult<TokenInfo> {
    let client = reqwest::Client::new();
    let token_url = format!(
        "https://login.microsoftonline.com/{}/oauth2/v2.0/token",
        config.tenant_id
    );

    let params = [
        ("client_id", config.client_id.as_str()),
        ("grant_type", "refresh_token"),
        ("refresh_token", refresh_token),
        ("scope", &config.scopes.join(" ")),
    ];

    let resp = client
        .post(&token_url)
        .form(&params)
        .send()
        .await
        .map_err(|e| ExcelError::CloudApiError(format!("Token refresh failed: {}", e)))?;

    if !resp.status().is_success() {
        let body = resp.text().await.unwrap_or_default();
        return Err(ExcelError::AuthRequired(format!(
            "Token refresh failed. Please run 'auth login' again. Error: {}",
            body
        )));
    }

    let tr: TokenResponse = resp
        .json()
        .await
        .map_err(|e| ExcelError::CloudApiError(format!("Failed to parse refresh response: {}", e)))?;

    let access_token = tr.access_token.ok_or_else(|| {
        ExcelError::CloudApiError("No access_token in refresh response".to_string())
    })?;

    let expires_at = tr.expires_in.map(|secs| {
        chrono::Utc::now() + chrono::Duration::seconds(secs as i64)
    });

    let scopes = tr
        .scope
        .map(|s| s.split_whitespace().map(|s| s.to_string()).collect())
        .unwrap_or_else(|| config.scopes.clone());

    Ok(TokenInfo {
        access_token,
        refresh_token: tr.refresh_token.or_else(|| Some(refresh_token.to_string())),
        expires_at,
        scopes,
    })
}

/// Load saved token from config directory
pub fn load_token() -> ExcelResult<Option<TokenInfo>> {
    let config_dir = config_dir();
    let token_path = config_dir.join("token.json");

    if !token_path.exists() {
        return Ok(None);
    }

    let content = std::fs::read_to_string(&token_path)?;
    let token: TokenInfo = serde_json::from_str(&content)?;

    if token.is_expired() {
        return Ok(None);
    }

    Ok(Some(token))
}

/// Save token to config directory
pub fn save_token(token: &TokenInfo) -> ExcelResult<()> {
    let config_dir = config_dir();
    std::fs::create_dir_all(&config_dir)?;

    let token_path = config_dir.join("token.json");
    let content = serde_json::to_string_pretty(token)?;
    std::fs::write(token_path, content)?;

    Ok(())
}

/// Remove saved token
pub fn remove_token() -> ExcelResult<()> {
    let token_path = config_dir().join("token.json");
    if token_path.exists() {
        std::fs::remove_file(token_path)?;
    }
    Ok(())
}

fn config_dir() -> std::path::PathBuf {
    dirs::config_dir()
        .unwrap_or_else(|| std::path::PathBuf::from("."))
        .join("excel-cli")
}
