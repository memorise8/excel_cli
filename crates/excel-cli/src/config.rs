use serde::{Deserialize, Serialize};
use std::path::PathBuf;

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct CliConfig {
    #[serde(default = "default_format")]
    pub default_format: String,
    pub graph_api: Option<GraphApiConfig>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct GraphApiConfig {
    pub client_id: String,
    pub tenant_id: String,
}

fn default_format() -> String {
    "json".to_string()
}

impl CliConfig {
    pub fn load() -> Self {
        let path = config_path();
        if path.exists() {
            let content = std::fs::read_to_string(&path).unwrap_or_default();
            toml::from_str(&content).unwrap_or_default()
        } else {
            Self::default()
        }
    }

    pub fn save(&self) -> Result<(), Box<dyn std::error::Error>> {
        let path = config_path();
        if let Some(parent) = path.parent() {
            std::fs::create_dir_all(parent)?;
        }
        let content = toml::to_string_pretty(self)?;
        std::fs::write(path, content)?;
        Ok(())
    }
}

fn config_path() -> PathBuf {
    dirs::config_dir()
        .unwrap_or_else(|| PathBuf::from("."))
        .join("excel-cli")
        .join("config.toml")
}
