mod builder;
mod config;
mod dispatch;
mod helpers;

use std::process;

#[tokio::main]
async fn main() {
    let matches = builder::build_cli().get_matches();

    match dispatch::dispatch(&matches).await {
        Ok(output) => {
            println!("{output}");
        }
        Err(e) => {
            let error_json = serde_json::json!({
                "error": true,
                "message": e.to_string(),
            });
            eprintln!("{}", serde_json::to_string_pretty(&error_json).unwrap());
            process::exit(1);
        }
    }
}
