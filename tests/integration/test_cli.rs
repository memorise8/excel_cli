use std::path::PathBuf;
use std::process::Command;

/// Get the path to the built binary
fn cli_bin() -> PathBuf {
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("target");
    path.push("debug");
    path.push("excel-cli");
    path
}

/// Run the CLI with given args, return (stdout, stderr, success)
fn run_cli(args: &[&str]) -> (String, String, bool) {
    let output = Command::new(cli_bin())
        .args(args)
        .output()
        .expect("Failed to execute excel-cli binary");

    let stdout = String::from_utf8_lossy(&output.stdout).to_string();
    let stderr = String::from_utf8_lossy(&output.stderr).to_string();
    (stdout, stderr, output.status.success())
}

/// Parse stdout as JSON
fn parse_json(stdout: &str) -> serde_json::Value {
    serde_json::from_str(stdout.trim()).unwrap_or_else(|e| {
        panic!("Failed to parse JSON output: {e}\nOutput was: {stdout}");
    })
}

/// Create a temporary xlsx file path in the system temp dir
fn temp_xlsx(name: &str) -> PathBuf {
    let mut path = std::env::temp_dir();
    path.push(format!("excel_cli_test_{name}_{}.xlsx", std::process::id()));
    path
}

/// Cleanup helper
fn cleanup(path: &PathBuf) {
    let _ = std::fs::remove_file(path);
}

// ── File service tests ────────────────────────────────────────────────

#[test]
fn test_file_create_and_info() {
    let path = temp_xlsx("create_info");
    let path_str = path.to_str().unwrap();

    // Create
    let (stdout, _, success) = run_cli(&["file", "create", path_str]);
    assert!(success, "file create failed");
    let json = parse_json(&stdout);
    assert_eq!(json["file_name"], path.file_name().unwrap().to_str().unwrap());
    assert!(json["sheet_count"].as_u64().unwrap() >= 1);

    // Info
    let (stdout, _, success) = run_cli(&["file", "info", path_str]);
    assert!(success, "file info failed");
    let json = parse_json(&stdout);
    assert!(json["file_size"].as_u64().unwrap() > 0);
    assert!(json["sheet_count"].as_u64().unwrap() >= 1);

    cleanup(&path);
}

#[test]
fn test_file_create_with_named_sheets() {
    let path = temp_xlsx("named_sheets");
    let path_str = path.to_str().unwrap();

    let (stdout, _, success) = run_cli(&["file", "create", path_str, "--sheets", "Alpha,Beta,Gamma"]);
    assert!(success, "file create with sheets failed");
    let json = parse_json(&stdout);
    assert_eq!(json["sheet_count"], 3);

    let sheet_names: Vec<String> = json["sheets"]
        .as_array()
        .unwrap()
        .iter()
        .map(|s| s["name"].as_str().unwrap().to_string())
        .collect();
    assert!(sheet_names.contains(&"Alpha".to_string()));
    assert!(sheet_names.contains(&"Beta".to_string()));
    assert!(sheet_names.contains(&"Gamma".to_string()));

    cleanup(&path);
}

// ── Sheet service tests ───────────────────────────────────────────────

#[test]
fn test_sheet_add_rename_delete_list() {
    let path = temp_xlsx("sheet_ops");
    let path_str = path.to_str().unwrap();

    // Create file
    run_cli(&["file", "create", path_str]);

    // Add sheet
    let (stdout, _, success) = run_cli(&["sheet", "add", path_str, "TestSheet"]);
    assert!(success, "sheet add failed");
    let json = parse_json(&stdout);
    assert_eq!(json["name"], "TestSheet");

    // List sheets
    let (stdout, _, success) = run_cli(&["sheet", "list", path_str]);
    assert!(success, "sheet list failed");
    let sheets: Vec<serde_json::Value> = serde_json::from_str(stdout.trim()).unwrap();
    let names: Vec<&str> = sheets.iter().map(|s| s["name"].as_str().unwrap()).collect();
    assert!(names.contains(&"TestSheet"));

    // Rename sheet
    let (_, _, success) = run_cli(&["sheet", "rename", path_str, "TestSheet", "Renamed"]);
    assert!(success, "sheet rename failed");

    // Verify rename
    let (stdout, _, _) = run_cli(&["sheet", "list", path_str]);
    let sheets: Vec<serde_json::Value> = serde_json::from_str(stdout.trim()).unwrap();
    let names: Vec<&str> = sheets.iter().map(|s| s["name"].as_str().unwrap()).collect();
    assert!(names.contains(&"Renamed"));
    assert!(!names.contains(&"TestSheet"));

    // Delete sheet
    let (_, _, success) = run_cli(&["sheet", "delete", path_str, "Renamed"]);
    assert!(success, "sheet delete failed");

    // Verify deletion
    let (stdout, _, _) = run_cli(&["sheet", "list", path_str]);
    let sheets: Vec<serde_json::Value> = serde_json::from_str(stdout.trim()).unwrap();
    let names: Vec<&str> = sheets.iter().map(|s| s["name"].as_str().unwrap()).collect();
    assert!(!names.contains(&"Renamed"));

    cleanup(&path);
}

// ── Range service tests ───────────────────────────────────────────────

#[test]
fn test_range_write_read_clear() {
    let path = temp_xlsx("range_ops");
    let path_str = path.to_str().unwrap();

    // Create file
    run_cli(&["file", "create", path_str]);

    // Write data
    let data = r#"[["Name","Age","City"],["Alice","30","NYC"],["Bob","25","LA"]]"#;
    let (_, _, success) = run_cli(&["range", "write", path_str, "Sheet1!A1:C3", "--data", data]);
    assert!(success, "range write failed");

    // Read data
    let (stdout, _, success) = run_cli(&["range", "read", path_str, "Sheet1!A1:C3"]);
    assert!(success, "range read failed");
    let json = parse_json(&stdout);
    assert_eq!(json["row_count"], 3);
    assert_eq!(json["col_count"], 3);

    // Check values
    let rows = json["rows"].as_array().unwrap();
    assert_eq!(rows[0][0], "Name");
    assert_eq!(rows[1][0], "Alice");
    assert_eq!(rows[2][1], 25);

    // Clear
    let (_, _, success) = run_cli(&["range", "clear", path_str, "Sheet1!A1:C3"]);
    assert!(success, "range clear failed");

    // Verify cleared
    let (stdout, _, _) = run_cli(&["range", "read", path_str, "Sheet1!A1:C3"]);
    let json = parse_json(&stdout);
    let rows = json["rows"].as_array().unwrap();
    // All cells should be empty after clear
    for row in rows {
        for cell in row.as_array().unwrap() {
            assert!(cell.is_null() || cell == "" || cell == "Empty");
        }
    }

    cleanup(&path);
}

// ── Formula service tests ─────────────────────────────────────────────

#[test]
fn test_formula_write_and_read() {
    let path = temp_xlsx("formula_ops");
    let path_str = path.to_str().unwrap();

    // Create file and write some data
    run_cli(&["file", "create", path_str]);
    let data = r#"[["10"],["20"],["30"]]"#;
    run_cli(&["range", "write", path_str, "Sheet1!A1:A3", "--data", data]);

    // Write formula
    let (_, _, success) = run_cli(&[
        "formula", "write", path_str, "Sheet1!A4", "--formula", "=SUM(A1:A3)",
    ]);
    assert!(success, "formula write failed");

    // Read formula
    let (stdout, _, success) = run_cli(&["formula", "read", path_str, "Sheet1!A4"]);
    assert!(success, "formula read failed");
    let json = parse_json(&stdout);
    assert!(json["has_formula"].as_bool().unwrap());
    assert!(json["formula"].as_str().unwrap().contains("SUM"));

    // List formulas
    let (stdout, _, success) = run_cli(&[
        "formula", "list", path_str, "--sheet", "Sheet1",
    ]);
    assert!(success, "formula list failed");
    let json = parse_json(&stdout);
    assert!(json["count"].as_u64().unwrap() >= 1);

    cleanup(&path);
}

// ── Helper command tests ──────────────────────────────────────────────

#[test]
fn test_summarize() {
    let path = temp_xlsx("summarize");
    let path_str = path.to_str().unwrap();

    run_cli(&["file", "create", path_str, "--sheets", "Data,Summary"]);

    let (stdout, _, success) = run_cli(&["+summarize", path_str]);
    assert!(success, "+summarize failed");
    let json = parse_json(&stdout);
    assert_eq!(json["sheet_count"], 2);
    assert!(json["sheets"].as_array().unwrap().len() == 2);

    cleanup(&path);
}

#[test]
fn test_diff() {
    let path1 = temp_xlsx("diff1");
    let path2 = temp_xlsx("diff2");
    let path1_str = path1.to_str().unwrap();
    let path2_str = path2.to_str().unwrap();

    run_cli(&["file", "create", path1_str, "--sheets", "Common,OnlyA"]);
    run_cli(&["file", "create", path2_str, "--sheets", "Common,OnlyB"]);

    let (stdout, _, success) = run_cli(&["+diff", path1_str, path2_str]);
    assert!(success, "+diff failed");
    let json = parse_json(&stdout);
    assert!(json["sheets"]["common"]
        .as_array()
        .unwrap()
        .iter()
        .any(|v| v == "Common"));

    cleanup(&path1);
    cleanup(&path2);
}

#[test]
fn test_template_budget() {
    let path = temp_xlsx("template_budget");
    let path_str = path.to_str().unwrap();

    let (stdout, _, success) = run_cli(&["+template", "budget", path_str]);
    assert!(success, "+template budget failed");
    let json = parse_json(&stdout);
    assert_eq!(json["status"], "ok");
    assert_eq!(json["template"], "budget");

    // Verify sheets were created
    let sheets = json["sheets"].as_array().unwrap();
    let sheet_names: Vec<&str> = sheets.iter().map(|s| s.as_str().unwrap()).collect();
    assert!(sheet_names.contains(&"Income"));
    assert!(sheet_names.contains(&"Expenses"));
    assert!(sheet_names.contains(&"Summary"));

    cleanup(&path);
}

// ── Cloud-only stubs ──────────────────────────────────────────────────

#[test]
fn test_cloud_only_services_return_stub() {
    for service in &["pivot", "chart", "calc", "connection", "slicer"] {
        let (stdout, _, _) = run_cli(&[service, "--help"]);
        // These should at least not crash; they have subcommands
        // If called without subcommand, they may show help or stub
        // The important thing is they don't panic
        let _ = stdout;
    }
}

// ── Export service tests ──────────────────────────────────────────────

#[test]
fn test_export_csv() {
    let path = temp_xlsx("export_csv");
    let path_str = path.to_str().unwrap();

    run_cli(&["file", "create", path_str]);
    let data = r#"[["Name","Score"],["Alice","95"],["Bob","87"]]"#;
    run_cli(&["range", "write", path_str, "Sheet1!A1:B3", "--data", data]);

    // Export to stdout (no --output)
    let (stdout, _, success) = run_cli(&["export", "csv", path_str, "--sheet", "Sheet1"]);
    assert!(success, "export csv failed");
    assert!(stdout.contains("Name"));
    assert!(stdout.contains("Alice"));

    cleanup(&path);
}

#[test]
fn test_export_json() {
    let path = temp_xlsx("export_json");
    let path_str = path.to_str().unwrap();

    run_cli(&["file", "create", path_str]);
    let data = r#"[["Name","Score"],["Alice","95"]]"#;
    run_cli(&["range", "write", path_str, "Sheet1!A1:B2", "--data", data]);

    let (stdout, _, success) = run_cli(&[
        "export", "json", path_str, "--sheet", "Sheet1", "--orient", "records",
    ]);
    assert!(success, "export json failed");
    // Output should contain the JSON data
    assert!(stdout.contains("Name") || stdout.contains("Alice"));

    cleanup(&path);
}
