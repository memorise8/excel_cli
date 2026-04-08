use clap::ArgMatches;
use excel_core::models::*;
use excel_core::output::{format_output, OutputFormat};
use excel_core::services::ExcelService;
use excel_core::services::local::{conditional, export, format, formula, named_range, table};
use excel_core::services::graph::GraphService;
use excel_core::services::graph::auth;
use excel_core::LocalService;
use std::path::Path;


pub async fn dispatch(matches: &ArgMatches) -> Result<String, Box<dyn std::error::Error>> {
    let format = OutputFormat::from_str(
        matches
            .get_one::<String>("format")
            .map(|s| s.as_str())
            .unwrap_or("json"),
    );

    let service = LocalService::new();

    match matches.subcommand() {
        // File service
        Some(("file", sub)) => dispatch_file(sub, &service, format).await,
        // Sheet service
        Some(("sheet", sub)) => dispatch_sheet(sub, &service, format).await,
        // Range service
        Some(("range", sub)) => dispatch_range(sub, &service, format).await,
        // Auth service
        Some(("auth", sub)) => dispatch_auth(sub).await,
        // Helper commands
        Some(("+summarize", sub)) => dispatch_summarize(sub, &service, format).await,
        Some(("+diff", sub)) => dispatch_diff(sub, format).await,
        // Phase 3 services (local implementations)
        Some(("formula", sub)) => dispatch_formula(sub, format).await,
        Some(("format", sub)) => dispatch_format(sub, format).await,
        Some(("conditional", sub)) => dispatch_conditional(sub, format).await,
        Some(("table", sub)) => dispatch_table(sub, format).await,
        Some(("named-range", sub)) => dispatch_named_range(sub, format).await,
        Some(("export", sub)) => dispatch_export(sub, format).await,
        // Cloud services
        Some(("pivot", sub)) => dispatch_pivot(sub, matches, format).await,
        Some(("chart", sub)) => dispatch_chart(sub, matches, format).await,
        Some(("calc", sub)) => dispatch_calc(sub, matches, format).await,
        Some(("connection", _)) | Some(("slicer", _)) => {
            Ok(r#"{"status": "not_yet_implemented", "message": "Connection and slicer services require Graph API and will be implemented in a future phase"}"#.to_string())
        }
        Some(("+validate", sub)) => dispatch_validate(sub, format).await,
        Some(("+convert", sub)) => dispatch_convert(sub, format).await,
        Some(("+merge", sub)) => dispatch_merge(sub, format).await,
        Some(("+template", sub)) => dispatch_template(sub, format).await,
        _ => Ok(r#"{"error": "Unknown command. Use --help for available commands."}"#.to_string()),
    }
}

async fn dispatch_file(
    matches: &ArgMatches,
    service: &LocalService,
    format: OutputFormat,
) -> Result<String, Box<dyn std::error::Error>> {
    match matches.subcommand() {
        Some(("create", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let sheets = sub
                .get_one::<String>("sheets")
                .map(|s| s.split(',').map(|n| n.trim().to_string()).collect());
            let info = service.file_create(Path::new(file), sheets)?;
            Ok(format_output(&info, format))
        }
        Some(("info", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let info = service.file_info(Path::new(file))?;
            Ok(format_output(&info, format))
        }
        Some(("save", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let output = sub.get_one::<String>("output").unwrap();
            service.file_save(Path::new(file), Path::new(output))?;
            Ok(format_output(
                &serde_json::json!({"status": "ok", "saved_to": output}),
                format,
            ))
        }
        Some(("convert", _sub)) => Ok(format_output(
            &serde_json::json!({"status": "not_yet_implemented"}),
            format,
        )),
        Some(("upload", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let token_info = auth::load_token()?
                .ok_or("Not authenticated. Run 'excel-cli auth login' first.")?;
            let graph = GraphService::new(Some(token_info.access_token));
            let result = graph.file_upload(file).await?;
            let item_id = result.get("id").and_then(|v| v.as_str()).unwrap_or("");
            let name = result.get("name").and_then(|v| v.as_str()).unwrap_or("");
            Ok(format_output(
                &serde_json::json!({
                    "status": "ok",
                    "action": "uploaded",
                    "item_id": item_id,
                    "name": name,
                    "message": format!("Use --item-id {} for subsequent cloud operations", item_id),
                }),
                format,
            ))
        }
        Some(("download", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let token_info = auth::load_token()?
                .ok_or("Not authenticated. Run 'excel-cli auth login' first.")?;
            let item_id = sub.get_one::<String>("item-id")
                .ok_or("--item-id is required for download")?;
            let graph = GraphService::new(Some(token_info.access_token));
            graph.file_download(item_id, file).await?;
            Ok(format_output(
                &serde_json::json!({"status": "ok", "action": "downloaded", "file": file}),
                format,
            ))
        }
        _ => Ok(r#"{"error": "Unknown file subcommand"}"#.to_string()),
    }
}

async fn dispatch_sheet(
    matches: &ArgMatches,
    service: &LocalService,
    format: OutputFormat,
) -> Result<String, Box<dyn std::error::Error>> {
    match matches.subcommand() {
        Some(("list", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let sheets = service.sheet_list(Path::new(file))?;
            Ok(format_output(&sheets, format))
        }
        Some(("add", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let name = sub.get_one::<String>("sheet").unwrap();
            let position = sub
                .get_one::<String>("position")
                .and_then(|p| p.parse().ok());
            let info = service.sheet_add(Path::new(file), name, position)?;
            Ok(format_output(&info, format))
        }
        Some(("rename", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let name = sub.get_one::<String>("sheet").unwrap();
            let new_name = sub.get_one::<String>("new-name").unwrap();
            service.sheet_rename(Path::new(file), name, new_name)?;
            Ok(format_output(
                &serde_json::json!({"status": "ok", "renamed": {"from": name, "to": new_name}}),
                format,
            ))
        }
        Some(("delete", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let name = sub.get_one::<String>("sheet").unwrap();
            service.sheet_delete(Path::new(file), name)?;
            Ok(format_output(
                &serde_json::json!({"status": "ok", "deleted": name}),
                format,
            ))
        }
        Some(("copy", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let name = sub.get_one::<String>("sheet").unwrap();
            let new_name = sub.get_one::<String>("new-name").map(|s| s.as_str());
            let info = service.sheet_copy(Path::new(file), name, new_name)?;
            Ok(format_output(&info, format))
        }
        _ => Ok(format_output(
            &serde_json::json!({"status": "not_yet_implemented"}),
            format,
        )),
    }
}

async fn dispatch_range(
    matches: &ArgMatches,
    service: &LocalService,
    format: OutputFormat,
) -> Result<String, Box<dyn std::error::Error>> {
    // Cloud routing
    if matches.get_flag("cloud") {
        return dispatch_range_cloud(matches, format).await;
    }

    match matches.subcommand() {
        Some(("read", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let range = sub.get_one::<String>("range").unwrap();
            let data = service.range_read(Path::new(file), range)?;
            Ok(format_output(&data, format))
        }
        Some(("write", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let range = sub.get_one::<String>("range").unwrap();

            let data: Vec<Vec<CellValue>> = if let Some(json_data) = sub.get_one::<String>("data")
            {
                let raw: Vec<Vec<serde_json::Value>> = serde_json::from_str(json_data)?;
                raw.into_iter()
                    .map(|row| {
                        row.into_iter()
                            .map(|v| match v {
                                serde_json::Value::Null => CellValue::Empty,
                                serde_json::Value::Bool(b) => CellValue::Bool(b),
                                serde_json::Value::Number(n) => {
                                    if let Some(i) = n.as_i64() {
                                        CellValue::Int(i)
                                    } else {
                                        CellValue::Float(n.as_f64().unwrap_or(0.0))
                                    }
                                }
                                serde_json::Value::String(s) => CellValue::String(s),
                                _ => CellValue::String(v.to_string()),
                            })
                            .collect()
                    })
                    .collect()
            } else if let Some(value) = sub.get_one::<String>("value") {
                vec![vec![CellValue::String(value.clone())]]
            } else {
                return Err("Either --data or --value is required".into());
            };

            service.range_write(Path::new(file), range, data)?;
            Ok(format_output(
                &serde_json::json!({"status": "ok", "range": range}),
                format,
            ))
        }
        Some(("clear", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let range = sub.get_one::<String>("range").unwrap();
            let values_only = sub.get_flag("values-only");
            service.range_clear(Path::new(file), range, values_only)?;
            Ok(format_output(
                &serde_json::json!({"status": "ok", "cleared": range}),
                format,
            ))
        }
        _ => Ok(format_output(
            &serde_json::json!({"status": "not_yet_implemented"}),
            format,
        )),
    }
}

async fn dispatch_auth(matches: &ArgMatches) -> Result<String, Box<dyn std::error::Error>> {
    match matches.subcommand() {
        Some(("login", sub)) => {
            let client_id = sub.get_one::<String>("client-id")
                .ok_or("--client-id is required. Register an Azure AD app first.")?;
            let tenant_id = sub.get_one::<String>("tenant-id")
                .map(|s| s.to_string())
                .unwrap_or_else(|| "common".to_string());

            let config = excel_core::services::graph::auth::AuthConfig {
                client_id: client_id.to_string(),
                tenant_id,
                ..Default::default()
            };

            let token = excel_core::services::graph::auth::device_code_login(&config).await?;
            excel_core::services::graph::auth::save_token(&token)?;

            Ok(serde_json::to_string_pretty(&serde_json::json!({
                "status": "ok",
                "message": "Authentication successful",
                "expires_at": token.expires_at,
                "scopes": token.scopes,
            }))?)
        }
        Some(("status", _)) => {
            let token = excel_core::services::graph::auth::load_token()?;
            match token {
                Some(t) => Ok(serde_json::to_string_pretty(&serde_json::json!({
                    "authenticated": true,
                    "expires_at": t.expires_at,
                    "scopes": t.scopes,
                }))?),
                None => Ok(serde_json::to_string_pretty(&serde_json::json!({
                    "authenticated": false,
                    "message": "Run 'excel-cli auth login' to authenticate"
                }))?),
            }
        }
        Some(("logout", _)) => {
            excel_core::services::graph::auth::remove_token()?;
            Ok(serde_json::to_string_pretty(&serde_json::json!({
                "status": "ok",
                "message": "Logged out successfully"
            }))?)
        }
        _ => Ok(r#"{"error": "Unknown auth subcommand"}"#.to_string()),
    }
}

async fn dispatch_summarize(
    matches: &ArgMatches,
    service: &LocalService,
    format: OutputFormat,
) -> Result<String, Box<dyn std::error::Error>> {
    let file = matches.get_one::<String>("file").unwrap();
    let info = service.file_info(Path::new(file))?;

    let summary = serde_json::json!({
        "file": info.file_name,
        "size_bytes": info.file_size,
        "sheet_count": info.sheet_count,
        "sheets": info.sheets.iter().map(|s| {
            serde_json::json!({
                "name": s.name,
                "rows": s.row_count,
                "cols": s.col_count,
                "visible": s.visible,
            })
        }).collect::<Vec<_>>(),
    });

    Ok(format_output(&summary, format))
}

async fn dispatch_diff(
    matches: &ArgMatches,
    format: OutputFormat,
) -> Result<String, Box<dyn std::error::Error>> {
    let file1 = matches.get_one::<String>("file1").unwrap();
    let file2 = matches.get_one::<String>("file2").unwrap();

    let service = LocalService::new();
    let info1 = service.file_info(Path::new(file1))?;
    let info2 = service.file_info(Path::new(file2))?;

    let sheets1: std::collections::HashSet<String> =
        info1.sheets.iter().map(|s| s.name.clone()).collect();
    let sheets2: std::collections::HashSet<String> =
        info2.sheets.iter().map(|s| s.name.clone()).collect();

    let only_in_1: Vec<&String> = sheets1.difference(&sheets2).collect();
    let only_in_2: Vec<&String> = sheets2.difference(&sheets1).collect();
    let common: Vec<&String> = sheets1.intersection(&sheets2).collect();

    let diff = serde_json::json!({
        "file1": info1.file_name,
        "file2": info2.file_name,
        "sheets": {
            "only_in_file1": only_in_1,
            "only_in_file2": only_in_2,
            "common": common,
        },
        "size_diff": info2.file_size as i64 - info1.file_size as i64,
    });

    Ok(format_output(&diff, format))
}

async fn dispatch_validate(
    matches: &ArgMatches,
    format: OutputFormat,
) -> Result<String, Box<dyn std::error::Error>> {
    let schema_path = matches.get_one::<String>("schema").unwrap();
    let file_path = matches.get_one::<String>("file").unwrap();

    let schema_content = std::fs::read_to_string(schema_path)?;
    let schema: serde_json::Value = serde_json::from_str(&schema_content)?;

    let service = LocalService::new();
    let info = service.file_info(Path::new(file_path))?;

    let mut errors: Vec<serde_json::Value> = Vec::new();

    // Validate required sheets
    if let Some(required_sheets) = schema.get("required_sheets").and_then(|v| v.as_array()) {
        for rs in required_sheets {
            if let Some(name) = rs.as_str() {
                if !info.sheets.iter().any(|s| s.name == name) {
                    errors.push(serde_json::json!({
                        "type": "missing_sheet",
                        "message": format!("Required sheet '{name}' not found"),
                    }));
                }
            }
        }
    }

    // Validate required columns per sheet
    if let Some(sheet_schemas) = schema.get("sheets").and_then(|v| v.as_object()) {
        for (sheet_name, sheet_schema) in sheet_schemas {
            if let Some(required_cols) = sheet_schema.get("required_columns").and_then(|v| v.as_array()) {
                let header_range = format!("{sheet_name}!A1:Z1");
                match service.range_read(Path::new(file_path), &header_range) {
                    Ok(data) => {
                        if let Some(first_row) = data.rows.first() {
                            let headers: Vec<String> = first_row.iter().map(|v| match v {
                                CellValue::String(s) => s.clone(),
                                _ => String::new(),
                            }).collect();
                            for col in required_cols {
                                if let Some(col_name) = col.as_str() {
                                    if !headers.iter().any(|h| h == col_name) {
                                        errors.push(serde_json::json!({
                                            "type": "missing_column",
                                            "sheet": sheet_name,
                                            "message": format!("Required column '{col_name}' not found in sheet '{sheet_name}'"),
                                        }));
                                    }
                                }
                            }
                        }
                    }
                    Err(_) => {
                        errors.push(serde_json::json!({
                            "type": "sheet_read_error",
                            "sheet": sheet_name,
                            "message": format!("Could not read sheet '{sheet_name}'"),
                        }));
                    }
                }
            }
        }
    }

    let result = serde_json::json!({
        "valid": errors.is_empty(),
        "file": file_path,
        "schema": schema_path,
        "error_count": errors.len(),
        "errors": errors,
    });

    Ok(format_output(&result, format))
}

async fn dispatch_convert(
    matches: &ArgMatches,
    format: OutputFormat,
) -> Result<String, Box<dyn std::error::Error>> {
    let file_path = matches.get_one::<String>("file").unwrap();
    let target_format = matches.get_one::<String>("to").unwrap();
    let output_path = matches.get_one::<String>("output").map(|s| s.clone());

    let service = LocalService::new();
    let info = service.file_info(Path::new(file_path))?;

    let default_output = match target_format.as_str() {
        "csv" => file_path.replace(".xlsx", ".csv"),
        "json" => file_path.replace(".xlsx", ".json"),
        "xlsx" => file_path.replace(".csv", ".xlsx").replace(".json", ".xlsx"),
        _ => return Err(format!("Unsupported format: {target_format}").into()),
    };
    let out = output_path.unwrap_or(default_output);

    match target_format.as_str() {
        "csv" => {
            let sheet_name = info.sheets.first().map(|s| s.name.clone()).unwrap_or_default();
            let max_row = info.sheets.first().and_then(|s| s.row_count).unwrap_or(0);
            let max_col = info.sheets.first().and_then(|s| s.col_count).unwrap_or(0);
            if max_row > 0 && max_col > 0 {
                let end_col = excel_core::models::range::col_index_to_letter(max_col as u32);
                let range_str = format!("{sheet_name}!A1:{end_col}{max_row}");
                let data = service.range_read(Path::new(file_path), &range_str)?;
                let mut csv_output = String::new();
                for row in &data.rows {
                    let line: Vec<String> = row.iter().map(|v| match v {
                        CellValue::Empty => String::new(),
                        CellValue::Bool(b) => b.to_string(),
                        CellValue::Int(i) => i.to_string(),
                        CellValue::Float(f) => f.to_string(),
                        CellValue::String(s) => {
                            if s.contains(',') || s.contains('"') || s.contains('\n') {
                                format!("\"{}\"", s.replace('"', "\"\""))
                            } else {
                                s.clone()
                            }
                        }
                        CellValue::Formula(fv) => fv.cached_value.as_ref()
                            .map(|v| format!("{v:?}")).unwrap_or_else(|| fv.formula.clone()),
                        CellValue::Error(e) => e.clone(),
                    }).collect();
                    csv_output.push_str(&line.join(","));
                    csv_output.push('\n');
                }
                std::fs::write(&out, csv_output)?;
            } else {
                std::fs::write(&out, "")?;
            }
        }
        "json" => {
            let sheet_name = info.sheets.first().map(|s| s.name.clone()).unwrap_or_default();
            let max_row = info.sheets.first().and_then(|s| s.row_count).unwrap_or(0);
            let max_col = info.sheets.first().and_then(|s| s.col_count).unwrap_or(0);
            if max_row > 0 && max_col > 0 {
                let end_col = excel_core::models::range::col_index_to_letter(max_col as u32);
                let range_str = format!("{sheet_name}!A1:{end_col}{max_row}");
                let data = service.range_read(Path::new(file_path), &range_str)?;
                let json_str = serde_json::to_string_pretty(&data)?;
                std::fs::write(&out, json_str)?;
            } else {
                std::fs::write(&out, "[]")?;
            }
        }
        _ => return Err(format!("Unsupported target format: {target_format}").into()),
    }

    Ok(format_output(
        &serde_json::json!({"status": "ok", "converted_to": out, "format": target_format}),
        format,
    ))
}

async fn dispatch_merge(
    matches: &ArgMatches,
    format: OutputFormat,
) -> Result<String, Box<dyn std::error::Error>> {
    let file1 = matches.get_one::<String>("file1").unwrap();
    let file2 = matches.get_one::<String>("file2").unwrap();

    let service = LocalService::new();
    let info1 = service.file_info(Path::new(file1))?;
    let info2 = service.file_info(Path::new(file2))?;

    let existing_names: std::collections::HashSet<String> = info1
        .sheets
        .iter()
        .map(|s| s.name.clone())
        .collect();

    let mut added_sheets: Vec<String> = Vec::new();

    for sheet in &info2.sheets {
        let mut name = sheet.name.clone();
        if existing_names.contains(&name) {
            name = format!("{name}_merged");
        }
        service.sheet_add(Path::new(file1), &name, None)?;
        added_sheets.push(name);
    }

    Ok(format_output(
        &serde_json::json!({
            "status": "ok",
            "merged_into": file1,
            "added_sheets": added_sheets,
        }),
        format,
    ))
}

async fn dispatch_template(
    matches: &ArgMatches,
    format: OutputFormat,
) -> Result<String, Box<dyn std::error::Error>> {
    let template = matches.get_one::<String>("template").unwrap();
    let output = matches.get_one::<String>("output").unwrap();

    let service = LocalService::new();

    match template.as_str() {
        "blank" => {
            service.file_create(Path::new(output), None)?;
        }
        "budget" => {
            service.file_create(
                Path::new(output),
                Some(vec!["Income".to_string(), "Expenses".to_string(), "Summary".to_string()]),
            )?;
            // Add headers
            service.range_write(
                Path::new(output),
                "Income!A1:C1",
                vec![vec![
                    CellValue::String("Date".to_string()),
                    CellValue::String("Description".to_string()),
                    CellValue::String("Amount".to_string()),
                ]],
            )?;
            service.range_write(
                Path::new(output),
                "Expenses!A1:D1",
                vec![vec![
                    CellValue::String("Date".to_string()),
                    CellValue::String("Category".to_string()),
                    CellValue::String("Description".to_string()),
                    CellValue::String("Amount".to_string()),
                ]],
            )?;
        }
        "tracker" => {
            service.file_create(
                Path::new(output),
                Some(vec!["Tasks".to_string(), "Done".to_string()]),
            )?;
            service.range_write(
                Path::new(output),
                "Tasks!A1:D1",
                vec![vec![
                    CellValue::String("ID".to_string()),
                    CellValue::String("Task".to_string()),
                    CellValue::String("Status".to_string()),
                    CellValue::String("Due Date".to_string()),
                ]],
            )?;
        }
        "sales" => {
            service.file_create(
                Path::new(output),
                Some(vec!["Sales".to_string(), "Products".to_string(), "Customers".to_string()]),
            )?;
            service.range_write(
                Path::new(output),
                "Sales!A1:E1",
                vec![vec![
                    CellValue::String("Date".to_string()),
                    CellValue::String("Product".to_string()),
                    CellValue::String("Customer".to_string()),
                    CellValue::String("Quantity".to_string()),
                    CellValue::String("Amount".to_string()),
                ]],
            )?;
        }
        // If it's a file path, copy it as template
        other if std::path::Path::new(other).exists() => {
            service.file_save(Path::new(other), Path::new(output))?;
        }
        _ => {
            return Err(format!(
                "Unknown template: '{template}'. Available: blank, budget, tracker, sales, or path to .xlsx file"
            ).into());
        }
    }

    let info = service.file_info(Path::new(output))?;
    Ok(format_output(
        &serde_json::json!({
            "status": "ok",
            "template": template,
            "created": output,
            "sheets": info.sheets.iter().map(|s| &s.name).collect::<Vec<_>>(),
        }),
        format,
    ))
}

// ═══════════════════════════════════════════════════════════════════════
// Cloud dispatch functions (Microsoft Graph API)
// ═══════════════════════════════════════════════════════════════════════

async fn dispatch_range_cloud(
    matches: &ArgMatches,
    format: OutputFormat,
) -> Result<String, Box<dyn std::error::Error>> {
    let token_info = auth::load_token()?
        .ok_or("Not authenticated. Run 'excel-cli auth login' first.")?;
    let item_id = matches
        .get_one::<String>("item-id")
        .ok_or("--item-id is required for cloud operations")?;
    let graph = GraphService::new(Some(token_info.access_token));

    match matches.subcommand() {
        Some(("read", sub)) => {
            let range = sub.get_one::<String>("range").unwrap();
            let with_format = sub.get_flag("with-format");

            let (sheet, range_part) = if range.contains('!') {
                let parts: Vec<&str> = range.splitn(2, '!').collect();
                (parts[0].to_string(), parts[1].to_string())
            } else {
                ("Sheet1".to_string(), range.to_string())
            };

            let data = graph.range_read(item_id, &sheet, &range_part).await?;

            if with_format {
                let font = graph.range_read_font(item_id, &sheet, &range_part).await?;
                let fill = graph.range_read_fill(item_id, &sheet, &range_part).await?;
                let borders = graph.range_read_borders(item_id, &sheet, &range_part).await?;
                let result = serde_json::json!({
                    "values": data.get("values"),
                    "formulas": data.get("formulas"),
                    "numberFormat": data.get("numberFormat"),
                    "format": {
                        "font": font,
                        "fill": fill,
                        "borders": borders,
                    }
                });
                Ok(format_output(&result, format))
            } else {
                Ok(format_output(&data, format))
            }
        }
        Some(("write", sub)) => {
            let range = sub.get_one::<String>("range").unwrap();
            let data_str = sub.get_one::<String>("data")
                .ok_or("--data is required")?;

            let (sheet, range_part) = if range.contains('!') {
                let parts: Vec<&str> = range.splitn(2, '!').collect();
                (parts[0].to_string(), parts[1].to_string())
            } else {
                ("Sheet1".to_string(), range.to_string())
            };

            let values: serde_json::Value = serde_json::from_str(data_str)?;
            let body = serde_json::json!({"values": values});
            let _result = graph.range_write(item_id, &sheet, &range_part, body).await?;
            Ok(format_output(
                &serde_json::json!({"status": "ok", "range": range, "cloud": true}),
                format,
            ))
        }
        _ => Ok(format_output(
            &serde_json::json!({"error": "Unknown or unsupported cloud range subcommand"}),
            format,
        )),
    }
}

async fn dispatch_calc(
    matches: &ArgMatches,
    root_matches: &ArgMatches,
    format: OutputFormat,
) -> Result<String, Box<dyn std::error::Error>> {
    let token_info = auth::load_token()?
        .ok_or("Not authenticated. Run 'excel-cli auth login' first.")?;
    let item_id = root_matches
        .get_one::<String>("item-id")
        .ok_or("--item-id is required for calc operations")?;
    let graph = GraphService::new(Some(token_info.access_token));

    match matches.subcommand() {
        Some(("now", _)) => {
            graph.calc_now(item_id).await?;
            Ok(format_output(
                &serde_json::json!({"status": "ok", "action": "recalculated", "item_id": item_id}),
                format,
            ))
        }
        Some(("mode", _)) => Ok(format_output(
            &serde_json::json!({"status": "not_yet_implemented"}),
            format,
        )),
        _ => Ok(r#"{"error": "Unknown calc subcommand. Use: now"}"#.to_string()),
    }
}

async fn dispatch_chart(
    matches: &ArgMatches,
    root_matches: &ArgMatches,
    format: OutputFormat,
) -> Result<String, Box<dyn std::error::Error>> {
    let token_info = auth::load_token()?
        .ok_or("Not authenticated. Run 'excel-cli auth login' first.")?;
    let item_id = root_matches
        .get_one::<String>("item-id")
        .ok_or("--item-id is required for chart operations")?;
    let graph = GraphService::new(Some(token_info.access_token));

    match matches.subcommand() {
        Some(("list", sub)) => {
            let sheet = sub.get_one::<String>("sheet")
                .ok_or("--sheet is required")?;
            let result = graph.chart_list(item_id, sheet).await?;
            Ok(format_output(&result, format))
        }
        Some(("create", sub)) => {
            let sheet = sub.get_one::<String>("sheet")
                .ok_or("--sheet is required")?;
            let chart_type = sub.get_one::<String>("type")
                .ok_or("--type is required")?;
            let source = sub.get_one::<String>("source")
                .ok_or("--source is required (data range)")?;
            let series_by = sub.get_one::<String>("series-by")
                .map(|s| s.as_str())
                .unwrap_or("Auto");
            let result = graph.chart_create(item_id, sheet, chart_type, source, series_by).await?;
            Ok(format_output(&result, format))
        }
        Some(("delete", sub)) => {
            let sheet = sub.get_one::<String>("sheet")
                .ok_or("--sheet is required")?;
            let name = sub.get_one::<String>("name")
                .ok_or("--name is required")?;
            graph.chart_delete(item_id, sheet, name).await?;
            Ok(format_output(
                &serde_json::json!({"status": "ok", "deleted": name}),
                format,
            ))
        }
        _ => Ok(format_output(
            &serde_json::json!({"status": "not_yet_implemented", "message": "Use list, create, or delete"}),
            format,
        )),
    }
}

async fn dispatch_pivot(
    matches: &ArgMatches,
    root_matches: &ArgMatches,
    format: OutputFormat,
) -> Result<String, Box<dyn std::error::Error>> {
    let token_info = auth::load_token()?
        .ok_or("Not authenticated. Run 'excel-cli auth login' first.")?;
    let item_id = root_matches
        .get_one::<String>("item-id")
        .ok_or("--item-id is required for pivot operations")?;
    let graph = GraphService::new(Some(token_info.access_token));

    match matches.subcommand() {
        Some(("list", sub)) => {
            let sheet = sub.get_one::<String>("sheet")
                .ok_or("--sheet is required")?;
            let result = graph.pivot_list(item_id, sheet).await?;
            Ok(format_output(&result, format))
        }
        Some(("refresh", sub)) => {
            let sheet = sub.get_one::<String>("sheet")
                .ok_or("--sheet is required")?;
            let name = sub.get_one::<String>("name")
                .ok_or("--name is required")?;
            graph.pivot_refresh(item_id, sheet, name).await?;
            Ok(format_output(
                &serde_json::json!({"status": "ok", "refreshed": name}),
                format,
            ))
        }
        _ => Ok(format_output(
            &serde_json::json!({"status": "not_yet_implemented", "message": "Use list or refresh"}),
            format,
        )),
    }
}

// ── Phase 3: Formula service ──────────────────────────────────────────

async fn dispatch_formula(
    matches: &ArgMatches,
    format: OutputFormat,
) -> Result<String, Box<dyn std::error::Error>> {
    match matches.subcommand() {
        Some(("read", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let range = sub.get_one::<String>("range").unwrap();
            let result = formula::read(Path::new(file), range)?;
            Ok(format_output(&result, format))
        }
        Some(("write", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let range = sub.get_one::<String>("range").unwrap();
            let formula_text = sub
                .get_one::<String>("formula")
                .ok_or("--formula is required")?;
            formula::write(Path::new(file), range, formula_text)?;
            Ok(format_output(
                &serde_json::json!({"status": "ok", "cell": range, "formula": formula_text}),
                format,
            ))
        }
        Some(("list", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let sheet = sub
                .get_one::<String>("sheet")
                .map(|s| s.as_str())
                .unwrap_or("Sheet1");
            let result = formula::list(Path::new(file), sheet)?;
            Ok(format_output(&result, format))
        }
        Some(("audit", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let range = sub.get_one::<String>("range").unwrap();
            let direction = sub
                .get_one::<String>("direction")
                .map(|s| s.as_str())
                .unwrap_or("precedents");
            let result = formula::audit(Path::new(file), range, direction)?;
            Ok(format_output(&result, format))
        }
        Some(("evaluate", _)) => Ok(format_output(
            &serde_json::json!({"status": "auth_required", "message": "Formula evaluation requires --cloud flag"}),
            format,
        )),
        _ => Ok(r#"{"error": "Unknown formula subcommand"}"#.to_string()),
    }
}

// ── Phase 3: Format service ───────────────────────────────────────────

async fn dispatch_format(
    matches: &ArgMatches,
    format: OutputFormat,
) -> Result<String, Box<dyn std::error::Error>> {
    match matches.subcommand() {
        Some(("font", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let range = sub.get_one::<String>("range").unwrap();
            let name = sub.get_one::<String>("name").map(|s| s.as_str());
            let size = sub
                .get_one::<String>("size")
                .and_then(|s| s.parse::<f64>().ok());
            let color = sub.get_one::<String>("color").map(|s| s.as_str());
            let bold = sub.get_flag("bold");
            let italic = sub.get_flag("italic");
            let underline = sub.get_flag("underline");
            format::font(Path::new(file), range, name, size, color, bold, italic, underline)?;
            Ok(format_output(
                &serde_json::json!({"status": "ok", "range": range, "action": "font"}),
                format,
            ))
        }
        Some(("fill", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let range = sub.get_one::<String>("range").unwrap();
            let color = sub
                .get_one::<String>("color")
                .ok_or("--color is required")?;
            let pattern = sub
                .get_one::<String>("pattern")
                .map(|s| s.as_str())
                .unwrap_or("solid");
            format::fill(Path::new(file), range, color, pattern)?;
            Ok(format_output(
                &serde_json::json!({"status": "ok", "range": range, "action": "fill", "color": color}),
                format,
            ))
        }
        Some(("border", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let range = sub.get_one::<String>("range").unwrap();
            let style = sub
                .get_one::<String>("style")
                .map(|s| s.as_str())
                .unwrap_or("thin");
            let color = sub.get_one::<String>("color").map(|s| s.as_str());
            let sides = sub
                .get_one::<String>("sides")
                .map(|s| s.as_str())
                .unwrap_or("all");
            format::border(Path::new(file), range, style, color, sides)?;
            Ok(format_output(
                &serde_json::json!({"status": "ok", "range": range, "action": "border"}),
                format,
            ))
        }
        Some(("align", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let range = sub.get_one::<String>("range").unwrap();
            let horizontal = sub.get_one::<String>("horizontal").map(|s| s.as_str());
            let vertical = sub.get_one::<String>("vertical").map(|s| s.as_str());
            let wrap = sub.get_flag("wrap");
            format::align(Path::new(file), range, horizontal, vertical, wrap)?;
            Ok(format_output(
                &serde_json::json!({"status": "ok", "range": range, "action": "align"}),
                format,
            ))
        }
        Some(("number", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let range = sub.get_one::<String>("range").unwrap();
            let format_code = sub.get_one::<String>("format").map(|s| s.as_str());
            let preset = sub.get_one::<String>("preset").map(|s| s.as_str());
            format::number_format(Path::new(file), range, format_code, preset)?;
            Ok(format_output(
                &serde_json::json!({"status": "ok", "range": range, "action": "number_format"}),
                format,
            ))
        }
        Some(("width", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let sheet = sub
                .get_one::<String>("sheet")
                .ok_or("--sheet is required")?;
            let col = sub.get_one::<String>("col").ok_or("--col is required")?;
            let width: f64 = sub
                .get_one::<String>("width")
                .ok_or("--width is required")?
                .parse()?;
            format::column_width(Path::new(file), sheet, col, width)?;
            Ok(format_output(
                &serde_json::json!({"status": "ok", "sheet": sheet, "col": col, "width": width}),
                format,
            ))
        }
        Some(("height", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let sheet = sub
                .get_one::<String>("sheet")
                .ok_or("--sheet is required")?;
            let row: u32 = sub
                .get_one::<String>("row")
                .ok_or("--row is required")?
                .parse()?;
            let height: f64 = sub
                .get_one::<String>("height")
                .ok_or("--height is required")?
                .parse()?;
            format::row_height(Path::new(file), sheet, row, height)?;
            Ok(format_output(
                &serde_json::json!({"status": "ok", "sheet": sheet, "row": row, "height": height}),
                format,
            ))
        }
        Some(("autofit", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let sheet = sub
                .get_one::<String>("sheet")
                .ok_or("--sheet is required")?;
            let cols = sub.get_one::<String>("cols").map(|s| s.as_str());
            format::autofit(Path::new(file), sheet, cols)?;
            Ok(format_output(
                &serde_json::json!({"status": "ok", "sheet": sheet, "action": "autofit"}),
                format,
            ))
        }
        Some(("style", _sub)) => Ok(format_output(
            &serde_json::json!({"status": "not_yet_implemented", "message": "Compound style will be supported in a future update"}),
            format,
        )),
        _ => Ok(r#"{"error": "Unknown format subcommand"}"#.to_string()),
    }
}

// ── Phase 3: Conditional formatting service ───────────────────────────

async fn dispatch_conditional(
    matches: &ArgMatches,
    format: OutputFormat,
) -> Result<String, Box<dyn std::error::Error>> {
    match matches.subcommand() {
        Some(("add", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let range = sub.get_one::<String>("range").unwrap();
            let rule_type = sub
                .get_one::<String>("type")
                .ok_or("--type is required")?;
            let operator = sub.get_one::<String>("operator").map(|s| s.as_str());
            let value = sub.get_one::<String>("value").map(|s| s.as_str());
            let format_json = sub.get_one::<String>("format-json").map(|s| s.as_str());
            let result =
                conditional::add(Path::new(file), range, rule_type, operator, value, format_json)?;
            Ok(format_output(&result, format))
        }
        Some(("list", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let sheet = sub
                .get_one::<String>("sheet")
                .ok_or("--sheet is required")?;
            let result = conditional::list(Path::new(file), sheet)?;
            Ok(format_output(&result, format))
        }
        Some(("delete", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let sheet = sub
                .get_one::<String>("sheet")
                .ok_or("--sheet is required")?;
            let index: usize = sub
                .get_one::<String>("index")
                .ok_or("--index is required")?
                .parse()?;
            conditional::delete(Path::new(file), sheet, index)?;
            Ok(format_output(
                &serde_json::json!({"status": "ok", "deleted_index": index}),
                format,
            ))
        }
        Some(("clear", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let range = sub.get_one::<String>("range").unwrap();
            conditional::clear(Path::new(file), range)?;
            Ok(format_output(
                &serde_json::json!({"status": "ok", "cleared": range}),
                format,
            ))
        }
        _ => Ok(r#"{"error": "Unknown conditional subcommand"}"#.to_string()),
    }
}

// ── Phase 3: Table service ────────────────────────────────────────────

async fn dispatch_table(
    matches: &ArgMatches,
    format: OutputFormat,
) -> Result<String, Box<dyn std::error::Error>> {
    match matches.subcommand() {
        Some(("list", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let tables = table::list(Path::new(file))?;
            Ok(format_output(&tables, format))
        }
        Some(("create", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let range = sub.get_one::<String>("range").unwrap();
            let name = sub.get_one::<String>("name").ok_or("--name is required")?;
            let style = sub.get_one::<String>("style").map(|s| s.as_str());
            let has_headers = sub.get_flag("has-headers");
            let info = table::create(Path::new(file), range, name, style, has_headers)?;
            Ok(format_output(&info, format))
        }
        Some(("read", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let name = sub.get_one::<String>("name").ok_or("--name is required")?;
            let data = table::read(Path::new(file), name)?;
            Ok(format_output(&data, format))
        }
        Some(("append", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let name = sub.get_one::<String>("name").ok_or("--name is required")?;
            let data_str = sub.get_one::<String>("data").ok_or("--data is required")?;
            let rows: Vec<Vec<serde_json::Value>> = serde_json::from_str(data_str)?;
            table::append(Path::new(file), name, rows)?;
            Ok(format_output(
                &serde_json::json!({"status": "ok", "table": name, "action": "append"}),
                format,
            ))
        }
        Some(("delete", _sub)) | Some(("resize", _sub)) | Some(("rename", _sub))
        | Some(("sort", _sub)) | Some(("filter", _sub)) | Some(("style", _sub))
        | Some(("total-row", _sub)) | Some(("column-add", _sub))
        | Some(("column-delete", _sub)) | Some(("to-range", _sub)) => Ok(format_output(
            &serde_json::json!({"status": "not_yet_implemented", "message": "This table operation is planned for a future update"}),
            format,
        )),
        _ => Ok(r#"{"error": "Unknown table subcommand"}"#.to_string()),
    }
}

// ── Phase 3: Named range service ──────────────────────────────────────

async fn dispatch_named_range(
    matches: &ArgMatches,
    format: OutputFormat,
) -> Result<String, Box<dyn std::error::Error>> {
    match matches.subcommand() {
        Some(("list", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let result = named_range::list(Path::new(file))?;
            Ok(format_output(&result, format))
        }
        Some(("create", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let name = sub.get_one::<String>("name").ok_or("--name is required")?;
            let refers_to = sub
                .get_one::<String>("refers-to")
                .ok_or("--refers-to is required")?;
            let scope = sub
                .get_one::<String>("scope")
                .map(|s| s.as_str())
                .unwrap_or("workbook");
            let result = named_range::create(Path::new(file), name, refers_to, scope)?;
            Ok(format_output(&result, format))
        }
        Some(("delete", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let name = sub.get_one::<String>("name").ok_or("--name is required")?;
            named_range::delete(Path::new(file), name)?;
            Ok(format_output(
                &serde_json::json!({"status": "ok", "deleted": name}),
                format,
            ))
        }
        Some(("update", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let name = sub.get_one::<String>("name").ok_or("--name is required")?;
            let refers_to = sub
                .get_one::<String>("refers-to")
                .ok_or("--refers-to is required")?;
            let result = named_range::update(Path::new(file), name, refers_to)?;
            Ok(format_output(&result, format))
        }
        Some(("resolve", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let name = sub.get_one::<String>("name").ok_or("--name is required")?;
            let result = named_range::resolve(Path::new(file), name)?;
            Ok(format_output(&result, format))
        }
        Some(("read", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let name = sub.get_one::<String>("name").ok_or("--name is required")?;
            let result = named_range::read_values(Path::new(file), name)?;
            Ok(format_output(&result, format))
        }
        _ => Ok(r#"{"error": "Unknown named-range subcommand"}"#.to_string()),
    }
}

// ── Phase 3: Export service ───────────────────────────────────────────

async fn dispatch_export(
    matches: &ArgMatches,
    format: OutputFormat,
) -> Result<String, Box<dyn std::error::Error>> {
    match matches.subcommand() {
        Some(("csv", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let sheet = sub
                .get_one::<String>("sheet")
                .ok_or("--sheet is required")?;
            let output_path = sub.get_one::<String>("output").map(|s| s.as_str());
            let delimiter = sub
                .get_one::<String>("delimiter")
                .and_then(|s| s.chars().next())
                .unwrap_or(',');
            let content = export::to_csv(Path::new(file), sheet, delimiter)?;
            let result = export::save_or_return(content, output_path)?;
            Ok(format_output(&serde_json::json!(serde_json::from_str::<serde_json::Value>(&result).unwrap_or(serde_json::Value::String(result))), format))
        }
        Some(("json", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let sheet = sub
                .get_one::<String>("sheet")
                .ok_or("--sheet is required")?;
            let output_path = sub.get_one::<String>("output").map(|s| s.as_str());
            let orient = sub
                .get_one::<String>("orient")
                .map(|s| s.as_str())
                .unwrap_or("records");
            let content = export::to_json(Path::new(file), sheet, orient)?;
            let result = export::save_or_return(content, output_path)?;
            Ok(format_output(&serde_json::json!(serde_json::from_str::<serde_json::Value>(&result).unwrap_or(serde_json::Value::String(result))), format))
        }
        Some(("html", sub)) => {
            let file = sub.get_one::<String>("file").unwrap();
            let sheet = sub
                .get_one::<String>("sheet")
                .ok_or("--sheet is required")?;
            let output_path = sub.get_one::<String>("output").map(|s| s.as_str());
            let content = export::to_html(Path::new(file), sheet)?;
            let result = export::save_or_return(content, output_path)?;
            Ok(format_output(&serde_json::json!(serde_json::from_str::<serde_json::Value>(&result).unwrap_or(serde_json::Value::String(result))), format))
        }
        Some(("pdf", _)) | Some(("screenshot", _)) => Ok(format_output(
            &serde_json::json!({"status": "auth_required", "message": "PDF/screenshot export requires --cloud flag"}),
            format,
        )),
        _ => Ok(r#"{"error": "Unknown export subcommand"}"#.to_string()),
    }
}
