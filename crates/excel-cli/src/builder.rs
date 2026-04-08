use clap::{Arg, Command};
use excel_core::registry::{self, FlagDef, OperationDef, ServiceDef};

/// Build the full clap Command tree from the operation registry
/// This mirrors gws's dynamic CLI generation pattern
pub fn build_cli() -> Command {
    let mut app = Command::new("excel-cli")
        .version(env!("CARGO_PKG_VERSION"))
        .about("CLI tool for Excel file manipulation — built for humans and AI agents")
        .long_about(
            "excel-cli provides a comprehensive set of commands for working with Excel files.\n\
             Local operations work offline with .xlsx files.\n\
             Cloud operations (marked with AUTH) require Microsoft Graph API authentication.\n\n\
             Use --format to switch output: json (default), table, csv",
        )
        .arg(
            Arg::new("format")
                .long("format")
                .short('f')
                .global(true)
                .help("Output format: json, table, csv")
                .default_value("json"),
        )
        .arg(
            Arg::new("cloud")
                .long("cloud")
                .global(true)
                .action(clap::ArgAction::SetTrue)
                .help("Use Microsoft Graph API for cloud-only operations"),
        )
        .arg(
            Arg::new("item-id")
                .long("item-id")
                .global(true)
                .help("OneDrive item ID for cloud operations"),
        );

    // Add all services from registry
    for service in registry::all_services() {
        app = app.subcommand(build_service_command(&service));
    }

    // Add helper commands (+verb)
    app = app
        .subcommand(
            Command::new("+summarize")
                .about("Summarize workbook structure and statistics")
                .arg(Arg::new("file").required(true).help("Excel file path")),
        )
        .subcommand(
            Command::new("+diff")
                .about("Compare two workbooks")
                .arg(Arg::new("file1").required(true).help("First Excel file"))
                .arg(Arg::new("file2").required(true).help("Second Excel file")),
        )
        .subcommand(
            Command::new("+validate")
                .about("Validate workbook data against a JSON schema")
                .arg(Arg::new("schema").required(true).help("JSON schema file"))
                .arg(Arg::new("file").required(true).help("Excel file to validate")),
        )
        .subcommand(
            Command::new("+convert")
                .about("Convert file format")
                .arg(Arg::new("file").required(true).help("Source file"))
                .arg(
                    Arg::new("to")
                        .long("to")
                        .short('t')
                        .required(true)
                        .help("Target format: xlsx, csv, json"),
                )
                .arg(
                    Arg::new("output")
                        .long("output")
                        .short('o')
                        .help("Output file path"),
                ),
        )
        .subcommand(
            Command::new("+merge")
                .about("Merge two workbooks into one")
                .arg(Arg::new("file1").required(true).help("Base workbook"))
                .arg(Arg::new("file2").required(true).help("Workbook to merge in")),
        )
        .subcommand(
            Command::new("+template")
                .about("Create workbook from a template")
                .arg(Arg::new("template").required(true).help("Template name or path"))
                .arg(Arg::new("output").required(true).help("Output file path")),
        )
        .subcommand(
            Command::new("auth")
                .about("Authentication management for Graph API")
                .subcommand(
                    Command::new("login")
                        .about("Authenticate with Microsoft Graph API")
                        .arg(
                            Arg::new("client-id")
                                .long("client-id")
                                .required(true)
                                .help("Azure AD application (client) ID"),
                        )
                        .arg(
                            Arg::new("tenant-id")
                                .long("tenant-id")
                                .default_value("common")
                                .help("Azure AD tenant ID (default: common)"),
                        ),
                )
                .subcommand(Command::new("status").about("Show authentication status"))
                .subcommand(Command::new("logout").about("Remove saved credentials")),
        );

    app
}

fn build_service_command(service: &ServiceDef) -> Command {
    let mut cmd = Command::new(service.name).about(service.description);

    for op in &service.operations {
        cmd = cmd.subcommand(build_operation_command(op));
    }

    cmd
}

fn build_operation_command(op: &OperationDef) -> Command {
    let mut cmd = Command::new(op.verb).about(build_op_description(op));

    if let Some(long) = op.long_description {
        cmd = cmd.long_about(long);
    }

    // Add positional args
    for arg_def in &op.args {
        let mut arg = Arg::new(arg_def.name).help(arg_def.description);

        if arg_def.required {
            arg = arg.required(true);
        }

        if let Some(default) = arg_def.default {
            arg = arg.default_value(default);
        }

        cmd = cmd.arg(arg);
    }

    // Add flags
    for flag_def in &op.flags {
        cmd = cmd.arg(build_flag(flag_def));
    }

    cmd
}

fn build_flag(flag: &FlagDef) -> Arg {
    let mut arg = Arg::new(flag.name)
        .long(flag.name)
        .help(flag.description);

    if let Some(short) = flag.short {
        arg = arg.short(short);
    }

    if flag.takes_value {
        arg = arg.num_args(1);
        if let Some(default) = flag.default {
            arg = arg.default_value(default);
        }
    } else {
        arg = arg.action(clap::ArgAction::SetTrue);
    }

    arg
}

fn build_op_description(op: &OperationDef) -> String {
    if op.auth_required {
        format!("{} [AUTH]", op.description)
    } else {
        op.description.to_string()
    }
}
