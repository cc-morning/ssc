use anyhow::Result;
use calamine::{open_workbook, Reader, Xlsx};
use clap::Parser;
use port_scanner::local_port_available;
use reqwest::{Client, Proxy, StatusCode};
use shadowsocks_service::{
    config::{Config, ConfigType, LocalConfig, ProtocolType},
    local,
    shadowsocks::{crypto::v1::CipherKind, ServerAddr, ServerConfig},
};
use std::{path::PathBuf, str::FromStr, time::Duration};
use tokio::time;

#[derive(Parser, Debug)]
#[clap(version = "0.0.1", author = "morning")]
struct Args {
    #[clap(required = true, value_name = "FILE", parse(from_os_str))]
    excel: PathBuf,
    #[clap(short = 's', long = "sheet", value_name = "SHEET")]
    excel_sheet: Option<String>,
    #[clap(short = 'H', long)]
    is_header: bool,
    #[clap(
        short = 'p',
        long = "port",
        value_name = "PORT",
        default_value = "1080"
    )]
    local_port: String,
    #[clap(short, long, value_name = "RETRY", default_value = "1")]
    retry: i32,
}

#[tokio::main]
async fn main() -> Result<()> {
    let args = Args::parse();

    if let Ok(mut excel) = open_workbook::<Xlsx<_>, PathBuf>(args.excel) {
        let sheet_name = if let Some(sheet_name) = args.excel_sheet {
            sheet_name
        } else {
            "Sheet1".to_string()
        };

        if let Some(Ok(range)) = excel.worksheet_range(&sheet_name) {
            let mut rows = range.rows().enumerate();

            if args.is_header {
                if let Some((_index, _row)) = rows.next() {}
            }

            let client = Client::builder()
                .proxy(Proxy::all(format!("http://127.0.0.1:{}", args.local_port))?)
                .build()?;

            let local_config = LocalConfig::new_with_addr(
                ServerAddr::SocketAddr(format!("127.0.0.1:{}", args.local_port).parse()?),
                ProtocolType::Http,
            );

            for (index, row) in rows {
                let addr = if let Some(value) = row.get(0) {
                    value.to_string()
                } else {
                    panic!("Line {} `addr` error!", index);
                };
                let port_str = if let Some(value) = row.get(1) {
                    value.to_string()
                } else {
                    panic!("Line {} `port` error!", index);
                };
                let port: u16 = port_str.parse()?;
                let method = if let Some(value) = row.get(2) {
                    value.to_string()
                } else {
                    panic!("Line {} `method` error!", index);
                };
                let password = if let Some(value) = row.get(3) {
                    value.to_string()
                } else {
                    panic!("Line {} `password` error!", index);
                };

                let mut config = Config::new(ConfigType::Local);
                config.local.push(local_config.clone());

                let server_config = ServerConfig::new(
                    ServerAddr::DomainName(addr.clone(), port),
                    password.clone(),
                    match CipherKind::from_str(&method) {
                        Ok(method) => method,
                        Err(_) => panic!("Line {} `method` error!", index),
                    },
                );
                config.server.push(server_config);

                let server = local::create(config.clone()).await?;

                let local_port = args.local_port.parse::<u16>()?;

                'wait_port: loop {
                    if local_port_available(local_port) {
                        break 'wait_port;
                    }

                    time::sleep(Duration::from_secs_f32(0.25)).await;
                }

                let task = tokio::spawn(async move {
                    if let Err(e) = server.run().await {
                        panic!("{:?}", e);
                    }
                });

                loop {
                    if !local_port_available(local_port) {
                        break;
                    }

                    time::sleep(Duration::from_secs_f32(0.25)).await;
                }

                let server_config = &[addr, format!("{}", port), method, password];

                let mut success = false;
                for _ in 0..args.retry {
                    if let Ok(response) = client.get("https://www.google.com").send().await {
                        if StatusCode::OK.eq(&response.status()) {
                            success = true;
                            break;
                        }
                    }
                }

                if !success {
                    eprintln!("Err : {:?}", server_config);
                }

                task.abort();
                if !task.await.unwrap_err().is_cancelled() {
                    panic!("Server stop failure!");
                }
            }
        }
    }

    Ok(())
}
