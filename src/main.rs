use calamine::{open_workbook_auto, Reader};
use clap::Parser;

#[derive(Parser)]
#[command(version, about)]
struct Cli {
    /// Sets the input file to use
    input: String,
}

fn main() {
    let cli = Cli::parse();
    let path = cli.input;

    let mut xl = open_workbook_auto(path).expect("Failed to open file");

    for (sheet_name, range) in xl.worksheets() {
        for row in range.rows() {
            let line = row
                .iter()
                .map(|c| c.to_string())
                .collect::<Vec<_>>()
                .join("\t");
            println!("[{}]{}", sheet_name, line);
        }
    }

    if let Some(Ok(mut vba)) = xl.vba_project() {
        let vba = vba.to_mut();
        for module in vba.get_module_names() {
            if let Ok(s) = vba.get_module(module) {
                // initialize string vec
                let lines = s
                    .lines()
                    .filter(|l| !l.starts_with("Attribute "))
                    .collect::<Vec<_>>();
                if !lines.is_empty() {
                    println!("[{}]", module);
                    lines.iter().for_each(|l| println!("{}", l));
                }
            }
        }
    }
}
