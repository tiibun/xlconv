use std::{
    io::{Read, Seek},
    path::PathBuf,
    process::exit,
};

use calamine::{open_workbook_auto, Data, Range, Reader, Sheets};
use clap::Parser;

#[derive(Parser)]
#[command(version, about)]
struct Cli {
    /// Sets the input Excel file to parse
    input: PathBuf,

    /// Sets if print formula insted of value
    #[clap(short, long)]
    formula: bool,
}

fn main() {
    let cli = Cli::parse();
    let path = cli.input;
    let print_formula = cli.formula;

    let mut xl = match open_workbook_auto(path) {
        Ok(s) => s,
        Err(e) => {
            eprintln!("Error: {}", e);
            exit(1);
        }
    };

    for (sheet_name, range) in xl.worksheets() {
        let formulas = if print_formula {
            Some(match xl.worksheet_formula(&sheet_name) {
                Ok(f) => f,
                Err(e) => {
                    eprintln!("Error: {}", e);
                    exit(1);
                }
            })
        } else {
            None
        };

        for line in format_rows(range, formulas) {
            println!("[{}]{}", sheet_name, line);
        }
    }

    print_vba(xl)
}

fn format_rows(
    range: Range<Data>,
    formulas: Option<Range<String>>,
) -> impl Iterator<Item = String> {
    (0..=range.height()).map(move |row| {
        (0..=range.width())
            .map(|col| {
                let row = row.try_into().unwrap();
                let col = col.try_into().unwrap();
                get_value_or_formula(
                    range.get_value((row, col)),
                    formulas.as_ref().and_then(|f| f.get_value((row, col))),
                )
            })
            .collect::<Vec<_>>()
            .join("\t")
    })
}

fn get_value_or_formula(value: Option<&Data>, formula: Option<&String>) -> String {
    if let Some(f) = formula {
        if !f.is_empty() {
            return "=".to_string() + &f.to_string();
        }
    }
    if let Some(c) = value {
        c.to_string()
    } else {
        String::new()
    }
}

fn print_vba<RS: Read + Seek>(mut xl: Sheets<RS>) {
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
