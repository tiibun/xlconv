mod worksheets;

use std::{path::PathBuf, process::exit};

use calamine::open_workbook_auto;
use clap::Parser;
use worksheets::operate_worksheets;

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

    let xl = match open_workbook_auto(path) {
        Ok(s) => s,
        Err(e) => {
            eprintln!("Error: {}", e);
            exit(1);
        }
    };

    match operate_worksheets(xl, print_formula) {
        Ok(_) => (),
        Err(e) => {
            eprintln!("Error: {}", e);
            exit(1);
        }
    }
}
