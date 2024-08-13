use std::io::{Read, Seek};

use calamine::{Data, Error, Range, Reader, Sheets};

pub fn operate_worksheets<RS: Read + Seek>(
    mut xl: Sheets<RS>,
    print_formula: bool,
) -> Result<(), Error> {
    for (sheet_name, range) in xl.worksheets() {
        let formulas = if print_formula {
            Some(match xl.worksheet_formula(&sheet_name) {
                Ok(f) => f,
                Err(e) => {
                    return Err(e);
                }
            })
        } else {
            None
        };

        for line in format_rows(range, formulas) {
            println!("[{}]{}", sheet_name, line);
        }
    }

    print_vba(xl);
    Ok(())
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
