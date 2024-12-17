use std::io::{Read, Seek};

use calamine::{Data, Error, Range, Reader, Sheets};

pub fn operate_worksheets<RS: Read + Seek>(
    mut xl: Sheets<RS>,
    print_formula: bool,
) -> Result<(), Error> {
    for (sheet_name, range) in xl.worksheets() {
        let formulas = if print_formula {
            Some(xl.worksheet_formula(&sheet_name)?)
        } else {
            None
        };

        print_sheet_lines(&sheet_name, range, formulas);
    }

    print_vba(xl);
    Ok(())
}

fn print_sheet_lines(
    sheet_name: &str,
    range: Range<Data>,
    formulas: Option<Range<String>>,
) {
    for line in format_rows(range, formulas) {
        println!("[{}]{}", sheet_name, line);
    }
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
    formula.filter(|f| !f.is_empty())
           .map(|f| "=".to_string() + f)
           .or_else(|| value.map(|c| c.to_string()))
           .unwrap_or_default()
}

fn print_vba<RS: Read + Seek>(mut xl: Sheets<RS>) {
    if let Some(Ok(mut vba)) = xl.vba_project() {
        let vba = vba.to_mut();
        for module_name in vba.get_module_names() {
            if let Ok(module_content) = vba.get_module(module_name) {
                print_vba_module(module_name, &module_content);
            }
        }
    }
}

fn print_vba_module(module_name: &str, module_content: &str) {
    let lines = module_content
        .lines()
        .filter(|line| !line.starts_with("Attribute "))
        .collect::<Vec<_>>();
    if !lines.is_empty() {
        println!("[{}]", module_name);
        lines.iter().for_each(|line| println!("{}", line));
    }
}
