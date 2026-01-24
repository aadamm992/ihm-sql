mod cli;
mod lot;

use crate::cli::Cli;
use crate::lot::Lot;
use calamine::{DataType, Reader, Xlsx, open_workbook};
use clap::Parser;
use std::fs;
use std::io::Write;

fn main() {
    // parse cli arguments
    let cli = Cli::parse();

    // load work book from the given parameter and return a Vec<Lot>
    let lots = load_workbook(
        &cli.file,
        cli.sheets.as_slice(),
        cli.fauf_row as u32,
        cli.fauf_col as u32,
        cli.charge_row as u32,
        cli.charge_col as u32,
        cli.modul_range_row,
        cli.modul_range_column,
    );

    // get current username
    let user_name = std::env::var("USERNAME").unwrap();

    // if excess flag is true, create transaction queries to the file
    if cli.exec {
        write_file(
            &lots,
            &format!(
                "C:\\Users\\{}\\Documents\\SQL Server Management Studio\\Traceability.sql",
                user_name
            ),
            create_sql,
        );
    }

    // print terminal to the fauf and charge
    if cli.list {
        println!("{}", create_list(&lots));
    }
}

/// Read Excel sheets and return a Vec<Lot>
///
/// # Arguments
///
/// * `file_path`: file full path
/// * `sheets`: necessary sheets names
///
/// returns: `Vec<Lot, Global>`
#[allow(clippy::too_many_arguments)]
fn load_workbook(
    file_path: &str,
    sheets: &[String],
    fauf_row: u32,
    fauf_col: u32,
    charge_row: u32,
    charge_col: u32,
    modul_range_row: usize,
    modul_range_column: usize,
) -> Vec<Lot> {
    // open workbook if existed, otherwise exit program with code 1
    let mut workbook: Xlsx<_> = match open_workbook(file_path) {
        Ok(wb) => wb,
        Err(_) => {
            eprintln!("\x1b[91mERROR: File is does not exist!\x1b[0m");
            std::process::exit(1);
        }
    };

    let mut lots = Vec::new();

    for sheet in sheets {
        match workbook.worksheet_range(sheet) {
            Ok(range) => {
                // get directly cell value which contain charge number
                let charge = if let Some(value) = range.get_value((charge_row, charge_col)) {
                    value.as_string().unwrap()
                } else {
                    eprintln!(
                        "\x1b[93mWARNING: Sheet \"{}\" is exist but target FAUF cell is empty.\x1b[0m",
                        sheet
                    );
                    continue;
                };

                // get directly cell value which contain fauf number
                let fauf = if let Some(value) = range.get_value((fauf_row, fauf_col)) {
                    value.as_string().unwrap()
                } else {
                    eprintln!(
                        "\x1b[93mWARNING: Sheet \"{}\" is exist but target Charge cell is empty.\x1b[0m",
                        sheet
                    );
                    continue;
                };

                // iterate used cells and collects if it is in the correct row and column
                let serials = range
                    .used_cells()
                    .filter(|(r, c, _)| *r >= modul_range_row && *c == modul_range_column)
                    .map(|(_, _, v)| v.as_string().unwrap())
                    .collect::<Vec<String>>();

                println!(
                    "\x1b[96msheet: {} -> {}/{} {:#?}\x1b[0m",
                    sheet, &fauf, &charge, serials
                );
                lots.push(Lot::new(sheet.to_string(), fauf, charge, serials));
            }
            Err(_) => eprintln!(
                "\x1b[93mWARNING: Sheet \"{}\" is does not exist.\x1b[0m",
                sheet
            ),
        };
    }
    lots
}

/// Write file from the `cli::Lot` parameters
///
/// # Arguments
///
/// * `lots`: `cli::Lot` slice
/// * `path`: target file full path
/// * `f`: function which determines the file content
fn write_file<F>(lots: &[Lot], path: &str, f: F)
where
    F: Fn(&[Lot]) -> String,
{
    let mut file = fs::OpenOptions::new()
        .write(true)
        .create(true)
        .truncate(true)
        .open(path)
        .expect("Cannot open file");

    if file.write_all(f(lots).as_bytes()).is_ok() {
        println!("\x1b[92mFile write is succesfully. \x1b[4m{}\x1b[0m", path);
    } else {
        eprintln!("\x1b[91mCannot write to file.\x1b[0m");
    }
}

/// Create and format a `String` from the `&[cli::Lot]` parameters
///
/// # Arguments
///
/// * `lots`: `cli::Lot` slice
///
/// returns: `String`
fn create_list(lots: &[Lot]) -> String {
    let mut res = String::new();
    lots.iter()
        .for_each(|l| res.insert_str(res.len(), &format!("{} - {}\n", l.fauf, l.charge)));
    res
}

/// Create and format `String` from the `&[cli::Lot]` parameters
///
/// # Arguments
///
/// * `lots`: `cli::Lot` slice
///
/// returns: `String`
fn create_sql(lots: &[Lot]) -> String {
    let mut exec = String::new();
    lots.iter().for_each(|lot| {
        exec.insert_str(
            exec.len(),
            &format!(
                "--sheet: {}\n--charge: {}\n--item count: {}\nEXEC [dbo].[DDS_MoveItemsFromFaufToFauf]\n\t@FAUF = '{}',\n\t@MODUL = '{}';\n\n",
                lot.sheet,
                lot.charge,
                lot.serials.len(),
                lot.fauf,
                lot.serials.join(",")
            ),
        );
    });
    exec
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_create_lots() {
        let lots = vec![Lot::new(
            "41".to_string(),
            "82036419".to_string(),
            "TC03909P0".to_string(),
            vec![
                "1072501425237413".to_string(),
                "1072501423237413".to_string(),
                "1072501436237413".to_string(),
                "1072501438237413".to_string(),
                "1072501442237413".to_string(),
            ],
        )];

        let result = create_sql(&lots);

        let expected = String::from(
            "--sheet: 41\n--charge: TC03909P0\n--item count: 5\nEXEC [dbo].[DDS_MoveItemsFromFaufToFauf]\n\t@FAUF = '82036419',\n\t@MODUL = '1072501425237413,1072501423237413,1072501436237413,1072501438237413,1072501442237413';\n\n",
        );

        assert_eq!(expected, result);
    }
}
