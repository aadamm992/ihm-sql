use clap::Parser;

#[derive(Parser)]
#[command(version, about = "Create SQL query from excel", long_about = None)]
pub struct Cli {
    /// File path
    pub file: String,

    /// Sheet names. Separated by comma
    #[arg(value_delimiter = ',', num_args = 1..)]
    pub sheets: Vec<String>,

    /// Create file with exec query
    #[arg(short, long)]
    pub exec: bool,

    /// Create a text file with Charge and FAUF
    #[arg(short, long)]
    pub list: bool,

    /// Set FAUF number row
    #[arg(long, default_value_t = 1)]
    pub fauf_row: usize,

    /// Set FAUF number column
    #[arg(long, default_value_t = 1)]
    pub fauf_col: usize,

    /// Set Charge number row
    #[arg(long, default_value_t = 1)]
    pub charge_row: usize,

    /// Set Charge number column
    #[arg(long, default_value_t = 0)]
    pub charge_col: usize,

    /// Set modul range starting row
    #[arg(long, default_value_t = 1)]
    pub modul_range_row: usize,

    /// Set modul range starting column
    #[arg(long, default_value_t = 6)]
    pub modul_range_column: usize,
}
