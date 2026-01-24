#[derive(Debug, PartialOrd, PartialEq)]
pub struct Lot {
    pub sheet: String,
    pub fauf: String,
    pub charge: String,
    pub serials: Vec<String>,
}

impl Lot {
    pub fn new(sheet: String, fauf: String, charge: String, serials: Vec<String>) -> Lot {
        Lot {
            sheet,
            fauf,
            charge,
            serials,
        }
    }
}
