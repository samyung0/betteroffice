//! `seed_state <file.docx> <out.bin>`: today's DOCX seed (spike diagnostics).
fn main() {
    let args: Vec<String> = std::env::args().collect();
    let base = std::fs::read(&args[1]).unwrap();
    let seed = office_service::seed(office_service::Format::Docx, &base).unwrap();
    std::fs::write(&args[2], seed).unwrap();
}
