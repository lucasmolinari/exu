use std::env; 
use std::io;
use std::fs;
use std::path::{Path, PathBuf};
use std::ffi::OsStr;

type Result<T> = io::Result<T>;

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() != 2 {
        eprintln!("Expected 1 argument but {} were supplied.", args.len() - 1);
        std::process::exit(0);
    }
    let source_path = match Path::new(args.get(1).unwrap()).canonicalize() {
        Ok(p) => {
            if p.extension().and_then(OsStr::to_str) != Some("xlsx") {
                eprintln!("File extension should be xlsx.");
                std::process::exit(0);
            }
            p
        },
        Err(_) => {
            eprintln!("Couldn't find the specified file.");
            std::process::exit(0);
        }
    };
    let file = match create_temp(&source_path) {
        Ok(f) => f,
        Err(e) => {
            eprintln!("Couldn't create temporary file: {}", e);
            std::process::exit(0);
        }
    };
    let contents = fs::read_to_string(file).unwrap();
    dbg!(contents);
}

fn create_temp(src_path: &PathBuf) -> Result<tempfile::NamedTempFile> {
    let mut srcf = fs::File::open(src_path)?;
    let mut tmpf = tempfile::NamedTempFile::new()?;
    io::copy(&mut srcf, &mut tmpf)?;
    Ok(tmpf)
}

