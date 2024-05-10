use std::env; 
use std::io;
use std::fs;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::ffi::OsStr;

use zip::read::ZipArchive;
use quick_xml::events::Event;
use quick_xml::reader::Reader;

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
    read_zipped(&file).unwrap();
}

fn create_temp(src_path: &PathBuf) -> Result<tempfile::NamedTempFile> {
    let mut srcf = fs::File::open(src_path)?;
    let mut tmpf = tempfile::NamedTempFile::new()?;
    io::copy(&mut srcf, &mut tmpf)?;
    Ok(tmpf)
}

fn read_zipped(file: &tempfile::NamedTempFile) -> Result<()> {
    let mut arch = ZipArchive::new(file).expect("Valid Zip File");
    
    for i in 0..arch.len() {
        let mut file = arch.by_index(i)?;

        if file.name().starts_with("xl/worksheets/sheet") {
            println!("{}", file.name());
            let mut xml = String::new();
            file.read_to_string(&mut xml)?;
            unlock(&xml)?;
        }
    }
    Ok(())
}

fn unlock(xml: &str) -> Result<()> {
    let mut reader = Reader::from_str(xml);

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf){
            Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
            Ok(Event::Eof) => break,
            Ok(Event::Start(ref e)) => println!("Found Start: {:?}", String::from_utf8_lossy(e)),
            Ok(Event::End(ref e)) => println!("Found End: {:?}", String::from_utf8_lossy(e)),
            Ok(Event::Text(ref e)) => {
            println!("Found: {:?}", String::from_utf8_lossy(e));
        }
        _ => {}
        };
    }
    Ok(())
}
