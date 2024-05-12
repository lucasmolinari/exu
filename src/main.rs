use std::env;
use std::ffi::OsStr;
use std::fs;
use std::io;
use std::io::Write;
use std::io::{Cursor, Read};
use std::path::{Path, PathBuf};

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use quick_xml::Writer;
use zip::read::ZipArchive;
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

type Result<T> = io::Result<T>;

fn main() {
    let args: Vec<String> = env::args().collect();

    if args.len() != 3 {
        eprintln!("Expected 2 arguments but {} were supplied.", args.len() - 1);
        println!("Usage: -- <source path> <destination path>");
        std::process::exit(0);
    }

    let src_path = match Path::new(args.get(1).unwrap()).canonicalize() {
        Ok(p) => {
            if p.extension().and_then(OsStr::to_str) != Some("xlsx") {
                eprintln!("File extension should be xlsx.");
                std::process::exit(0);
            }
            p
        }
        Err(_) => {
            eprintln!("Couldn't find the specified file.");
            std::process::exit(0);
        }
    };

    let dest_path = match Path::new(args.get(2).unwrap()).canonicalize() {
        Ok(p) => p,
        Err(_) => {
            eprintln!("Couldn't find destination path.");
            std::process::exit(0);
        }
    };

    read_zipped(&src_path, &dest_path).unwrap();
}

fn create_temp(src_path: &PathBuf) -> Result<tempfile::NamedTempFile> {
    let mut srcf = fs::File::open(src_path)?;
    let mut tmpf = tempfile::NamedTempFile::new()?;
    io::copy(&mut srcf, &mut tmpf)?;
    Ok(tmpf)
}

fn read_zipped(src: &PathBuf, dest: &PathBuf) -> Result<()> {
    let tsrc = create_temp(src)?;
    let mut arch = ZipArchive::new(tsrc)?;

    let rdest = fs::File::create(dest)?;
    let mut warch = ZipWriter::new(rdest);

    for i in 0..arch.len() {
        let mut file = arch.by_index(i)?;

        let options =
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
        warch.start_file(file.name(), options)?;

        if file.name().starts_with("xl/worksheets/sheet") {
            println!("{}", file.name());
            let mut xml = String::new();
            file.read_to_string(&mut xml)?;

            let unlocked = unlock(&xml)?;
            warch.write(unlocked.as_bytes())?;
            continue;
        }
    }
    Ok(())
}

fn unlock(xml: &str) -> Result<String> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
            Ok(Event::Eof) => break,
            Ok(Event::Empty(e)) => {
                if e.name().as_ref() == b"sheetProtection" {
                    continue;
                }
            }
            Ok(e) => {
                writer.write_event(e).unwrap();
            }
        };
    }
    let r = String::from_utf8(writer.into_inner().into_inner());
    match r {
        Ok(s) => Ok(s),
        Err(e) => Err(io::Error::new(io::ErrorKind::InvalidData, e)),
    }
}
