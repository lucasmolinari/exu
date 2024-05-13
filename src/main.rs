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
            let ext = p.extension().and_then(OsStr::to_str);
            if ext != Some("xlsx") && ext != Some("xlsm") {
                std::process::exit(0);
            }
            p
        }
        Err(e) => {
            eprintln!("Error reaching for the source path: {}.", e);
            std::process::exit(0);
        }
    };

    let uname = match src_path.file_name() {
        Some(name) => "unlocked_".to_string() + name.to_str().unwrap(),
        None => "unlocked.xlsx".to_string(),
    };
    let dest_path = Path::new(args.get(2).unwrap()).join(uname);

    match unlock(&src_path, &dest_path) {
        Ok(_) => println!("Process Finished."),
        Err(e) => eprintln!("{}", e),
    };
}

fn create_temp(src_path: &PathBuf) -> Result<tempfile::NamedTempFile> {
    let mut srcf = fs::File::open(src_path)?;
    let mut tmpf = tempfile::NamedTempFile::new()?;
    io::copy(&mut srcf, &mut tmpf)?;
    Ok(tmpf)
}

fn unlock(src: &PathBuf, dest: &PathBuf) -> Result<()> {
    let tsrc = create_temp(src)?;
    let mut arch = ZipArchive::new(tsrc)?;

    let rdest = fs::File::create(dest)?;
    let mut warch = ZipWriter::new(rdest);

    for i in 0..arch.len() {
        let mut file = arch.by_index(i)?;

        let options =
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
        warch.start_file(file.name(), options)?;

        let mut content = Vec::new();
        match file.read_to_end(&mut content) {
            Ok(_) => {}
            Err(e) => {
                println!("Received error trying to read [{}]: {}", file.name(), e);
                println!("File skipped.");
                continue;
            }
        };
        if file.name().starts_with("xl/worksheets/sheet") {
            let cstr = std::str::from_utf8(&content)
                .map_err(|e| io::Error::new(io::ErrorKind::InvalidData, e))?;

            let unlocked = remove_tag(&cstr, "sheetProtection")?;
            warch.write(unlocked.as_bytes())?;

            continue;
        }
        warch.write_all(&content)?;
    }
    Ok(())
}

fn remove_tag(xml: &str, tag: &str) -> Result<String> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Err(e) => panic!("Error at position {}: {:?}", reader.buffer_position(), e),
            Ok(Event::Eof) => break,
            Ok(Event::Empty(e)) => {
                if e.name().as_ref() != tag.as_bytes() {
                    writer.write_event(Event::Empty(e)).unwrap();
                }
            }
            Ok(e) => {
                writer.write_event(e).unwrap();
            }
        };
    }

    match String::from_utf8(writer.into_inner().into_inner()) {
        Ok(s) => Ok(s),
        Err(e) => Err(io::Error::new(io::ErrorKind::InvalidData, e)),
    }
}
