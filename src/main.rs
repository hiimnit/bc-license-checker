use clap::Parser;
use encoding_rs::WINDOWS_1252;
use encoding_rs_io::DecodeReaderBytesBuilder;
use std::error::Error;
use std::fs::{self, File};
use std::io::{self, BufReader, Read, Write};

#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
struct Args {
    #[arg(short, long, help = "Path to detailed permission report text file", value_hint=clap::ValueHint::FilePath)]
    license: Option<String>,
    #[arg(short, long, help = "Path to exported objects in xlsx format", value_hint=clap::ValueHint::FilePath)]
    objects: String,
}

fn read_file(
    path: &String,
    encoding: &'static encoding_rs::Encoding,
) -> Result<String, Box<dyn Error>> {
    let file = fs::File::open(path)?;
    let transcoded = DecodeReaderBytesBuilder::new()
        .encoding(Some(encoding))
        .build(file);
    let mut reader = io::BufReader::new(transcoded);

    let mut result = String::new();
    reader.read_to_string(&mut result)?;

    Ok(result)
}

fn main() -> Result<(), bclicensechecker::SendSyncError> {
    let args = Args::parse();

    let license = args.license.map(|license| {
        read_file(&license, WINDOWS_1252).expect("Could not read the license info file!")
    });
    let objects_reader =
        BufReader::new(File::open(args.objects).expect("Could not read the objects file!"));

    let (missing_objects, missing_ranges) = bclicensechecker::compare(license, objects_reader)?;

    if missing_objects.is_empty() {
        println!("No missing objects found!");
        return Ok(());
    }

    let output_file_path = "missing-permissions.csv";
    let output_file = fs::File::create(output_file_path)?;
    let mut output_file = io::LineWriter::new(output_file);

    output_file.write_all(b"ObjectType,FromObjectID,ToObjectID,Read,Insert,Modify,Delete,Execute,AvailableRange,Used,ObjectTypeRemaining,CompanyObjectPermissionID\n")?;

    for object in &missing_objects {
        println!(
            "{} {}\t{}",
            object.id,
            object.object_type.format(),
            object.name
        );
    }

    for range in missing_ranges {
        let line = [
            range.object_type.format(),
            &range.range_from.to_string(),
            &range.range_to.to_string(),
            "Direct",
            "Direct",
            "Direct",
            "Direct",
            "Direct",
            "50000 - 99999",
            &range.quantity.to_string(),
            "0",
            "0",
        ];
        output_file.write_all(line.join(",").as_bytes())?;
        output_file.write_all(b"\n")?;
    }

    output_file.flush()?;

    println!("Wrote missing permissions to {output_file_path}");

    // TODO print objects that are not needed?

    Ok(())
}
