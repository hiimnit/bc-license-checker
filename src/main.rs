use calamine::{open_workbook, Reader, Xlsx};
use clap::Parser;
use dialoguer::console::Term;
use dialoguer::{theme::ColorfulTheme, Select};
use encoding_rs::WINDOWS_1252;
use encoding_rs_io::DecodeReaderBytesBuilder;
use std::error::Error;
use std::fs;
use std::io;
use std::io::Read;
use std::io::Seek;
use std::ops::RangeInclusive;

#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
struct Args {
    #[arg(short, long, help = "Path to detailed permission report text file", value_hint=clap::ValueHint::FilePath)]
    license: String,
    #[arg(short, long, help = "Path to exported objects in xlsx format", value_hint=clap::ValueHint::FilePath)]
    objects: String,
}

#[derive(Debug)]
struct ObjectRange {
    object_type: String,
    quantity: i64,
    range_from: i64,
    range_to: i64,
    permission: String,
}

#[derive(Debug)]
struct Object {
    object_type: String,
    id: i64,
    name: String,
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

    return Ok(result);
}

fn pick_sheet<RS: Read + Seek>(excel: &Xlsx<RS>) -> Result<String, &str> {
    let sheet_names = excel.sheet_names();

    match sheet_names.len() {
        0 => return Err("No sheets"),
        1 => return Ok(sheet_names.first().unwrap().clone()),
        _ => {
            let selection = Select::with_theme(&ColorfulTheme::default())
                .items(sheet_names)
                .default(0)
                .interact_on_opt(&Term::stderr())
                .or(Err("Terminal error"))?;

            return match selection {
                Some(index) => Ok(sheet_names[index].clone()),
                None => Err("Select a sheet!"),
            };
        }
    }
}

fn main() -> Result<(), Box<dyn Error>> {
    let args = Args::parse();

    let mut licensed_object_ranges: Vec<ObjectRange> = Vec::from([
        ObjectRange {
            object_type: "TableData".to_string(),
            quantity: 10,
            range_from: 50000,
            range_to: 50009,
            permission: "RIMDX".to_string(),
        },
        ObjectRange {
            object_type: "Page".to_string(),
            quantity: 100,
            range_from: 50000,
            range_to: 50099,
            permission: "X".to_string(),
        },
        ObjectRange {
            object_type: "Report".to_string(),
            quantity: 100,
            range_from: 50000,
            range_to: 50099,
            permission: "X".to_string(),
        },
        ObjectRange {
            object_type: "Codeunit".to_string(),
            quantity: 100,
            range_from: 50000,
            range_to: 50099,
            permission: "X".to_string(),
        },
        ObjectRange {
            object_type: "XMLPort".to_string(),
            quantity: 100,
            range_from: 50000,
            range_to: 50099,
            permission: "X".to_string(),
        },
        ObjectRange {
            object_type: "Query".to_string(),
            quantity: 100,
            range_from: 50000,
            range_to: 50099,
            permission: "X".to_string(),
        },
    ]);

    let checked_range: RangeInclusive<i64> = 50000..=99999;

    let mut objects: Vec<Object> = Vec::new();

    let license_file =
        read_file(&args.license, WINDOWS_1252).expect("Could not read the license info file!");

    let skip = license_file
        .lines()
        .into_iter()
        .skip_while(|p| *p != "Object Assignment");

    for line in skip
        .skip(5)
        .take_while(|p| *p != "Module Objects and Permissions")
        .filter(|p| *p != "")
    {
        let words = line.split_whitespace();
        if let &[object_type, quantity, range_from, range_to, permission] =
            words.collect::<Vec<&str>>().as_slice()
        {
            licensed_object_ranges.push(ObjectRange {
                object_type: String::from(object_type),
                quantity: quantity.parse::<i64>()?,
                range_from: range_from.parse::<i64>()?,
                range_to: range_to.parse::<i64>()?,
                permission: String::from(permission),
            });
        } else {
            unimplemented!("Unimplemented license format.");
        }
    }

    let mut excel: Xlsx<_> = open_workbook(args.objects).expect("Could not read the objects file!");
    let selected_sheet = pick_sheet(&excel)?;

    if let Some(Ok(r)) = excel.worksheet_range(&selected_sheet) {
        for row in r.rows().skip(1) {
            if let [object_type, object_id, name, ..] = row {
                objects.push(Object {
                    object_type: String::from(object_type.to_string()),
                    id: if object_id.is_int() {
                        object_id.get_int().unwrap()
                    } else if object_id.is_float() {
                        object_id.get_float().unwrap() as i64
                    } else {
                        -1
                    },
                    name: String::from(name.to_string()),
                });
            } else {
                unimplemented!("Unimplemented row format.");
            }
        }
    }

    for object in objects.iter().filter(|e| checked_range.contains(&e.id)) {
        let found_index = licensed_object_ranges.iter().position(|e| {
            e.object_type.to_lowercase() == object.object_type.to_lowercase()
                && (e.range_from..=e.range_to).contains(&object.id)
        });

        match found_index {
            Some(_) => {}
            None => println!(
                "{} {} {} not found",
                object.object_type, object.id, object.name
            ),
        }
    }

    Ok(())
}
