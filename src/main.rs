use calamine::{open_workbook, Reader, Xlsx};
use clap::Parser;
use dialoguer::console::Term;
use dialoguer::{theme::ColorfulTheme, Select};
use encoding_rs::WINDOWS_1252;
use encoding_rs_io::DecodeReaderBytesBuilder;
use std::error::Error;
use std::fs;
use std::io::{self, Read, Seek, Write};
use std::ops::RangeInclusive;

#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
struct Args {
    #[arg(short, long, help = "Path to detailed permission report text file", value_hint=clap::ValueHint::FilePath)]
    license: String,
    #[arg(short, long, help = "Path to exported objects in xlsx format", value_hint=clap::ValueHint::FilePath)]
    objects: String,
}

#[derive(Debug, PartialEq, Eq)]
enum ObjectType {
    TableData,
    Table,
    Report,
    Codeunit,
    XMLport,
    MenuSuite,
    Page,
    Query,
    System,
    FieldNumber,
    PageExtension,
    TableExtension,
    Enum,
    EnumExtension,
    Profile,
    ProfileExtension,
    PermissionSet,
    PermissionSetExtension,
    ReportExtension,
}

impl ObjectType {
    pub fn from(object_type: &str) -> Self {
        match object_type {
            "TableData" => Self::TableData,
            "Table" => Self::Table,
            "Report" => Self::Report,
            "Codeunit" => Self::Codeunit,
            "XMLport" | "XMLPort" => Self::XMLport,
            "MenuSuite" => Self::MenuSuite,
            "Page" => Self::Page,
            "Query" => Self::Query,
            "System" => Self::System,
            "FieldNumber" => Self::FieldNumber,
            "PageExtension" => Self::PageExtension,
            "TableExtension" => Self::TableExtension,
            "Enum" => Self::Enum,
            "EnumExtension" => Self::EnumExtension,
            "Profile" => Self::Profile,
            "ProfileExtension" => Self::ProfileExtension,
            "PermissionSet" => Self::PermissionSet,
            "PermissionSetExtension" => Self::PermissionSetExtension,
            "ReportExtension" => Self::ReportExtension,
            _ => unimplemented!("Unknown object type {object_type}!"),
        }
    }

    pub fn to_string(self: &Self) -> &str {
        match self {
            ObjectType::TableData => "TableData",
            ObjectType::Table => "Table",
            ObjectType::Report => "Report",
            ObjectType::Codeunit => "Codeunit",
            // different case for P is intentional
            ObjectType::XMLport => "XMLPort",
            ObjectType::MenuSuite => "MenuSuite",
            ObjectType::Page => "Page",
            ObjectType::Query => "Query",
            ObjectType::System => "System",
            ObjectType::FieldNumber => "FieldNumber",
            ObjectType::PageExtension => "PageExtension",
            ObjectType::TableExtension => "TableExtension",
            ObjectType::Enum => "Enum",
            ObjectType::EnumExtension => "EnumExtension",
            ObjectType::Profile => "Profile",
            ObjectType::ProfileExtension => "ProfileExtension",
            ObjectType::PermissionSet => "PermissionSet",
            ObjectType::PermissionSetExtension => "PermissionSetExtension",
            ObjectType::ReportExtension => "ReportExtension",
        }
    }

    pub fn is_licensed(self: &Self) -> bool {
        match self {
            ObjectType::TableData
            | ObjectType::Report
            | ObjectType::Codeunit
            | ObjectType::XMLport
            | ObjectType::Query
            | ObjectType::Page => true,

            ObjectType::Table
            | ObjectType::MenuSuite
            | ObjectType::System
            | ObjectType::FieldNumber
            | ObjectType::PageExtension
            | ObjectType::TableExtension
            | ObjectType::Enum
            | ObjectType::EnumExtension
            | ObjectType::Profile
            | ObjectType::ProfileExtension
            | ObjectType::PermissionSet
            | ObjectType::PermissionSetExtension
            | ObjectType::ReportExtension => false,
        }
    }
}

#[derive(Debug)]
struct ObjectRange {
    object_type: ObjectType,
    quantity: i64,
    range_from: i64,
    range_to: i64,
    permission: String,
}

impl ObjectRange {
    pub fn new(object_type: &str, range_from: i64, range_to: i64, permission: &str) -> Self {
        Self {
            object_type: ObjectType::from(object_type),
            quantity: range_to - range_from + 1,
            range_from,
            range_to,
            permission: permission.to_owned(),
        }
    }
}

#[derive(Debug)]
struct Object {
    object_type: ObjectType,
    id: i64,
    name: String,
}

impl Object {
    pub fn new(object_type: &str, id: i64, name: &str) -> Self {
        Self {
            object_type: ObjectType::from(object_type),
            id,
            name: name.to_owned(),
        }
    }
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
        ObjectRange::new("TableData", 50000, 50009, "RIMDX"),
        ObjectRange::new("Page", 50000, 50099, "X"),
        ObjectRange::new("Report", 50000, 50099, "X"),
        ObjectRange::new("Codeunit", 50000, 50099, "X"),
        ObjectRange::new("XMLPort", 50000, 50099, "X"),
        ObjectRange::new("Query", 50000, 50099, "X"),
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
        if let &[object_type, _, range_from, range_to, permission] =
            words.collect::<Vec<&str>>().as_slice()
        {
            licensed_object_ranges.push(ObjectRange::new(
                object_type,
                range_from.parse::<i64>()?,
                range_to.parse::<i64>()?,
                permission,
            ));
        } else {
            unimplemented!("Unimplemented license format.");
        }
    }

    let mut excel: Xlsx<_> = open_workbook(args.objects).expect("Could not read the objects file!");
    let selected_sheet = pick_sheet(&excel)?;

    if let Some(Ok(r)) = excel.worksheet_range(&selected_sheet) {
        for row in r.rows().skip(1) {
            if let [object_type, object_id, name, ..] = row {
                objects.push(Object::new(
                    &object_type.to_string(),
                    if object_id.is_int() {
                        object_id.get_int().unwrap()
                    } else if object_id.is_float() {
                        object_id.get_float().unwrap() as i64
                    } else {
                        unimplemented!("Object id is not a number {}!", object_id.to_string());
                    },
                    &name.to_string(),
                ));
            } else {
                unimplemented!("Unimplemented row format.");
            }
        }
    }

    let mut missing_objects: Vec<Object> = Vec::new();

    for object in objects
        .into_iter()
        .filter(|e| e.object_type.is_licensed())
        .filter(|e| checked_range.contains(&e.id))
    {
        let found_index = licensed_object_ranges.iter().position(|e| {
            e.object_type == object.object_type && (e.range_from..=e.range_to).contains(&object.id)
        });

        match found_index {
            Some(_) => {}
            None => {
                missing_objects.push(object);
            }
        }
    }

    // TODO print stats - how many objects found?

    if missing_objects.is_empty() {
        println!("No missing objects found!");
        return Ok(());
    }

    let path = "missing-permissions.csv";

    let file = fs::File::create(&path)?;
    let mut file = io::LineWriter::new(file);

    file.write_all(b"ObjectType,FromObjectID,ToObjectID,Read,Insert,Modify,Delete,Execute,AvailableRange,Used,ObjectTypeRemaining,CompanyObjectPermissionID\n")?;

    for object in &missing_objects {
        println!(
            "{} {}\t{}",
            object.id,
            object.object_type.to_string(),
            object.name
        );

        let object_id = object.id.to_string();
        let quantity = (object.id - object.id + 1).to_string();
        let line = vec![
            object.object_type.to_string(),
            &object_id,
            &object_id,
            "Direct",
            "Direct",
            "Direct",
            "Direct",
            "Direct",
            "50000 - 99999",
            &quantity,
            "0",
            "0",
        ];
        file.write_all(line.join(",").as_bytes())?;
        file.write_all(b"\n")?;
    }

    file.flush()?;

    println!("Wrote missing permissions to {path}");

    // TODO print objects that are not needed?

    // TODO make permission file input optional?

    Ok(())
}
