use calamine::{open_workbook_from_rs, Reader, Xlsx};
use itertools::Itertools;
use std::{
    error::Error,
    io::{Read, Seek},
    ops::RangeInclusive,
};

#[derive(Debug, PartialEq, Eq, Hash, Clone, Copy)]
pub enum ObjectType {
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
    pub fn from(object_type: &str) -> Result<Self, String> {
        match object_type {
            "TableData" => Ok(Self::TableData),
            "Table" => Ok(Self::Table),
            "Report" => Ok(Self::Report),
            "Codeunit" => Ok(Self::Codeunit),
            "XMLport" | "XMLPort" => Ok(Self::XMLport),
            "MenuSuite" => Ok(Self::MenuSuite),
            "Page" => Ok(Self::Page),
            "Query" => Ok(Self::Query),
            "System" => Ok(Self::System),
            "FieldNumber" => Ok(Self::FieldNumber),
            "PageExtension" => Ok(Self::PageExtension),
            "TableExtension" => Ok(Self::TableExtension),
            "Enum" => Ok(Self::Enum),
            "EnumExtension" => Ok(Self::EnumExtension),
            "Profile" => Ok(Self::Profile),
            "ProfileExtension" => Ok(Self::ProfileExtension),
            "PermissionSet" => Ok(Self::PermissionSet),
            "PermissionSetExtension" => Ok(Self::PermissionSetExtension),
            "ReportExtension" => Ok(Self::ReportExtension),
            _ => Err(format!("Invalid object type {}", object_type)),
        }
    }

    pub fn format(&self) -> &str {
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

    pub fn is_licensed(&self) -> bool {
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
pub struct ObjectRange {
    pub object_type: ObjectType,
    pub quantity: i64,
    pub range_from: i64,
    pub range_to: i64,
}

impl ObjectRange {
    pub fn new(object_type: &str, range_from: i64, range_to: i64) -> Result<Self, String> {
        Ok(Self {
            object_type: ObjectType::from(object_type)?,
            quantity: range_to - range_from + 1,
            range_from,
            range_to,
        })
    }

    pub fn new_with_type(object_type: ObjectType, range_from: i64, range_to: i64) -> Self {
        Self {
            object_type,
            quantity: range_to - range_from + 1,
            range_from,
            range_to,
        }
    }

    pub fn start_new(object_type: ObjectType, id: i64) -> Self {
        Self {
            object_type,
            quantity: 1,
            range_from: id,
            range_to: id,
        }
    }

    pub fn increase_range_to(&mut self) {
        self.range_to += 1;
        self.quantity += 1;
    }
}

#[derive(Debug)]
pub struct Object {
    pub object_type: ObjectType,
    pub id: i64,
    pub name: String,
}

impl Object {
    pub fn new(object_type: &str, id: i64, name: &str) -> Result<Self, String> {
        Ok(Self {
            object_type: ObjectType::from(object_type)?,
            id,
            name: name.to_owned(),
        })
    }
}

fn pick_sheet<RS: Read + Seek>(excel: &Xlsx<RS>) -> Result<String, &'static str> {
    let sheet_names = excel.sheet_names();

    sheet_names
        .first()
        .map(|e| e.into())
        .ok_or("Invalid input: No sheets.")
}

fn merge_missing_objects(missing_objects: &[Object]) -> Vec<ObjectRange> {
    let groups = missing_objects.iter().into_group_map_by(|e| e.object_type);

    let mut ranges: Vec<ObjectRange> = vec![];

    for (object_type, mut objects) in groups {
        objects.sort_by_key(|e| e.id);

        let first_object = objects.first().unwrap();

        let mut range = ObjectRange::start_new(object_type, first_object.id);

        for object in objects.iter().skip(1) {
            if object.id == range.range_to + 1 {
                range.increase_range_to();
                continue;
            }

            ranges.push(range);
            range = ObjectRange::start_new(object_type, object.id);
        }

        ranges.push(range);
    }

    ranges
}

fn parse_license(license: Option<String>) -> Result<Vec<ObjectRange>, Box<dyn Error>> {
    let mut licensed_object_ranges: Vec<ObjectRange> = Vec::from([
        ObjectRange::new_with_type(ObjectType::TableData, 50000, 50009),
        ObjectRange::new_with_type(ObjectType::Page, 50000, 50099),
        ObjectRange::new_with_type(ObjectType::Report, 50000, 50099),
        ObjectRange::new_with_type(ObjectType::Codeunit, 50000, 50099),
        ObjectRange::new_with_type(ObjectType::XMLport, 50000, 50099),
        ObjectRange::new_with_type(ObjectType::Query, 50000, 50099),
    ]);

    if let Some(license) = license {
        let skip = license.lines().skip_while(|p| *p != "Object Assignment");

        for line in skip
            .skip(5)
            .take_while(|p| *p != "Module Objects and Permissions")
            .filter(|p| !p.is_empty())
        {
            let words = line.split_whitespace();
            if let &[object_type, _, range_from, range_to, _] =
                words.collect::<Vec<&str>>().as_slice()
            {
                licensed_object_ranges.push(ObjectRange::new(
                    object_type,
                    range_from.parse::<i64>()?,
                    range_to.parse::<i64>()?,
                )?);
            } else {
                return Err("Invalid license format: Expected object_type range_from range_to split by whitespace.".into());
            }
        }
    }

    Ok(licensed_object_ranges)
}

fn parse_objects<RS>(objects_reader: RS) -> Result<Vec<Object>, Box<dyn Error>>
where
    RS: Read + Seek,
{
    let mut objects: Vec<Object> = Vec::new();
    let mut excel: Xlsx<_> = open_workbook_from_rs(objects_reader)?;
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
                        return Err(format!(
                            "Invalid object id: Expected a number, got {}.",
                            object_id
                        )
                        .into());
                    },
                    &name.to_string(),
                )?);
            } else {
                return Err("Invalid objects row: Expected three columns.".into());
            }
        }
    }

    Ok(objects)
}

pub fn compare<RS>(
    license: Option<String>,
    objects_reader: RS,
) -> Result<(Vec<Object>, Vec<ObjectRange>), Box<dyn Error>>
where
    RS: Read + Seek,
{
    let licensed_object_ranges = parse_license(license)?;
    let objects = parse_objects(objects_reader)?;

    let checked_range: RangeInclusive<i64> = 50000..=99999;
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

    let missing_ranges = merge_missing_objects(&missing_objects);

    Ok((missing_objects, missing_ranges))
}
