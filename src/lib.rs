//! Simple library to convert JSON files to Excel (xlsx).
use std::io::{Cursor, Read, Seek, Write};

use serde_json::{Deserializer, Value};
use zip::write::FileOptions;

use crate::result::{XlsxExporResult, XlsxExportError};

pub mod result;

pub fn json_to_xlsx(reader: impl Read, mut output: impl Write) -> XlsxExporResult<()> {
    let mut stream = Deserializer::from_reader(reader).into_iter::<Value>();

    let main_array: Vec<Value> = match stream.next() {
        Some(Ok(Value::Array(list))) => list,
        Some(Ok(_)) => return Err(XlsxExportError::NotAnArray),
        Some(Err(e)) => return Err(XlsxExportError::JsonError(e)),
        None => return Err(XlsxExportError::NotAnArray),
    };

    if main_array.is_empty() {
        return Err(XlsxExportError::EmptyArray);
    }

    let first_item = main_array.first().unwrap();
    let first_item = match first_item {
        Value::Object(o) => o,
        _ => return Err(XlsxExportError::ExpectedObject),
    };
    let headers: Vec<String> = first_item.keys().cloned().collect();

    let mut buffer = Cursor::new(Vec::new());
    {
        let mut zip = zip::ZipWriter::new(&mut buffer);
        let options = FileOptions::default();

        write_content_types(&mut zip, options)?;
        write_rels(&mut zip, options)?;
        write_workbook(&mut zip, options)?;
        write_sheet1(&mut zip, options, &headers, main_array)?;

        zip.finish()?;
    }

    output.write_all(&buffer.into_inner())?;
    Ok(())
}

fn write_content_types<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    options: FileOptions,
) -> zip::result::ZipResult<()> {
    let xml = r#"
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
            <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        </Types>
    "#;
    zip.start_file("[Content_Types].xml", options)?;
    zip.write_all(xml.trim_start().as_bytes())
        .map_err(zip::result::ZipError::Io)
}

fn write_rels<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    options: FileOptions,
) -> zip::result::ZipResult<()> {
    zip.start_file("_rels/.rels", options)?;
    let xml = r#"
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
    "#;
    zip.write_all(xml.trim_start().as_bytes())
        .map_err(zip::result::ZipError::Io)
}

fn write_workbook<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    options: FileOptions,
) -> zip::result::ZipResult<()> {
    zip.start_file("xl/workbook.xml", options)?;
    let xml = r#"
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <sheets>
                <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
            </sheets>
        </workbook>
    "#;
    zip.write_all(xml.trim_start().as_bytes())
        .map_err(zip::result::ZipError::Io)?;

    // Add relationships
    zip.start_file("xl/_rels/workbook.xml.rels", options)?;
    let rels = r#"
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
        </Relationships>
    "#;
    zip.write_all(rels.trim_start().as_bytes())
        .map_err(zip::result::ZipError::Io)
}

fn write_sheet1<W: Write + Seek>(
    zip: &mut zip::ZipWriter<W>,
    options: FileOptions,
    headers: &[String],
    main_array: Vec<Value>,
) -> zip::result::ZipResult<()> {
    zip.start_file("xl/worksheets/sheet1.xml", options)?;

    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>"#);

    let mut row_idx = 1;

    // Write header
    xml.push_str(&format!(r#"<row r="{}">"#, row_idx));
    for (i, header) in headers.iter().enumerate() {
        let col = column_letter(i + 1);
        xml.push_str(&format!(
            r#"<c r="{}{}" t="str"><v>{}</v></c>"#,
            col, row_idx, header
        ));
    }
    xml.push_str("</row>");
    row_idx += 1;

    // Write data rows
    for row in main_array {
        let value: Value = row;
        if let Value::Object(o) = value {
            xml.push_str(&format!(r#"<row r="{}">"#, row_idx));
            for (i, key) in headers.iter().enumerate() {
                let col = column_letter(i + 1);
                let value = o.get(key).map_or("".to_string(), |v| {
                    v.to_string().trim_matches('"').to_string()
                });
                xml.push_str(&format!(
                    r#"<c r="{}{}" t="str"><v>{}</v></c>"#,
                    col, row_idx, value
                ));
            }
            xml.push_str("</row>");
        }

        row_idx += 1;
    }

    xml.push_str("</sheetData></worksheet>");
    zip.write_all(xml.as_bytes())
        .map_err(zip::result::ZipError::Io)
}

fn column_letter(mut n: usize) -> String {
    let mut col = String::new();
    while n > 0 {
        let rem = (n - 1) % 26;
        col.insert(0, (b'A' + rem as u8) as char);
        n = (n - 1) / 26;
    }
    col
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    fn run_and_extract_xlsx(json_input: &str) -> Vec<u8> {
        let reader = Cursor::new(json_input);
        let mut output = Vec::new();
        let result = json_to_xlsx(reader, &mut output);
        assert!(
            result.is_ok(),
            "Expected OK but got error: {:?}",
            result.err()
        );
        output
    }

    #[test]
    fn test_valid_json_conversion() {
        let json = r#"
            [
                { "name": "Alice", "age": 30 },
                { "name": "Bob", "age": 25 }
            ]
        "#;

        let output = run_and_extract_xlsx(json);
        assert!(
            output.starts_with(b"PK"),
            "Expected XLSX (zip) to start with PK"
        );
        assert!(
            output.len() > 100,
            "Output seems too small to be valid XLSX"
        );
    }

    #[test]
    fn test_non_array_json_root() {
        let json = r#"{ "name": "Not an array" }"#;
        let reader = Cursor::new(json);
        let mut output = Vec::new();
        let result = json_to_xlsx(reader, &mut output);
        assert!(matches!(result, Err(XlsxExportError::NotAnArray)));
    }

    #[test]
    fn test_empty_array() {
        let json = r#"[]"#;
        let reader = Cursor::new(json);
        let mut output = Vec::new();
        let result = json_to_xlsx(reader, &mut output);
        assert!(matches!(result, Err(XlsxExportError::EmptyArray)));
    }

    #[test]
    fn test_array_with_non_object_elements() {
        let json = r#"[1, 2, 3]"#;
        let reader = Cursor::new(json);
        let mut output = Vec::new();
        let result = json_to_xlsx(reader, &mut output);
        assert!(matches!(result, Err(XlsxExportError::ExpectedObject)));
    }

    #[test]
    fn test_malformed_json() {
        let json = r#"[{ "name": "John""#; // Missing closing brace
        let reader = Cursor::new(json);
        let mut output = Vec::new();
        let result = json_to_xlsx(reader, &mut output);
        assert!(matches!(result, Err(XlsxExportError::JsonError(_))));
    }

    #[test]
    fn test_missing_fields_in_some_objects() {
        let json = r#"
            [
                { "name": "Alice", "age": 30 },
                { "name": "Bob" }
            ]
        "#;

        let output = run_and_extract_xlsx(json);
        assert!(output.starts_with(b"PK"));
    }
}
