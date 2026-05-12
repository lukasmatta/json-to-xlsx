//! Simple library to convert a JSON array of objects to an Excel (xlsx) file.
//!
//! The input must be a JSON array where every element is a JSON object. Keys
//! from the first object become the column headers. Missing fields in later
//! objects produce empty cells. Non-object elements are silently skipped.
use std::io::{Cursor, Read, Seek, Write};

use serde_json::{Deserializer, Value};
use zip::write::FileOptions;

use crate::result::{XlsxExportError, XlsxExportResult};

pub mod result;

/// Convert a JSON array of objects to an xlsx file.
///
/// `reader` must contain a JSON array of objects. The keys of the first object
/// determine the column headers and their order. `output` receives the raw xlsx
/// bytes (a zip archive).
///
/// # Errors
///
/// Returns [`XlsxExportError::NotAnArray`] if the root value is not an array,
/// [`XlsxExportError::EmptyArray`] if the array is empty,
/// [`XlsxExportError::ExpectedObject`] if the first element is not an object,
/// [`XlsxExportError::JsonError`] on malformed JSON, and
/// [`XlsxExportError::IoError`] / [`XlsxExportError::ZipError`] on write failures.
///
/// # Examples
///
/// Write to a file:
///
/// ```no_run
/// use std::fs::File;
/// use std::io::BufReader;
/// use json_to_xlsx::json_to_xlsx;
///
/// fn main() -> Result<(), Box<dyn std::error::Error>> {
///     let input = BufReader::new(File::open("input.json")?);
///     let output = File::create("output.xlsx")?;
///     json_to_xlsx(input, output)?;
///     Ok(())
/// }
/// ```
///
/// Write to an in-memory buffer (useful for HTTP responses or tests):
///
/// ```
/// use std::io::Cursor;
/// use json_to_xlsx::json_to_xlsx;
///
/// fn main() -> Result<(), Box<dyn std::error::Error>> {
///     let json = br#"[{"name":"Alice","score":99}]"#;
///     let mut buf = Vec::new();
///     json_to_xlsx(Cursor::new(json), &mut buf)?;
///     // buf now contains a valid .xlsx file
///     Ok(())
/// }
/// ```
pub fn json_to_xlsx(reader: impl Read, mut output: impl Write) -> XlsxExportResult<()> {
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
            col,
            row_idx,
            xml_escape(header)
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
                let cell_ref = format!("{}{}", col, row_idx);
                let cell = match o.get(key) {
                    None | Some(Value::Null) => String::new(),
                    Some(v) => cell_xml(&cell_ref, v),
                };
                xml.push_str(&cell);
            }
            xml.push_str("</row>");
            row_idx += 1;
        }
    }

    xml.push_str("</sheetData></worksheet>");
    zip.write_all(xml.as_bytes())
        .map_err(zip::result::ZipError::Io)
}

fn cell_xml(cell_ref: &str, value: &Value) -> String {
    match value {
        Value::String(s) => format!(
            r#"<c r="{}" t="str"><v>{}</v></c>"#,
            cell_ref,
            xml_escape(s)
        ),
        Value::Number(n) => format!(r#"<c r="{}"><v>{}</v></c>"#, cell_ref, n),
        Value::Bool(b) => format!(
            r#"<c r="{}" t="b"><v>{}</v></c>"#,
            cell_ref,
            if *b { 1 } else { 0 }
        ),
        Value::Array(_) | Value::Object(_) => {
            let s = value.to_string();
            format!(
                r#"<c r="{}" t="str"><v>{}</v></c>"#,
                cell_ref,
                xml_escape(&s)
            )
        }
        Value::Null => String::new(),
    }
}

fn xml_escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for ch in s.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            c => out.push(c),
        }
    }
    out
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

    fn extract_sheet_xml(xlsx_bytes: Vec<u8>) -> String {
        let cursor = Cursor::new(xlsx_bytes);
        let mut archive = zip::ZipArchive::new(cursor).expect("valid zip");
        let mut file = archive
            .by_name("xl/worksheets/sheet1.xml")
            .expect("sheet1.xml present");
        let mut contents = String::new();
        file.read_to_string(&mut contents).expect("valid utf-8");
        contents
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
    fn test_xml_special_characters_in_values() {
        let json = r#"[{"name": "Tom & Jerry", "note": "<b>bold</b>"}]"#;
        let xml = extract_sheet_xml(run_and_extract_xlsx(json));
        assert!(xml.contains("Tom &amp; Jerry"), "& should be escaped");
        assert!(
            xml.contains("&lt;b&gt;bold&lt;/b&gt;"),
            "< > should be escaped"
        );
    }

    #[test]
    fn test_xml_special_characters_in_headers() {
        let json = r#"[{"a & b": 1, "<col>": 2}]"#;
        let xml = extract_sheet_xml(run_and_extract_xlsx(json));
        assert!(xml.contains("a &amp; b"), "& in header should be escaped");
        assert!(
            xml.contains("&lt;col&gt;"),
            "< > in header should be escaped"
        );
    }

    #[test]
    fn test_number_cells_have_no_type_attr() {
        let json = r#"[{"score": 42, "ratio": 3.14}]"#;
        let xml = extract_sheet_xml(run_and_extract_xlsx(json));
        // Data is in row 2; number cells must have no t= attribute
        assert!(
            xml.contains(r#"<c r="A2"><v>42</v></c>"#),
            "integer cell should have no type attr"
        );
        assert!(
            xml.contains(r#"<c r="B2"><v>3.14</v></c>"#),
            "float cell should have no type attr"
        );
    }

    #[test]
    fn test_boolean_cells() {
        let json = r#"[{"active": true, "deleted": false}]"#;
        let xml = extract_sheet_xml(run_and_extract_xlsx(json));
        assert!(xml.contains("t=\"b\""), "booleans should use t=b");
        assert!(xml.contains("<v>1</v>"), "true should be 1");
        assert!(xml.contains("<v>0</v>"), "false should be 0");
    }

    #[test]
    fn test_null_produces_empty_cell() {
        let json = r#"[{"name": "Alice", "age": null}]"#;
        let xml = extract_sheet_xml(run_and_extract_xlsx(json));
        assert!(
            !xml.contains("null"),
            "null should not appear as the string 'null'"
        );
    }

    #[test]
    fn test_string_with_quotes_escaped() {
        let json = r#"[{"msg": "say \"hi\""}]"#;
        let xml = extract_sheet_xml(run_and_extract_xlsx(json));
        assert!(
            xml.contains("say &quot;hi&quot;"),
            "quotes inside strings should be XML-escaped"
        );
    }

    #[test]
    fn test_non_object_rows_skipped_without_gap() {
        // Mixed array: valid object, non-object, valid object.
        // The non-object should be silently skipped and the third row
        // should appear as row 3 (not row 4 with a gap).
        let json = r#"[{"a": 1}, "oops", {"a": 2}]"#;
        let xml = extract_sheet_xml(run_and_extract_xlsx(json));
        assert!(
            xml.contains(r#"<row r="2">"#),
            "first data row should be row 2"
        );
        assert!(
            xml.contains(r#"<row r="3">"#),
            "second data row should be row 3, no gap"
        );
        assert!(!xml.contains(r#"<row r="4">"#), "there should be no row 4");
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
