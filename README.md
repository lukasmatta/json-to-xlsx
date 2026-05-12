# json-to-xlsx

[![Crates.io](https://img.shields.io/crates/v/json-to-xlsx.svg)](https://crates.io/crates/json-to-xlsx)
[![Crates.io downloads](https://img.shields.io/crates/d/json-to-xlsx.svg)](https://crates.io/crates/json-to-xlsx)
[![docs.rs](https://img.shields.io/docsrs/json-to-xlsx)](https://docs.rs/json-to-xlsx)

Convert a JSON array of objects to an Excel (`.xlsx`) file.

## Installation

```toml
[dependencies]
json-to-xlsx = "0.1"
```

## Usage

```rust
use std::fs::File;
use std::io::BufReader;
use json_to_xlsx::json_to_xlsx;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = BufReader::new(File::open("input.json")?);
    let output = File::create("output.xlsx")?;
    json_to_xlsx(input, output)?;
    Ok(())
}
```

## Input format

The JSON must be an array of objects. Keys from the **first** object become the column headers (in insertion order). Later rows can omit fields — those cells will be empty. Non-object elements are silently skipped.

```json
[
  { "name": "Alice", "age": 30, "active": true },
  { "name": "Bob",   "age": 25 }
]
```

## Type mapping

| JSON type | Excel cell |
| --- | --- |
| string | string |
| number | numeric |
| boolean | `1` / `0` |
| null | empty |
| object / array | serialized as string |

## Error handling

`json_to_xlsx` returns `XlsxExportError` on failure:

- `NotAnArray` — root value is not a JSON array
- `EmptyArray` — the array has no elements
- `ExpectedObject` — first element is not an object
- `JsonError` — malformed JSON
- `IoError` / `ZipError` — write failure

## Benchmarks

Benchmarks cover 100 / 10 000 / 100 000 rows at 6 columns each.

```sh
cargo bench
```

An HTML report with graphs is written to `target/criterion/report/index.html`.
