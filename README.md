json-to-xlsx
========
[![Crates.io version](https://img.shields.io/crates/v/json-to-xlsx.svg)](https://crates.io/crates/json-to-xlsx)

Info
----

Simple library to convert JSON files to Excel (xlsx).

How to use
----

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
