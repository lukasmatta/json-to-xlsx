json-to-xlsx
========
[![Crates.io version](https://img.shields.io/crates/v/json-to-xlsx.svg)](https://crates.io/crates/json-to-xlsx)

Info
----

Simple library to convert JSON files to Excel (xlsx).

How to use
----

```rust
let json_input = File::open("input.json").unwrap();
let xlsx_output = File::create("output.xlsx").unwrap();

let buf_reader = BufReader::new(json_input);

match json_to_xlsx(buf_reader, xlsx_output) {
    Ok(_) => println!("Success!"),
    Err(e) => println!("{e}"),
}
```
