use criterion::{Criterion, black_box, criterion_group, criterion_main};
use std::io::{BufReader, BufWriter, Cursor};

use json_to_xlsx::json_to_xlsx;

fn conversion_benchmark(c: &mut Criterion) {
    let json = r#"[
        { "city": "New York", "temp": 25 },
        { "city": "Prague", "temp": 8 },
        { "city": "Tokyo", "temp": 23 },
        { "city": "Cape Town", "temp": 16 }
     ]"#;

    c.bench_function("json_to_xlsx_mem", |b| {
        b.iter(|| {
            let input = Cursor::new(json.as_bytes());
            let mut output_buf = Vec::new();
            let output = Cursor::new(&mut output_buf);

            let reader = BufReader::new(input);
            let writer = BufWriter::new(output);

            let _ = json_to_xlsx(black_box(reader), black_box(writer));
        })
    });
}

criterion_group!(benches, conversion_benchmark);
criterion_main!(benches);
