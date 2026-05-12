use criterion::{BenchmarkId, Criterion, black_box, criterion_group, criterion_main};
use std::io::{BufReader, BufWriter, Cursor};

use json_to_xlsx::json_to_xlsx;

fn generate_json(rows: usize) -> String {
    let mut s = String::from("[");
    for i in 0..rows {
        if i > 0 {
            s.push(',');
        }
        s.push_str(&format!(
            r#"{{"id":{i},"name":"User {i}","email":"user{i}@example.com","city":"City {i}","score":{score},"active":{active}}}"#,
            i = i,
            score = i % 100,
            active = i % 2 == 0,
        ));
    }
    s.push(']');
    s
}

fn bench_by_row_count(c: &mut Criterion) {
    let mut group = c.benchmark_group("json_to_xlsx");

    for rows in [100, 10_000, 100_000] {
        let json = generate_json(rows);
        group.bench_with_input(BenchmarkId::from_parameter(rows), &json, |b, json| {
            b.iter(|| {
                let reader = BufReader::new(Cursor::new(json.as_bytes()));
                let mut output_buf = Vec::new();
                let writer = BufWriter::new(Cursor::new(&mut output_buf));
                let _ = json_to_xlsx(black_box(reader), black_box(writer));
            });
        });
    }

    group.finish();
}

criterion_group!(benches, bench_by_row_count);
criterion_main!(benches);
