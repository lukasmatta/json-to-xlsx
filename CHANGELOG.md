# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.4](https://github.com/lukasmatta/json-to-xlsx/compare/v0.1.3...v0.1.4) - 2026-05-12

### Other

- Improve description, add llms.txt
- Extend README.md
- Add downloads badge
- Improve benchmark

## [0.1.3](https://github.com/lukasmatta/json-to-xlsx/compare/v0.1.2...v0.1.3) - 2026-04-23

### Fixed

- skip non-object rows without gaps and add From<serde_json::Error>
- serialize JSON values by type instead of string coercion
- escape XML special characters in cell values and headers

### Other

- apply cargo fmt
- add function doc comment and fix README example to use ? operator

## [0.1.2](https://github.com/lukasmatta/json-to-xlsx/compare/v0.1.1...v0.1.2) - 2025-09-29

### Other

- add criterion benchmark
- Simplify sample code

## [0.1.1](https://github.com/lukasmatta/json-to-xlsx/compare/v0.1.0...v0.1.1) - 2025-07-05

### Other

- Update README.md
- Add crate version badge
