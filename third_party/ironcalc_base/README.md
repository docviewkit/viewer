# IronCalc Base

## DocViewKit integration

DocViewKit vendors IronCalc Base 0.8.2 for its optional spreadsheet calculation Wasm module. The upstream `MIT OR Apache-2.0` license declared in [Cargo.toml](Cargo.toml) remains in force; the retained MIT text is [ironcalc-MIT.txt](../licenses/ironcalc-MIT.txt). DocViewKit's project license does not replace upstream copyrights. Use the [root build and test instructions](../../README.md) for the integrated runtime; the upstream library notes follow below.

[![Crates.io][crates-badge]][crates-url]
[![MIT licensed][mit-badge]][mit-url]

[crates-badge]: https://img.shields.io/crates/v/ironcalc_base.svg
[crates-url]: https://crates.io/crates/ironcalc_base
[mit-badge]: https://img.shields.io/badge/license-MIT-blue.svg
[mit-url]: https://github.com/ironcalc/ironcalc/blob/master/LICENSE


## About

IronCalc Base is the engine of the IronCalc ecosystem


## Build
To build the library

```bash
$ cargo build --release
```

## Tests

To run the tests:

```bash
$ cargo test
```
