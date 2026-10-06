//! Format-neutral spreadsheet calculation results shared by parsers and the lazy calculator.

use std::collections::{HashMap, HashSet};

const RESULT_MAGIC: &[u8; 8] = b"OVCALCV2";
pub const REQUIRED_MESSAGE: &str = "XLSX uncached or volatile formulas and conditional-format predicates require the optional calculation module";

pub(crate) fn calculated_condition(kind: &str) -> bool {
    matches!(
        kind,
        "cellIs"
            | "expression"
            | "containsText"
            | "notContainsText"
            | "beginsWith"
            | "endsWith"
            | "containsBlanks"
            | "notContainsBlanks"
            | "containsErrors"
            | "notContainsErrors"
            | "duplicateValues"
            | "uniqueValues"
            | "top10"
            | "aboveAverage"
            | "timePeriod"
    )
}

#[cfg(any(feature = "native-formats", feature = "calculation-service"))]
pub(crate) fn formula_calls_any(formula: &str, names: &[&str]) -> bool {
    let upper = formula.to_ascii_uppercase();
    names.iter().any(|name| {
        upper.match_indices(name).any(|(start, matched)| {
            let end = start + matched.len();
            (start == 0
                || !matches!(upper.as_bytes()[start - 1], b'A'..=b'Z' | b'0'..=b'9' | b'_' | b'.'))
                && upper.as_bytes().get(end) == Some(&b'(')
        })
    })
}

#[derive(Clone, Copy, Debug, Eq, Hash, Ord, PartialEq, PartialOrd)]
pub struct CellAddress {
    pub sheet: u32,
    pub row: i32,
    pub column: i32,
}

#[derive(Clone, Debug, PartialEq)]
pub enum FormulaValue {
    Blank,
    Boolean(bool),
    Error(String),
    Number(f64),
    Text(String),
}

#[derive(Clone, Debug, Default)]
pub struct CalculationResults {
    values: HashMap<CellAddress, FormulaValue>,
    array_cells: HashSet<CellAddress>,
    skipped_cells: HashSet<CellAddress>,
    conditional_matches: HashMap<(u32, u32), HashSet<(u32, u32)>>,
}

impl CalculationResults {
    pub fn insert_conditional_rule(&mut self, sheet: u32, priority: u32) {
        self.conditional_matches
            .entry((sheet, priority))
            .or_default();
    }

    pub fn insert_conditional_match(&mut self, address: CellAddress, priority: u32) {
        if address.row > 0 && address.column > 0 {
            self.conditional_matches
                .entry((address.sheet, priority))
                .or_default()
                .insert((address.column as u32 - 1, address.row as u32 - 1));
        }
    }

    pub fn conditional_matches(&self, sheet: u32, priority: u32) -> Option<&HashSet<(u32, u32)>> {
        self.conditional_matches.get(&(sheet, priority))
    }
    pub fn insert_value(&mut self, address: CellAddress, value: FormulaValue) {
        self.values.insert(address, value);
    }

    pub fn insert_array_cell(&mut self, address: CellAddress) {
        self.array_cells.insert(address);
    }

    pub fn insert_skipped_cell(&mut self, address: CellAddress) {
        self.skipped_cells.insert(address);
    }

    pub fn value(&self, address: CellAddress) -> Result<&FormulaValue, String> {
        if self.skipped_cells.contains(&address) {
            return Err("formula result depends on time or randomness".to_owned());
        }
        self.values
            .get(&address)
            .ok_or_else(|| "calculation result is missing".to_owned())
    }

    pub fn is_array_cell(&self, address: CellAddress) -> bool {
        self.array_cells.contains(&address)
    }

    pub fn encode(&self) -> Result<Vec<u8>, String> {
        let mut writer = Writer(RESULT_MAGIC.to_vec());
        let mut values = self.values.iter().collect::<Vec<_>>();
        values.sort_by_key(|(address, _)| **address);
        writer.count(values.len())?;
        for (address, value) in values {
            writer.address(*address);
            match value {
                FormulaValue::Blank => writer.u8(0),
                FormulaValue::Boolean(value) => {
                    writer.u8(1);
                    writer.u8(u8::from(*value));
                }
                FormulaValue::Error(value) => {
                    writer.u8(2);
                    writer.string(value)?;
                }
                FormulaValue::Number(value) => {
                    if !value.is_finite() {
                        return Err("calculation result is not finite".to_owned());
                    }
                    writer.u8(3);
                    writer.0.extend_from_slice(&value.to_le_bytes());
                }
                FormulaValue::Text(value) => {
                    writer.u8(4);
                    writer.string(value)?;
                }
            }
        }
        let mut array_cells = self.array_cells.iter().copied().collect::<Vec<_>>();
        array_cells.sort_unstable();
        writer.count(array_cells.len())?;
        for address in array_cells {
            writer.address(address);
        }
        let mut skipped_cells = self.skipped_cells.iter().copied().collect::<Vec<_>>();
        skipped_cells.sort_unstable();
        writer.count(skipped_cells.len())?;
        for address in skipped_cells {
            writer.address(address);
        }
        let mut rules = self.conditional_matches.iter().collect::<Vec<_>>();
        rules.sort_by_key(|(key, _)| **key);
        writer.count(rules.len())?;
        for (&(sheet, priority), cells) in rules {
            writer.u32(sheet);
            writer.u32(priority);
            let mut cells = cells.iter().copied().collect::<Vec<_>>();
            cells.sort_unstable();
            writer.count(cells.len())?;
            for (column, row) in cells {
                writer.u32(column);
                writer.u32(row);
            }
        }
        Ok(writer.0)
    }

    pub fn decode(bytes: &[u8], max_items: usize) -> Result<Self, String> {
        let mut reader = Reader::new(bytes)?;
        let mut results = Self::default();
        let value_count = reader.count(max_items)?;
        for _ in 0..value_count {
            let address = reader.address()?;
            if address.row < 1 || address.column < 1 {
                return Err("calculated cell address is invalid".to_owned());
            }
            let value = match reader.u8()? {
                0 => FormulaValue::Blank,
                1 => FormulaValue::Boolean(match reader.u8()? {
                    0 => false,
                    1 => true,
                    _ => return Err("invalid calculated boolean".to_owned()),
                }),
                2 => FormulaValue::Error(reader.string()?),
                3 => {
                    let value = f64::from_le_bytes(reader.take(8)?.try_into().unwrap());
                    if !value.is_finite() {
                        return Err("calculation result is not finite".to_owned());
                    }
                    FormulaValue::Number(value)
                }
                4 => FormulaValue::Text(reader.string()?),
                _ => return Err("invalid calculation result value".to_owned()),
            };
            if results.values.insert(address, value).is_some() {
                return Err("duplicate calculated cell".to_owned());
            }
        }
        let array_count = reader.count(max_items.saturating_sub(value_count))?;
        for _ in 0..array_count {
            results.array_cells.insert(reader.address()?);
        }
        let skipped_count = reader.count(
            max_items
                .saturating_sub(value_count)
                .saturating_sub(array_count),
        )?;
        for _ in 0..skipped_count {
            results.skipped_cells.insert(reader.address()?);
        }
        let mut remaining = max_items.saturating_sub(value_count + array_count + skipped_count);
        let rule_count = reader.count(remaining)?;
        remaining -= rule_count;
        for _ in 0..rule_count {
            let sheet = reader.u32()?;
            let priority = reader.u32()?;
            let count = reader.count(remaining)?;
            remaining -= count;
            if results.conditional_matches.contains_key(&(sheet, priority)) {
                return Err("duplicate conditional-format rule".to_owned());
            }
            results.insert_conditional_rule(sheet, priority);
            for _ in 0..count {
                let column = reader.u32()?;
                let row = reader.u32()?;
                if row >= 1_048_576 || column >= 16_384 {
                    return Err("conditional-format address is invalid".to_owned());
                }
                results.insert_conditional_match(
                    CellAddress {
                        sheet,
                        row: row as i32 + 1,
                        column: column as i32 + 1,
                    },
                    priority,
                );
            }
        }
        if reader.offset != bytes.len() {
            return Err("calculation payload has trailing bytes".to_owned());
        }
        Ok(results)
    }
}

struct Writer(Vec<u8>);

impl Writer {
    fn u8(&mut self, value: u8) {
        self.0.push(value);
    }

    fn u32(&mut self, value: u32) {
        self.0.extend_from_slice(&value.to_le_bytes());
    }

    fn count(&mut self, value: usize) -> Result<(), String> {
        self.u32(
            value
                .try_into()
                .map_err(|_| "calculation collection is too large")?,
        );
        Ok(())
    }

    fn string(&mut self, value: &str) -> Result<(), String> {
        self.count(value.len())?;
        self.0.extend_from_slice(value.as_bytes());
        Ok(())
    }

    fn address(&mut self, address: CellAddress) {
        self.u32(address.sheet);
        self.0.extend_from_slice(&address.row.to_le_bytes());
        self.0.extend_from_slice(&address.column.to_le_bytes());
    }
}

struct Reader<'a> {
    bytes: &'a [u8],
    offset: usize,
}

impl<'a> Reader<'a> {
    fn new(bytes: &'a [u8]) -> Result<Self, String> {
        if !bytes.starts_with(RESULT_MAGIC) {
            return Err("calculation payload has an invalid header".to_owned());
        }
        Ok(Self { bytes, offset: 8 })
    }

    fn take(&mut self, length: usize) -> Result<&'a [u8], String> {
        let end = self
            .offset
            .checked_add(length)
            .filter(|end| *end <= self.bytes.len())
            .ok_or_else(|| "calculation payload is truncated".to_owned())?;
        let value = &self.bytes[self.offset..end];
        self.offset = end;
        Ok(value)
    }

    fn u8(&mut self) -> Result<u8, String> {
        Ok(self.take(1)?[0])
    }

    fn u32(&mut self) -> Result<u32, String> {
        Ok(u32::from_le_bytes(self.take(4)?.try_into().unwrap()))
    }

    fn count(&mut self, maximum: usize) -> Result<usize, String> {
        let count = self.u32()? as usize;
        (count <= maximum)
            .then_some(count)
            .ok_or_else(|| "calculation collection exceeds the configured limit".to_owned())
    }

    fn string(&mut self) -> Result<String, String> {
        let length = self.u32()? as usize;
        std::str::from_utf8(self.take(length)?)
            .map(str::to_owned)
            .map_err(|_| "calculation string is not UTF-8".to_owned())
    }

    fn address(&mut self) -> Result<CellAddress, String> {
        Ok(CellAddress {
            sheet: self.u32()?,
            row: i32::from_le_bytes(self.take(4)?.try_into().unwrap()),
            column: i32::from_le_bytes(self.take(4)?.try_into().unwrap()),
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn calculation_results_round_trip() {
        let address = CellAddress {
            sheet: 0,
            row: 1,
            column: 1,
        };
        let mut results = CalculationResults::default();
        results.insert_value(address, FormulaValue::Number(2.4));
        results.insert_array_cell(address);
        results.insert_conditional_rule(0, 2);
        results.insert_conditional_match(address, 1);
        let decoded = CalculationResults::decode(&results.encode().unwrap(), 10).unwrap();
        assert_eq!(decoded.value(address), Ok(&FormulaValue::Number(2.4)));
        assert!(decoded.is_array_cell(address));
        assert_eq!(
            decoded.conditional_matches(0, 1),
            Some(&HashSet::from([(0, 0)]))
        );
        assert_eq!(decoded.conditional_matches(0, 2), Some(&HashSet::new()));
        assert!(CalculationResults::decode(&results.encode().unwrap(), 4).is_err());
    }
}
