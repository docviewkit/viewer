//! Read-only XLSX formula evaluation behind a narrow rendering seam.
#![allow(dead_code)] // Preview-only helpers remain covered while the calculator uses full workbooks.

use std::collections::{HashMap, HashSet};
use std::hash::{BuildHasher, Hasher, RandomState};

use ironcalc_base::{
    Model,
    cell::CellValue,
    expressions::{
        parser::{
            DefinedNameS, Node, Parser, new_parser_english,
            static_analysis::add_implicit_intersection, stringify::to_english_string,
        },
        types::{CellReferenceIndex, CellReferenceRC},
    },
    types::CellType,
};

use crate::{
    calculation::{
        CalculationResults, CellAddress, FormulaValue as ResultValue, formula_calls_any,
    },
    limits::Limits,
    package::Package,
    xml::{XmlEvent, decode_xml_text, parse_ooxml as parse_xml},
};

use super::local_name;

type CoordinateMap<V> = HashMap<(i32, i32), V, FastBuildHasher>;
type CellMap<V> = HashMap<CellReferenceIndex, V, FastBuildHasher>;
type CellSet = HashSet<CellReferenceIndex, FastBuildHasher>;
type ReferenceSet = HashSet<PreviewReference, FastBuildHasher>;
type SharedFormulaMap = HashMap<(u32, u32), Node, FastBuildHasher>;

#[derive(Clone)]
struct FastBuildHasher {
    seed: u64,
}

impl Default for FastBuildHasher {
    fn default() -> Self {
        let mut hasher = RandomState::new().build_hasher();
        hasher.write_u64(0x517c_c1b7_2722_0a95);
        Self {
            seed: hasher.finish(),
        }
    }
}

impl BuildHasher for FastBuildHasher {
    type Hasher = FastHasher;

    fn build_hasher(&self) -> Self::Hasher {
        FastHasher(self.seed)
    }
}

struct FastHasher(u64);

impl FastHasher {
    #[inline]
    fn mix(&mut self, value: u64) {
        self.0 ^= value.wrapping_mul(0x9e37_79b1_85eb_ca87);
        self.0 = self.0.rotate_left(27).wrapping_mul(0x94d0_49bb_1331_11eb);
    }
}

impl Hasher for FastHasher {
    fn finish(&self) -> u64 {
        let mut value = self.0;
        value ^= value >> 33;
        value = value.wrapping_mul(0xff51_afd7_ed55_8ccd);
        value ^= value >> 33;
        value
    }

    fn write(&mut self, bytes: &[u8]) {
        for chunk in bytes.chunks(8) {
            let mut value = [0_u8; 8];
            value[..chunk.len()].copy_from_slice(chunk);
            self.mix(u64::from_le_bytes(value));
        }
    }

    fn write_u32(&mut self, value: u32) {
        self.mix(u64::from(value));
    }

    fn write_i32(&mut self, value: i32) {
        self.mix(value as u32 as u64);
    }
}

#[derive(Clone, Debug)]
pub(super) struct FormulaSheet {
    pub(super) name: String,
    pub(super) part: Option<String>,
}

#[derive(Clone, Debug, PartialEq)]
pub(super) enum FormulaValue {
    Blank,
    Boolean(bool),
    Error(String),
    Number(f64),
    Text(String),
}

pub(super) struct FormulaEvaluator {
    model: Model<'static>,
    source_to_model_sheet: Vec<u32>,
    array_cells: HashSet<(u32, i32, i32)>,
    skipped_cells: HashSet<(u32, i32, i32)>,
    value_overrides: HashMap<(u32, i32, i32), FormulaValue>,
    root_scope: Option<(u32, FormulaRootBounds)>,
    formula_cells: HashSet<(u32, i32, i32)>,
}

#[derive(Clone, Copy, Debug)]
pub(super) struct FormulaRootBounds {
    pub(super) max_row: i32,
    pub(super) max_column: i32,
}

impl FormulaEvaluator {
    pub(super) fn for_sheet(
        package: &Package<'_>,
        workbook_part: &str,
        sheets: &[FormulaSheet],
        shared_strings: &[String],
        used_row_counts: Vec<u32>,
        target_sheet: u32,
        root_bounds: Option<FormulaRootBounds>,
    ) -> Result<Self, String> {
        let target_sheet_index = usize::try_from(target_sheet)
            .map_err(|_| "formula sheet exceeds the supported range")?;
        if sheets.get(target_sheet_index).is_none() {
            return Err("formula sheet is missing".to_owned());
        }
        let bounds = WorkbookBounds::new(sheets, used_row_counts);
        let defined_names = parse_defined_names(package, workbook_part)?;
        let parser_defined_names = defined_names
            .iter()
            .filter(|name| {
                !name.name.starts_with("_xlnm.") && !has_nondeterministic_function(&name.formula)
            })
            .map(|name| {
                (
                    name.name.clone(),
                    name.scope,
                    bound_whole_column_references(&name.formula, name.scope, &bounds),
                )
            })
            .collect::<Vec<DefinedNameS>>();
        let sheet_names = sheets
            .iter()
            .map(|sheet| sheet.name.clone())
            .collect::<Vec<_>>();
        let mut parser =
            new_parser_english(sheet_names.clone(), parser_defined_names, HashMap::new());
        let mut loaded_inputs = vec![None::<CoordinateMap<InputValue>>; sheets.len()];
        load_preview_inputs(
            package,
            sheets,
            shared_strings,
            target_sheet,
            &mut loaded_inputs,
        )?;
        let formulas_by_cell = loaded_inputs[target_sheet_index]
            .as_ref()
            .expect("target sheet inputs were loaded")
            .iter()
            .filter_map(|(&(row, column), value)| match value {
                InputValue::Formula { formula, .. } if !formula.is_empty() => {
                    Some(((target_sheet, row, column), formula.clone()))
                }
                InputValue::Literal(_) | InputValue::Formula { .. } => None,
            })
            .collect::<HashMap<_, _>>();
        let hyperlink_indirect_cells = loaded_inputs[target_sheet_index]
            .as_ref()
            .expect("target sheet inputs were loaded")
            .iter()
            .filter_map(|(&(row, column), value)| match value {
                InputValue::Formula { formula, .. } => {
                    hyperlink_indirect_cell(formula).map(|(reference_row, reference_column)| {
                        HyperlinkIndirectCell {
                            sheet: target_sheet,
                            row,
                            column,
                            reference_row,
                            reference_column,
                            match_address: formulas_by_cell
                                .get(&(target_sheet, reference_row, reference_column))
                                .and_then(|formula| parse_indirect_match_address(formula)),
                        }
                    })
                }
                InputValue::Literal(_) => None,
            })
            .collect::<Vec<_>>();
        let roots = loaded_inputs[target_sheet_index]
            .as_ref()
            .expect("target sheet inputs were loaded")
            .iter()
            .filter_map(|(&(row, column), value)| {
                (matches!(value, InputValue::Formula { .. })
                    && root_bounds
                        .is_none_or(|bounds| row <= bounds.max_row && column <= bounds.max_column))
                .then_some(CellReferenceIndex {
                    sheet: target_sheet,
                    row,
                    column,
                })
            })
            .collect::<Vec<_>>();
        let formula_cells = roots
            .iter()
            .map(|cell| (cell.sheet, cell.row, cell.column))
            .collect();
        let mut scheduled = roots.iter().copied().collect::<CellSet>();
        let mut pending = roots;
        let mut processed = CellSet::default();
        let mut expanded_ranges = ReferenceSet::default();
        let mut shared_master_nodes = SharedFormulaMap::default();
        let mut shared_prepared_nodes = SharedFormulaMap::default();
        let mut model_inputs = CellMap::<PreviewModelInput>::default();
        let mut array_cells = HashSet::new();
        let mut skipped_cells = HashSet::new();
        while let Some(cell) = pending.pop() {
            if !processed.insert(cell) {
                continue;
            }
            load_preview_inputs(
                package,
                sheets,
                shared_strings,
                cell.sheet,
                &mut loaded_inputs,
            )?;
            let Some(value) = loaded_inputs[cell.sheet as usize]
                .as_ref()
                .and_then(|inputs| inputs.get(&(cell.row, cell.column)))
                .cloned()
            else {
                continue;
            };
            match value {
                InputValue::Literal(value) => {
                    model_inputs.insert(cell, PreviewModelInput::Literal(value));
                }
                InputValue::Formula {
                    formula,
                    shared,
                    array_range,
                } => {
                    let shared_key = shared.map(|shared| (cell.sheet, shared));
                    let cached_node = formula.is_empty().then(|| {
                        shared_key
                            .and_then(|key| shared_prepared_nodes.get(&key))
                            .cloned()
                    });
                    let (node, is_hyperlink_indirect) = if let Some(node) = cached_node.flatten() {
                        (node, false)
                    } else {
                        let formula = preview_formula(
                            &mut parser,
                            &sheet_names,
                            &loaded_inputs[cell.sheet as usize],
                            cell,
                            formula,
                            shared,
                            &mut shared_master_nodes,
                        )?;
                        if formula.is_empty() {
                            continue;
                        }
                        if has_nondeterministic_function(&formula) {
                            skipped_cells.insert((cell.sheet, cell.row, cell.column));
                            continue;
                        }
                        let is_hyperlink_indirect = hyperlink_indirect_cell(&formula).is_some();
                        let display_formula =
                            rewrite_address_defaults(&rewrite_hyperlink_display(&formula));
                        let indirect_match = parse_indirect_match_address(&display_formula);
                        let static_address = indirect_match
                            .is_none()
                            .then(|| parse_static_address(&display_formula))
                            .flatten();
                        let dynamic_reference = has_dynamic_reference_function(&display_formula);
                        if dynamic_reference && !is_hyperlink_indirect && indirect_match.is_none() {
                            return Err(format!(
                                "sheet formula {} at row {}, column {} uses a dynamic cell reference: {}",
                                sheet_names[cell.sheet as usize],
                                cell.row,
                                cell.column,
                                display_formula
                            ));
                        }
                        if let Some(address) = indirect_match {
                            queue_indirect_address_inputs(
                                package,
                                sheets,
                                shared_strings,
                                &bounds,
                                &mut loaded_inputs,
                                (cell.sheet, address),
                                (&mut pending, &mut scheduled),
                            )?;
                        } else if let Some(address) = static_address {
                            queue_static_address_input(
                                package,
                                sheets,
                                shared_strings,
                                &mut loaded_inputs,
                                (cell.sheet, address),
                                (&mut pending, &mut scheduled),
                            )?;
                        }
                        let bounded = bound_whole_column_references(
                            &display_formula,
                            Some(cell.sheet),
                            &bounds,
                        );
                        let node = parse_cell_formula_node(
                            &mut parser,
                            &bounded,
                            &sheet_names[cell.sheet as usize],
                            cell.row,
                            cell.column,
                            array_range.is_some() || has_dynamic_array_function(&bounded),
                        )
                        .1;
                        if let Some(key) = shared_key
                            && !is_hyperlink_indirect
                            && !dynamic_reference
                            && indirect_match.is_none()
                            && static_address.is_none()
                        {
                            shared_prepared_nodes.insert(key, node.clone());
                        }
                        (node, is_hyperlink_indirect)
                    };
                    if matches!(node, Node::ParseErrorKind { .. }) {
                        return Err("preview formula could not be parsed".to_owned());
                    }
                    let context = CellReferenceRC {
                        sheet: sheet_names[cell.sheet as usize].clone(),
                        row: cell.row,
                        column: cell.column,
                    };
                    let mut references = Vec::new();
                    collect_preview_references(
                        &node,
                        &mut parser,
                        &context,
                        &sheet_names,
                        16,
                        &mut references,
                    )?;
                    for reference in references {
                        if reference.is_range() && !expanded_ranges.insert(reference) {
                            continue;
                        }
                        load_preview_inputs(
                            package,
                            sheets,
                            shared_strings,
                            reference.sheet,
                            &mut loaded_inputs,
                        )?;
                        let Some(inputs) = loaded_inputs[reference.sheet as usize].as_ref() else {
                            continue;
                        };
                        queue_preview_reference(inputs, reference, &mut pending, &mut scheduled);
                    }
                    if let Some(range) = array_range {
                        for row in range.start_row..=range.end_row {
                            for column in range.start_column..=range.end_column {
                                array_cells.insert((cell.sheet, row, column));
                            }
                        }
                    }
                    if !is_hyperlink_indirect {
                        model_inputs.insert(cell, PreviewModelInput::Formula(node));
                    }
                }
            }
        }

        let mut model = Model::new_empty("Workbook.xlsx", "en", "UTC", "en")?;
        model.rename_sheet_by_index(0, &sheet_names[0])?;
        for sheet_name in &sheet_names[1..] {
            model.add_sheet(sheet_name)?;
        }
        for defined_name in &defined_names {
            if defined_name.name.starts_with("_xlnm.")
                || has_nondeterministic_function(&defined_name.formula)
            {
                continue;
            }
            let formula =
                bound_whole_column_references(&defined_name.formula, defined_name.scope, &bounds);
            model.new_defined_name(&defined_name.name, defined_name.scope, &formula)?;
        }
        let mut sorted_inputs = model_inputs.into_iter().collect::<Vec<_>>();
        sorted_inputs.sort_by_key(|(cell, _)| (cell.sheet, cell.row, cell.column));
        for (cell, input) in sorted_inputs {
            match input {
                PreviewModelInput::Literal(value)
                    if !value.is_empty()
                        && !array_cells.contains(&(cell.sheet, cell.row, cell.column)) =>
                {
                    model.set_user_input(cell.sheet, cell.row, cell.column, value)?;
                }
                PreviewModelInput::Formula(node) => {
                    model.update_cell_with_parsed_formula(
                        cell.sheet,
                        cell.row,
                        cell.column,
                        node,
                    )?;
                }
                PreviewModelInput::Literal(_) => {}
            }
        }
        model.evaluate();
        let source_to_model_sheet = (0..sheets.len())
            .map(|index| u32::try_from(index).unwrap_or(u32::MAX))
            .collect::<Vec<_>>();
        let mut hyperlink_values = Vec::new();
        for hyperlink in &hyperlink_indirect_cells {
            let reference = if let Some(match_address) = hyperlink.match_address {
                resolve_indirect_match_address(
                    &model,
                    &source_to_model_sheet,
                    sheets,
                    &bounds,
                    hyperlink.sheet,
                    match_address,
                )?
            } else {
                match model.get_cell_value_by_index(
                    hyperlink.sheet,
                    hyperlink.reference_row,
                    hyperlink.reference_column,
                )? {
                    CellValue::String(reference) => Some(reference),
                    CellValue::None | CellValue::Boolean(_) | CellValue::Number(_) => None,
                }
            };
            let Some(reference) = reference else {
                continue;
            };
            let Some((sheet, row, column)) =
                parse_cross_sheet_reference(&reference, hyperlink.sheet, sheets)
            else {
                continue;
            };
            let value = model.get_cell_value_by_index(sheet, row, column)?;
            let cell_type = model.get_cell_type(sheet, row, column)?;
            hyperlink_values.push((
                hyperlink.sheet,
                hyperlink.row,
                hyperlink.column,
                cell_value_input(value, cell_type),
            ));
        }
        for (sheet, row, column, value) in hyperlink_values {
            model.set_user_input(sheet, row, column, value)?;
        }
        if !hyperlink_indirect_cells.is_empty() {
            model.evaluate();
        }
        Ok(Self {
            model,
            source_to_model_sheet,
            array_cells,
            skipped_cells,
            value_overrides: HashMap::new(),
            root_scope: root_bounds.map(|bounds| (target_sheet, bounds)),
            formula_cells,
        })
    }

    pub(super) fn from_package(
        package: &Package<'_>,
        workbook_part: &str,
        sheets: &[FormulaSheet],
        shared_strings: &[String],
    ) -> Result<Self, String> {
        if sheets.is_empty() {
            return Err("formula workbook has no sheets".to_owned());
        }
        let mut inputs = Vec::new();
        let mut used_row_counts = vec![1_u32; sheets.len()];
        for (sheet_index, sheet) in sheets.iter().enumerate() {
            let Some(part) = sheet.part.as_deref() else {
                continue;
            };
            let bytes = package
                .required_part(part)
                .map_err(|error| error.to_string())?;
            let sheet_inputs = parse_sheet_inputs(
                &bytes,
                package,
                u32::try_from(sheet_index).map_err(|_| "too many formula sheets")?,
                shared_strings,
                part,
            )?;
            used_row_counts[sheet_index] = sheet_inputs
                .iter()
                .map(|input| input.row as u32)
                .max()
                .unwrap_or(1);
            inputs.extend(sheet_inputs);
        }
        let bounds = WorkbookBounds::new(sheets, used_row_counts);
        let defined_names = parse_defined_names(package, workbook_part)?;
        let model_order = formula_sheet_order(sheets, &inputs, &defined_names);
        let mut source_to_model_sheet = vec![0_u32; sheets.len()];
        for (model_sheet, source_sheet) in model_order.iter().copied().enumerate() {
            source_to_model_sheet[source_sheet] =
                u32::try_from(model_sheet).map_err(|_| "too many formula sheets")?;
        }

        let first_sheet = &sheets[model_order[0]];
        let mut model = Model::new_empty("Workbook.xlsx", "en", "UTC", "en")?;
        model.rename_sheet_by_index(0, &first_sheet.name)?;
        for source_sheet in &model_order[1..] {
            model.add_sheet(&sheets[*source_sheet].name)?;
        }

        for defined_name in &defined_names {
            if defined_name.name.starts_with("_xlnm.")
                || has_nondeterministic_function(&defined_name.formula)
            {
                continue;
            }
            let formula =
                bound_whole_column_references(&defined_name.formula, defined_name.scope, &bounds);
            let scope = defined_name
                .scope
                .and_then(|scope| source_to_model_sheet.get(scope as usize).copied());
            model.new_defined_name(&defined_name.name, scope, &formula)?;
        }
        let parser_defined_names = defined_names
            .iter()
            .filter(|name| {
                !name.name.starts_with("_xlnm.") && !has_nondeterministic_function(&name.formula)
            })
            .map(|name| {
                (
                    name.name.clone(),
                    name.scope
                        .and_then(|scope| source_to_model_sheet.get(scope as usize).copied()),
                    bound_whole_column_references(&name.formula, name.scope, &bounds),
                )
            })
            .collect::<Vec<DefinedNameS>>();
        let model_sheet_names = model_order
            .iter()
            .map(|source| sheets[*source].name.clone())
            .collect::<Vec<_>>();
        let mut formula_parser = new_parser_english(
            model_sheet_names.clone(),
            parser_defined_names,
            HashMap::new(),
        );
        let array_cells = inputs
            .iter()
            .filter_map(|input| match &input.value {
                InputValue::Formula {
                    array_range: Some(range),
                    ..
                } => Some((input.sheet, *range)),
                InputValue::Literal(_)
                | InputValue::Formula {
                    array_range: None, ..
                } => None,
            })
            .flat_map(|(sheet, range)| {
                (range.start_row..=range.end_row).flat_map(move |row| {
                    (range.start_column..=range.end_column).map(move |column| (sheet, row, column))
                })
            })
            .collect::<HashSet<_>>();
        let formula_cells = inputs
            .iter()
            .filter(|input| matches!(input.value, InputValue::Formula { .. }))
            .map(|input| (input.sheet, input.row, input.column))
            .collect::<HashSet<_>>();
        let formulas_by_cell = inputs
            .iter()
            .filter_map(|input| match &input.value {
                InputValue::Formula { formula, .. } if !formula.is_empty() => {
                    Some(((input.sheet, input.row, input.column), formula.as_str()))
                }
                InputValue::Literal(_) | InputValue::Formula { .. } => None,
            })
            .collect::<HashMap<_, _>>();
        let hyperlink_indirect_cells = inputs
            .iter()
            .filter_map(|input| match &input.value {
                InputValue::Formula { formula, .. } => {
                    hyperlink_indirect_cell(formula).map(|(reference_row, reference_column)| {
                        HyperlinkIndirectCell {
                            sheet: input.sheet,
                            row: input.row,
                            column: input.column,
                            reference_row,
                            reference_column,
                            match_address: formulas_by_cell
                                .get(&(input.sheet, reference_row, reference_column))
                                .and_then(|formula| parse_indirect_match_address(formula)),
                        }
                    })
                }
                InputValue::Literal(_) => None,
            })
            .collect::<Vec<_>>();
        let hyperlink_cells = hyperlink_indirect_cells
            .iter()
            .map(|cell| CellReferenceIndex {
                sheet: source_to_model_sheet[cell.sheet as usize],
                row: cell.row,
                column: cell.column,
            })
            .collect::<HashSet<_>>();
        let hyperlink_helper_cells = hyperlink_indirect_cells
            .iter()
            .flat_map(|cell| {
                [
                    (cell.sheet, cell.row, cell.column),
                    (cell.sheet, cell.reference_row, cell.reference_column),
                ]
            })
            .collect::<HashSet<_>>();
        let hyperlink_sheets = hyperlink_indirect_cells
            .iter()
            .map(|cell| source_to_model_sheet[cell.sheet as usize])
            .collect::<HashSet<_>>();
        // These navigation formulas are display-only in many templates. Prove that no formula
        // reads them before using output overrides; otherwise preserve normal recalculation.
        let mut hyperlinks_have_dependents = false;
        let mut shared_formulas = HashMap::<(u32, u32), SharedFormula>::new();
        let mut skipped_shared_formulas = HashSet::<(u32, u32)>::new();
        let mut skipped_cells = HashSet::<(u32, i32, i32)>::new();
        let mut followers = Vec::new();
        for input in inputs {
            let model_sheet = source_to_model_sheet[input.sheet as usize];
            match input.value {
                InputValue::Literal(value) => {
                    if !array_cells.contains(&(input.sheet, input.row, input.column))
                        && !value.is_empty()
                    {
                        model.set_user_input(model_sheet, input.row, input.column, value)?;
                    }
                }
                InputValue::Formula {
                    formula,
                    shared,
                    array_range,
                } if !formula.is_empty() => {
                    if has_nondeterministic_function(&formula) {
                        skipped_cells.insert((input.sheet, input.row, input.column));
                        if let Some(shared) = shared {
                            skipped_shared_formulas.insert((model_sheet, shared));
                        }
                        continue;
                    }
                    let display_formula =
                        rewrite_address_defaults(&rewrite_hyperlink_display(&formula));
                    let bounded =
                        bound_whole_column_references(&display_formula, Some(input.sheet), &bounds);
                    let (formula, node) = import_cell_formula_and_node(
                        &mut formula_parser,
                        &bounded,
                        &sheets[input.sheet as usize].name,
                        input.row,
                        input.column,
                        array_range.is_some() || has_dynamic_array_function(&bounded),
                        &model_sheet_names,
                    );
                    if !hyperlink_helper_cells.contains(&(input.sheet, input.row, input.column))
                        && (has_dynamic_reference_function(&bounded)
                            || node_may_reference_cells(
                                &node,
                                &mut formula_parser,
                                &CellReferenceRC {
                                    sheet: sheets[input.sheet as usize].name.clone(),
                                    row: input.row,
                                    column: input.column,
                                },
                                &hyperlink_cells,
                                &hyperlink_sheets,
                                &model_sheet_names,
                                16,
                            )
                            || shared.is_some()
                                && node_may_reference_sheet(&node, &hyperlink_sheets))
                    {
                        hyperlinks_have_dependents = true;
                    }
                    if let Some(range) = array_range {
                        if input.row != range.start_row || input.column != range.start_column {
                            return Err(format!(
                                "array formula anchor does not match its range on sheet {}",
                                input.sheet
                            ));
                        }
                        range.width()?;
                        range.height()?;
                        model.update_cell_with_formula(
                            model_sheet,
                            input.row,
                            input.column,
                            formula,
                        )?;
                    } else {
                        model.update_cell_with_formula(
                            model_sheet,
                            input.row,
                            input.column,
                            formula,
                        )?;
                    }
                    if let Some(shared) = shared {
                        shared_formulas.insert((model_sheet, shared), SharedFormula { node });
                    }
                }
                InputValue::Formula {
                    shared: Some(shared),
                    ..
                } => {
                    followers.push((input.sheet, model_sheet, input.row, input.column, shared));
                }
                InputValue::Formula { shared: None, .. } => {}
            }
        }
        for (source_sheet, sheet, row, column, shared) in followers {
            if skipped_shared_formulas.contains(&(sheet, shared)) {
                skipped_cells.insert((source_sheet, row, column));
                continue;
            }
            let master = shared_formulas
                .get(&(sheet, shared))
                .ok_or_else(|| format!("shared formula {shared} has no master on sheet {sheet}"))?;
            let context = CellReferenceRC {
                sheet: model_sheet_names[sheet as usize].clone(),
                row,
                column,
            };
            let formula = format!(
                "={}",
                restore_sheet_name_quotes(
                    &to_english_string(&master.node, &context),
                    &model_sheet_names,
                )
            );
            model.update_cell_with_formula(sheet, row, column, formula)?;
        }
        model.evaluate();
        let mut hyperlink_references = Vec::new();
        for hyperlink in &hyperlink_indirect_cells {
            let model_sheet = source_to_model_sheet[hyperlink.sheet as usize];
            let reference = if let Some(match_address) = hyperlink.match_address {
                resolve_indirect_match_address(
                    &model,
                    &source_to_model_sheet,
                    sheets,
                    &bounds,
                    hyperlink.sheet,
                    match_address,
                )?
            } else {
                match model.get_cell_value_by_index(
                    model_sheet,
                    hyperlink.reference_row,
                    hyperlink.reference_column,
                )? {
                    CellValue::String(reference) => Some(reference),
                    CellValue::None | CellValue::Boolean(_) | CellValue::Number(_) => None,
                }
            };
            if let Some(reference) = reference
                && parse_cross_sheet_reference(&reference, hyperlink.sheet, sheets).is_some()
            {
                hyperlink_references.push((*hyperlink, reference));
            }
        }
        let mut value_overrides = HashMap::new();
        let mut hyperlink_values = Vec::new();
        for (hyperlink, reference) in &hyperlink_references {
            let model_sheet = source_to_model_sheet[hyperlink.sheet as usize];
            let Some((target_source_sheet, target_row, target_column)) =
                parse_cross_sheet_reference(reference, hyperlink.sheet, sheets)
            else {
                continue;
            };
            let target_model_sheet = source_to_model_sheet[target_source_sheet as usize];
            let value =
                model.get_cell_value_by_index(target_model_sheet, target_row, target_column)?;
            let cell_type = model.get_cell_type(target_model_sheet, target_row, target_column)?;
            if hyperlinks_have_dependents {
                hyperlink_values.push((
                    model_sheet,
                    hyperlink.row,
                    hyperlink.column,
                    cell_value_input(value, cell_type),
                ));
            } else {
                // Avoid a second whole-workbook evaluation solely to materialize hyperlink labels.
                value_overrides.insert(
                    (hyperlink.sheet, hyperlink.row, hyperlink.column),
                    formula_value(value, cell_type),
                );
            }
        }
        for (sheet, row, column, value) in hyperlink_values {
            model.set_user_input(sheet, row, column, value)?;
        }
        if hyperlinks_have_dependents && !hyperlink_references.is_empty() {
            model.evaluate();
        }
        Ok(Self {
            model,
            source_to_model_sheet,
            array_cells,
            skipped_cells,
            value_overrides,
            root_scope: None,
            formula_cells,
        })
    }

    pub(super) fn is_array_cell(
        &self,
        sheet: u32,
        zero_based_row: u32,
        zero_based_column: u32,
    ) -> bool {
        let Ok(row) = to_one_based_i32(zero_based_row, "formula row") else {
            return false;
        };
        let Ok(column) = to_one_based_i32(zero_based_column, "formula column") else {
            return false;
        };
        self.array_cells.contains(&(sheet, row, column))
    }

    pub(super) fn value(
        &self,
        sheet: u32,
        zero_based_row: u32,
        zero_based_column: u32,
    ) -> Result<FormulaValue, String> {
        let row = to_one_based_i32(zero_based_row, "formula row")?;
        let column = to_one_based_i32(zero_based_column, "formula column")?;
        if let Some(value) = self.value_overrides.get(&(sheet, row, column)) {
            return Ok(value.clone());
        }
        if self.skipped_cells.contains(&(sheet, row, column)) {
            return Err("formula result depends on time or randomness".to_owned());
        }
        let sheet = self
            .source_to_model_sheet
            .get(sheet as usize)
            .copied()
            .ok_or_else(|| "formula sheet exceeds the evaluator range".to_owned())?;
        let cell_type = self.model.get_cell_type(sheet, row, column)?;
        self.model
            .get_cell_value_by_index(sheet, row, column)
            .map(|value| formula_value(value, cell_type))
    }

    pub(super) fn is_deferred(
        &self,
        sheet: u32,
        zero_based_row: u32,
        zero_based_column: u32,
    ) -> bool {
        let Some((target_sheet, bounds)) = self.root_scope else {
            return false;
        };
        if sheet != target_sheet {
            return false;
        }
        zero_based_row
            .checked_add(1)
            .is_some_and(|row| row > bounds.max_row as u32)
            || zero_based_column
                .checked_add(1)
                .is_some_and(|column| column > bounds.max_column as u32)
    }

    fn into_results(self) -> CalculationResults {
        let mut results = CalculationResults::default();
        for &(sheet, row, column) in self.array_cells.iter() {
            results.insert_array_cell(CellAddress { sheet, row, column });
        }
        for &(sheet, row, column) in self.skipped_cells.iter() {
            results.insert_skipped_cell(CellAddress { sheet, row, column });
        }
        for &(sheet, row, column) in self.formula_cells.union(&self.array_cells) {
            let address = CellAddress { sheet, row, column };
            let Ok(source_row) = u32::try_from(row - 1) else {
                continue;
            };
            let Ok(source_column) = u32::try_from(column - 1) else {
                continue;
            };
            let Ok(value) = self.value(sheet, source_row, source_column) else {
                continue;
            };
            results.insert_value(
                address,
                match value {
                    FormulaValue::Blank => ResultValue::Blank,
                    FormulaValue::Boolean(value) => ResultValue::Boolean(value),
                    FormulaValue::Error(value) => ResultValue::Error(value),
                    FormulaValue::Number(value) => ResultValue::Number(value),
                    FormulaValue::Text(value) => ResultValue::Text(value),
                },
            );
        }
        results
    }
}

pub(crate) fn calculate(bytes: &[u8], limits: Limits) -> Result<CalculationResults, String> {
    let package = Package::open(bytes, limits).map_err(|error| error.to_string())?;
    let workbook_part = package
        .relationships(None)
        .map_err(|error| error.to_string())?
        .into_iter()
        .find(|relationship| {
            relationship.type_uri.ends_with("/officeDocument") && !relationship.external
        })
        .map(|relationship| relationship.target)
        .ok_or_else(|| "calculation input has no workbook".to_owned())?;
    let relationships = package
        .relationships(Some(&workbook_part))
        .map_err(|error| error.to_string())?;
    let workbook_bytes = package
        .required_part(&workbook_part)
        .map_err(|error| error.to_string())?;
    let mut raw_sheets = Vec::<(String, String)>::new();
    let mut date_1904 = false;
    parse_xml(&workbook_bytes, limits, |event| {
        if let XmlEvent::StartElement {
            name, attributes, ..
        } = &event
            && local_name(name) == "workbookPr"
        {
            date_1904 = matches!(attribute(attributes, "date1904"), Some("1" | "true"));
        }
        if let XmlEvent::StartElement {
            name, attributes, ..
        } = event
            && local_name(name) == "sheet"
        {
            let name = attribute(&attributes, "name")
                .ok_or_else(|| formula_error(&workbook_part, "formula sheet has no name"))?;
            let relationship = attributes
                .iter()
                .find(|attribute| local_name(attribute.name) == "id")
                .map(|attribute| attribute.value)
                .ok_or_else(|| {
                    formula_error(&workbook_part, "formula sheet has no relationship")
                })?;
            raw_sheets.push((name.to_owned(), relationship.to_owned()));
        }
        Ok(())
    })
    .map_err(|error| error.to_string())?;
    let sheets = raw_sheets
        .into_iter()
        .map(|(name, id)| FormulaSheet {
            name,
            part: relationships
                .iter()
                .find(|relationship| relationship.id == id && !relationship.external)
                .map(|relationship| relationship.target.clone()),
        })
        .collect::<Vec<_>>();
    let shared_strings = match relationships.iter().find(|relationship| {
        relationship.type_uri.ends_with("/sharedStrings") && !relationship.external
    }) {
        Some(relationship) => parse_formula_shared_strings(&package, &relationship.target)?,
        None => Vec::new(),
    };
    let mut evaluator =
        FormulaEvaluator::from_package(&package, &workbook_part, &sheets, &shared_strings)?;
    let evaluated_rules =
        calculate_conditional_rules(&package, &sheets, &mut evaluator, date_1904)?;
    let matched = evaluator
        .model
        .conditional_dxf_matches()
        .map(|((sheet, row, column), priority)| {
            let source = evaluator
                .source_to_model_sheet
                .iter()
                .position(|index| *index == sheet)
                .unwrap() as u32;
            (
                CellAddress {
                    sheet: source,
                    row,
                    column,
                },
                priority,
            )
        })
        .collect::<Vec<_>>();
    let mut results = evaluator.into_results();
    for (sheet, priority) in evaluated_rules {
        results.insert_conditional_rule(sheet, priority);
    }
    for (address, priority) in matched {
        results.insert_conditional_match(address, priority);
    }
    Ok(results)
}

fn calculate_conditional_rules(
    package: &Package<'_>,
    sheets: &[FormulaSheet],
    evaluator: &mut FormulaEvaluator,
    date_1904: bool,
) -> Result<Vec<(u32, u32)>, String> {
    use ironcalc_base::cf_types::{
        CfRule, ConditionalFormatting, PeriodType, TextOperator, ValueOperator,
    };
    use ironcalc_base::expressions::utils::parse_reference_a1;
    // Reserve the existing formula payload first so optional formatting cannot
    // make an otherwise usable calculation result exceed the receiver's budget.
    let mut budget = package
        .limits()
        .max_document_objects
        .saturating_sub(evaluator.formula_cells.len())
        .saturating_sub(evaluator.array_cells.len().saturating_mul(2))
        .saturating_sub(evaluator.skipped_cells.len());
    let mut evaluated_rules = Vec::new();
    for (source_sheet, sheet) in sheets.iter().enumerate() {
        let Some(part) = &sheet.part else { continue };
        let bytes = package.required_part(part).map_err(|e| e.to_string())?;
        let sheet_index = evaluator.source_to_model_sheet[source_sheet];
        let mut depth = 0;
        let mut range = String::new();
        let mut rule = None::<(HashMap<String, String>, Vec<String>)>;
        let mut formula = None::<String>;
        let mut parsed = Vec::new();
        parse_xml(&bytes, package.limits(), |event| {
            match event {
                XmlEvent::StartElement {
                    name,
                    attributes,
                    empty,
                } => {
                    match local_name(name) {
                        "conditionalFormatting" if depth == 1 => {
                            range = attribute(&attributes, "sqref")
                                .unwrap_or_default()
                                .to_owned();
                        }
                        "cfRule" if depth == 2 && !range.is_empty() => {
                            let attributes = attributes
                                .iter()
                                .map(|a| {
                                    Ok((a.name.to_owned(), decode_xml_text(a.value)?.into_owned()))
                                })
                                .collect::<Result<HashMap<_, _>, crate::diagnostic::Diagnostic>>(
                                )?;
                            if empty {
                                parsed.push((range.clone(), attributes, Vec::new()));
                            } else {
                                rule = Some((attributes, Vec::new()));
                            }
                        }
                        "formula" if rule.is_some() => {
                            formula = Some(String::new());
                        }
                        _ => {}
                    }
                    depth += usize::from(!empty);
                }
                XmlEvent::EndElement { name } => {
                    depth = depth.saturating_sub(1);
                    match local_name(name) {
                        "formula" => {
                            if let (Some((_, formulas)), Some(value)) =
                                (rule.as_mut(), formula.take())
                            {
                                formulas.push(value);
                            }
                        }
                        "cfRule" if depth == 2 => {
                            if let Some((attributes, formulas)) = rule.take() {
                                parsed.push((range.clone(), attributes, formulas));
                            }
                        }
                        "conditionalFormatting" if depth == 1 => {
                            range.clear();
                        }
                        _ => {}
                    }
                }
                XmlEvent::Text(text) | XmlEvent::Cdata(text) => {
                    if let Some(value) = formula.as_mut() {
                        value.push_str(&decode_xml_text(text)?);
                    }
                }
            }
            Ok(())
        })
        .map_err(|e| e.to_string())?;
        for (range, attrs, formulas) in parsed {
            let get = |key: &str| attrs.get(key).map(String::as_str).unwrap_or("");
            let kind = get("type");
            if !crate::calculation::calculated_condition(kind) {
                continue;
            }
            let priority = get("priority").parse::<u32>().unwrap_or(u32::MAX);
            // Evaluate the original areas, preserving their shared formula anchor.
            // A hard cell budget prevents whole-column rules from unbounded work.
            let count = range.split_whitespace().try_fold(1usize, |count, area| {
                let (first, last) = area.split_once(':').unwrap_or((area, area));
                let a = parse_reference_a1(first)?;
                let b = parse_reference_a1(last)?;
                if a.row < 1
                    || a.column < 1
                    || b.row > 1_048_576
                    || b.column > 16_384
                    || a.row > b.row
                    || a.column > b.column
                {
                    return None;
                }
                count.checked_add(
                    ((b.row - a.row + 1) as usize)
                        .checked_mul((b.column - a.column + 1) as usize)?,
                )
            });
            let Some(remaining) = count.and_then(|count| budget.checked_sub(count)) else {
                continue;
            };
            if formulas.iter().any(|f| has_nondeterministic_function(f)) {
                continue;
            }
            let first = formulas.first().cloned().unwrap_or_default();
            let dxf_id = priority; // Native renderer joins matches to the author's rule priority.
            let stop_if_true = false; // Preserve every match; host merges all rule kinds together.
            let condition = match kind {
                "expression" if !first.is_empty() => CfRule::Formula {
                    formula: first,
                    dxf_id,
                    stop_if_true,
                },
                "cellIs" if !first.is_empty() => {
                    let operator = match get("operator") {
                        "equal" => ValueOperator::Equal,
                        "notEqual" => ValueOperator::NotEqual,
                        "lessThan" => ValueOperator::LessThan,
                        "lessThanOrEqual" => ValueOperator::LessThanOrEqual,
                        "greaterThan" => ValueOperator::GreaterThan,
                        "greaterThanOrEqual" => ValueOperator::GreaterThanOrEqual,
                        "between" => ValueOperator::Between,
                        "notBetween" => ValueOperator::NotBetween,
                        _ => continue,
                    };
                    CfRule::CellIs {
                        operator,
                        formula: first,
                        formula2: formulas.get(1).cloned(),
                        dxf_id,
                        stop_if_true,
                    }
                }
                "containsText" | "notContainsText" | "beginsWith" | "endsWith"
                    if !first.is_empty() =>
                {
                    CfRule::Formula {
                        formula: first,
                        dxf_id,
                        stop_if_true,
                    }
                }
                "containsText" | "notContainsText" | "beginsWith" | "endsWith" => CfRule::Text {
                    operator: match kind {
                        "notContainsText" => TextOperator::DoesNotContain,
                        "beginsWith" => TextOperator::BeginsWith,
                        "endsWith" => TextOperator::EndsWith,
                        _ => TextOperator::Contains,
                    },
                    value: get("text").to_owned(),
                    dxf_id,
                    stop_if_true,
                },
                "containsBlanks" => CfRule::Blanks {
                    dxf_id,
                    stop_if_true,
                },
                "notContainsBlanks" => CfRule::NotBlanks {
                    dxf_id,
                    stop_if_true,
                },
                "containsErrors" => CfRule::Errors {
                    dxf_id,
                    stop_if_true,
                },
                "notContainsErrors" => CfRule::NoErrors {
                    dxf_id,
                    stop_if_true,
                },
                "duplicateValues" => CfRule::DuplicateValues {
                    dxf_id,
                    stop_if_true,
                },
                "uniqueValues" => CfRule::UniqueValues {
                    dxf_id,
                    stop_if_true,
                },
                "top10" => {
                    let rank = get("rank").parse::<u32>().unwrap_or(10);
                    let percent = matches!(get("percent"), "1" | "true");
                    if matches!(get("bottom"), "1" | "true") {
                        CfRule::Bottom10 {
                            rank,
                            percent,
                            dxf_id,
                            stop_if_true,
                        }
                    } else {
                        CfRule::Top10 {
                            rank,
                            percent,
                            dxf_id,
                            stop_if_true,
                        }
                    }
                }
                "aboveAverage"
                    if !matches!(get("equalAverage"), "1" | "true")
                        && matches!(get("stdDev"), "" | "0") =>
                {
                    if matches!(get("aboveAverage"), "0" | "false") {
                        CfRule::BelowAverage {
                            dxf_id,
                            stop_if_true,
                        }
                    } else {
                        CfRule::AboveAverage {
                            dxf_id,
                            stop_if_true,
                        }
                    }
                }
                "timePeriod" if date_1904 => continue,
                "timePeriod" if !first.is_empty() => CfRule::Formula {
                    formula: first,
                    dxf_id,
                    stop_if_true,
                },
                "timePeriod" => {
                    let time_period = match get("timePeriod") {
                        "yesterday" => PeriodType::Yesterday,
                        "today" => PeriodType::Today,
                        "tomorrow" => PeriodType::Tomorrow,
                        "last7Days" => PeriodType::Last7Days,
                        // Engine fallback uses Monday weeks; OOXML authors normally supply
                        // a WEEKDAY formula. Do not silently substitute another week boundary.
                        "lastWeek" | "thisWeek" | "nextWeek" => continue,
                        "lastMonth" => PeriodType::LastMonth,
                        "thisMonth" => PeriodType::ThisMonth,
                        "nextMonth" => PeriodType::NextMonth,
                        _ => continue,
                    };
                    CfRule::TimePeriod {
                        time_period,
                        date1: None,
                        date2: None,
                        dxf_id,
                        stop_if_true,
                    }
                }
                _ => continue,
            };
            budget = remaining;
            evaluated_rules.push((source_sheet as u32, priority));
            evaluator.model.workbook.worksheets[sheet_index as usize]
                .conditional_formatting
                .push(ConditionalFormatting {
                    range,
                    cf_rule: condition,
                    priority,
                });
        }
    }
    evaluator.model.evaluate_conditional_formatting();
    Ok(evaluated_rules)
}

fn parse_formula_shared_strings(package: &Package<'_>, part: &str) -> Result<Vec<String>, String> {
    let bytes = package
        .required_part(part)
        .map_err(|error| error.to_string())?;
    let mut depth = 0_usize;
    let mut item_depth = None;
    let mut text_depth = None;
    let mut value = String::new();
    let mut values = Vec::new();
    parse_xml(&bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement { name, empty, .. } => {
                match local_name(name) {
                    "si" => {
                        item_depth = (!empty).then_some(depth);
                        value.clear();
                        if empty {
                            values.push(String::new());
                        }
                    }
                    "t" if item_depth.is_some() => text_depth = (!empty).then_some(depth),
                    _ => {}
                }
                depth = depth.saturating_add(usize::from(!empty));
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                if local_name(name) == "t" {
                    text_depth = None;
                }
                if local_name(name) == "si" && item_depth == Some(depth) {
                    values.push(std::mem::take(&mut value));
                    item_depth = None;
                }
            }
            XmlEvent::Text(text) | XmlEvent::Cdata(text) if text_depth.is_some() => {
                value.push_str(&decode_xml_text(text)?);
            }
            XmlEvent::Text(_) | XmlEvent::Cdata(_) => {}
        }
        Ok(())
    })
    .map_err(|error| error.to_string())?;
    Ok(values)
}

#[cfg(test)]
fn import_cell_formula(
    parser: &mut Parser<'_>,
    formula: &str,
    sheet_name: &str,
    row: i32,
    column: i32,
    array_formula: bool,
    sheet_names: &[String],
) -> String {
    import_cell_formula_and_node(
        parser,
        formula,
        sheet_name,
        row,
        column,
        array_formula,
        sheet_names,
    )
    .0
}

fn import_cell_formula_and_node(
    parser: &mut Parser<'_>,
    formula: &str,
    sheet_name: &str,
    row: i32,
    column: i32,
    array_formula: bool,
    sheet_names: &[String],
) -> (String, Node) {
    let (context, node) =
        parse_cell_formula_node(parser, formula, sheet_name, row, column, array_formula);
    (
        format!(
            "={}",
            restore_sheet_name_quotes(&to_english_string(&node, &context), sheet_names)
        ),
        node,
    )
}

fn parse_cell_formula_node<'a>(
    parser: &mut Parser<'a>,
    formula: &str,
    sheet_name: &str,
    row: i32,
    column: i32,
    array_formula: bool,
) -> (CellReferenceRC, Node) {
    let context = CellReferenceRC {
        sheet: sheet_name.to_owned(),
        row,
        column,
    };
    let mut node = parser.parse(formula, &context);
    if !array_formula {
        add_implicit_intersection(&mut node, true);
    }
    (context, node)
}

fn node_may_reference_cells(
    node: &Node,
    parser: &mut Parser<'_>,
    context: &CellReferenceRC,
    targets: &HashSet<CellReferenceIndex>,
    target_sheets: &HashSet<u32>,
    sheet_names: &[String],
    remaining_depth: u8,
) -> bool {
    match node {
        Node::ReferenceKind {
            sheet_index,
            absolute_row,
            absolute_column,
            row,
            column,
            ..
        } => targets.contains(&CellReferenceIndex {
            sheet: *sheet_index,
            row: if *absolute_row {
                *row
            } else {
                context.row + *row
            },
            column: if *absolute_column {
                *column
            } else {
                context.column + *column
            },
        }),
        Node::RangeKind {
            sheet_index,
            absolute_row1,
            absolute_column1,
            row1,
            column1,
            absolute_row2,
            absolute_column2,
            row2,
            column2,
            ..
        } => {
            let row1 = if *absolute_row1 {
                *row1
            } else {
                context.row + *row1
            };
            let row2 = if *absolute_row2 {
                *row2
            } else {
                context.row + *row2
            };
            let column1 = if *absolute_column1 {
                *column1
            } else {
                context.column + *column1
            };
            let column2 = if *absolute_column2 {
                *column2
            } else {
                context.column + *column2
            };
            targets.iter().any(|target| {
                target.sheet == *sheet_index
                    && target.row >= row1.min(row2)
                    && target.row <= row1.max(row2)
                    && target.column >= column1.min(column2)
                    && target.column <= column1.max(column2)
            })
        }
        Node::DefinedNameKind((_, scope, formula)) if remaining_depth > 0 => {
            let sheet = scope
                .and_then(|sheet| sheet_names.get(sheet as usize))
                .cloned()
                .unwrap_or_else(|| context.sheet.clone());
            let defined_context = CellReferenceRC {
                sheet,
                row: context.row,
                column: context.column,
            };
            has_dynamic_reference_function(formula) || {
                let defined_node = parser.parse(formula, &defined_context);
                node_may_reference_cells(
                    &defined_node,
                    parser,
                    &defined_context,
                    targets,
                    target_sheets,
                    sheet_names,
                    remaining_depth - 1,
                )
            }
        }
        Node::OpRangeKind { left, right }
        | Node::OpConcatenateKind { left, right }
        | Node::OpSumKind { left, right, .. }
        | Node::OpProductKind { left, right, .. }
        | Node::OpPowerKind { left, right }
        | Node::CompareKind { left, right, .. } => {
            node_may_reference_cells(
                left,
                parser,
                context,
                targets,
                target_sheets,
                sheet_names,
                remaining_depth,
            ) || node_may_reference_cells(
                right,
                parser,
                context,
                targets,
                target_sheets,
                sheet_names,
                remaining_depth,
            )
        }
        Node::FunctionKind { args, .. } | Node::NamedFunctionKind { args, .. } => {
            args.iter().any(|argument| {
                node_may_reference_cells(
                    argument,
                    parser,
                    context,
                    targets,
                    target_sheets,
                    sheet_names,
                    remaining_depth,
                )
            })
        }
        Node::LambdaDefKind { body, .. } => node_may_reference_cells(
            body,
            parser,
            context,
            targets,
            target_sheets,
            sheet_names,
            remaining_depth,
        ),
        Node::LambdaCallKind { lambda, args } => {
            node_may_reference_cells(
                lambda,
                parser,
                context,
                targets,
                target_sheets,
                sheet_names,
                remaining_depth,
            ) || args.iter().any(|argument| {
                node_may_reference_cells(
                    argument,
                    parser,
                    context,
                    targets,
                    target_sheets,
                    sheet_names,
                    remaining_depth,
                )
            })
        }
        Node::ImplicitIntersection { child, .. } | Node::SpillRangeOperator { child } => {
            node_may_reference_cells(
                child,
                parser,
                context,
                targets,
                target_sheets,
                sheet_names,
                remaining_depth,
            )
        }
        Node::UnaryKind { right, .. } => node_may_reference_cells(
            right,
            parser,
            context,
            targets,
            target_sheets,
            sheet_names,
            remaining_depth,
        ),
        Node::WrongReferenceKind { sheet_name, .. } | Node::WrongRangeKind { sheet_name, .. } => {
            sheet_name.as_ref().is_none_or(|name| {
                sheet_names
                    .iter()
                    .position(|sheet| sheet == name)
                    .is_none_or(|sheet| target_sheets.contains(&(sheet as u32)))
            })
        }
        Node::ParseErrorKind { .. } | Node::TableNameKind(_) if !targets.is_empty() => true,
        Node::BooleanKind(_)
        | Node::NumberKind(_)
        | Node::StringKind(_)
        | Node::ArrayKind(_)
        | Node::DefinedNameKind(_)
        | Node::NamedVariableKind { .. }
        | Node::ErrorKind(_)
        | Node::EmptyArgKind
        | Node::ParseErrorKind { .. }
        | Node::TableNameKind(_) => false,
    }
}

fn node_may_reference_sheet(node: &Node, target_sheets: &HashSet<u32>) -> bool {
    match node {
        Node::ReferenceKind { sheet_index, .. } | Node::RangeKind { sheet_index, .. } => {
            target_sheets.contains(sheet_index)
        }
        Node::OpRangeKind { left, right }
        | Node::OpConcatenateKind { left, right }
        | Node::OpSumKind { left, right, .. }
        | Node::OpProductKind { left, right, .. }
        | Node::OpPowerKind { left, right }
        | Node::CompareKind { left, right, .. } => {
            node_may_reference_sheet(left, target_sheets)
                || node_may_reference_sheet(right, target_sheets)
        }
        Node::FunctionKind { args, .. } | Node::NamedFunctionKind { args, .. } => args
            .iter()
            .any(|argument| node_may_reference_sheet(argument, target_sheets)),
        Node::LambdaDefKind { body, .. } => node_may_reference_sheet(body, target_sheets),
        Node::LambdaCallKind { lambda, args } => {
            node_may_reference_sheet(lambda, target_sheets)
                || args
                    .iter()
                    .any(|argument| node_may_reference_sheet(argument, target_sheets))
        }
        Node::ImplicitIntersection { child, .. } | Node::SpillRangeOperator { child } => {
            node_may_reference_sheet(child, target_sheets)
        }
        Node::UnaryKind { right, .. } => node_may_reference_sheet(right, target_sheets),
        Node::BooleanKind(_)
        | Node::NumberKind(_)
        | Node::StringKind(_)
        | Node::WrongReferenceKind { .. }
        | Node::WrongRangeKind { .. }
        | Node::ArrayKind(_)
        | Node::DefinedNameKind(_)
        | Node::TableNameKind(_)
        | Node::NamedVariableKind { .. }
        | Node::ErrorKind(_)
        | Node::ParseErrorKind { .. }
        | Node::EmptyArgKind => false,
    }
}

fn has_dynamic_reference_function(formula: &str) -> bool {
    formula_contains_identifier(formula, "INDIRECT")
        || formula_contains_identifier(formula, "OFFSET")
}

fn hyperlink_indirect_cell(formula: &str) -> Option<(i32, i32)> {
    let formula = formula.trim();
    const HYPERLINK: &str = "HYPERLINK";
    if formula.len() <= HYPERLINK.len()
        || !formula[..HYPERLINK.len()].eq_ignore_ascii_case(HYPERLINK)
        || formula.as_bytes().get(HYPERLINK.len()) != Some(&b'(')
        || matching_parenthesis(formula, HYPERLINK.len()) != Some(formula.len() - 1)
    {
        return None;
    }
    let arguments = split_formula_arguments(&formula[HYPERLINK.len() + 1..formula.len() - 1])?;
    let display = arguments.get(1)?.trim();
    const INDIRECT: &str = "INDIRECT";
    if display.len() <= INDIRECT.len()
        || !display[..INDIRECT.len()].eq_ignore_ascii_case(INDIRECT)
        || display.as_bytes().get(INDIRECT.len()) != Some(&b'(')
        || matching_parenthesis(display, INDIRECT.len()) != Some(display.len() - 1)
    {
        return None;
    }
    let reference = display[INDIRECT.len() + 1..display.len() - 1]
        .trim()
        .replace('$', "");
    parse_a1(&reference).map(|(column, row)| (row, column))
}

fn parse_indirect_match_address(formula: &str) -> Option<IndirectMatchAddress> {
    let address = exact_formula_call_arguments(formula, "ADDRESS")?;
    if address.len() != 5 {
        return None;
    }
    let matches = exact_formula_call_arguments(address[0], "MATCH")?;
    if !(2..=3).contains(&matches.len())
        || matches
            .get(2)
            .is_some_and(|value| !value.trim().is_empty() && value.trim() != "0")
    {
        return None;
    }
    let (lookup_column, lookup_row) = parse_local_a1(matches[0])?;
    let indirect = exact_formula_call_arguments(matches[1], "INDIRECT")?;
    if indirect.is_empty() || indirect.len() > 2 {
        return None;
    }
    let sheet_name_reference = address[4].trim();
    let (sheet_name_column, sheet_name_row) = parse_local_a1(sheet_name_reference)?;
    if !formula_contains_identifier(indirect[0], sheet_name_reference.trim_matches('$')) {
        return None;
    }
    let range_column = parse_indirect_whole_column(indirect[0])?;
    let result_column = address[1].trim().parse::<i32>().ok()?;
    (result_column > 0).then_some(IndirectMatchAddress {
        lookup_row,
        lookup_column,
        sheet_name_row,
        sheet_name_column,
        range_column,
        result_column,
    })
}

fn parse_static_address(formula: &str) -> Option<StaticAddress> {
    let address = exact_formula_call_arguments(formula, "ADDRESS")?;
    if address.len() != 5 {
        return None;
    }
    let row = address[0].trim().parse::<i32>().ok()?;
    let column = address[1].trim().parse::<i32>().ok()?;
    let (sheet_name_column, sheet_name_row) = parse_local_a1(address[4])?;
    (row > 0 && column > 0).then_some(StaticAddress {
        row,
        column,
        sheet_name_row,
        sheet_name_column,
    })
}

fn resolve_preview_static_text(
    package: &Package<'_>,
    sheets: &[FormulaSheet],
    shared_strings: &[String],
    loaded_inputs: &mut [Option<CoordinateMap<InputValue>>],
    cell: CellReferenceIndex,
    visited: &mut HashSet<CellReferenceIndex>,
) -> Result<Option<String>, String> {
    if !visited.insert(cell) || visited.len() > 64 {
        return Ok(None);
    }
    load_preview_inputs(package, sheets, shared_strings, cell.sheet, loaded_inputs)?;
    let value = loaded_inputs[cell.sheet as usize]
        .as_ref()
        .and_then(|inputs| inputs.get(&(cell.row, cell.column)))
        .cloned();
    match value {
        Some(InputValue::Literal(value)) => Ok(value
            .strip_prefix('\'')
            .map(str::to_owned)
            .or_else(|| (!value.is_empty()).then_some(value))),
        Some(InputValue::Formula { formula, .. }) if !formula.is_empty() => {
            let formula = formula.trim().trim_start_matches('=');
            if let Some(value) = filename_fallback_sheet_name(formula) {
                return Ok(Some(value));
            }
            let Some((sheet, row, column)) =
                parse_cross_sheet_reference(formula, cell.sheet, sheets)
            else {
                return Ok(None);
            };
            resolve_preview_static_text(
                package,
                sheets,
                shared_strings,
                loaded_inputs,
                CellReferenceIndex { sheet, row, column },
                visited,
            )
        }
        Some(InputValue::Formula { .. }) | None => Ok(None),
    }
}

fn filename_fallback_sheet_name(formula: &str) -> Option<String> {
    let arguments = exact_formula_call_arguments(formula, "IF")?;
    if arguments.len() != 3
        || !formula_contains_identifier(arguments[0], "ISERROR")
        || !formula_contains_identifier(arguments[0], "CELL")
        || !arguments[0].to_ascii_lowercase().contains("\"filename\"")
    {
        return None;
    }
    let fallback = arguments[1].trim();
    fallback
        .strip_prefix('"')?
        .strip_suffix('"')
        .map(|value| value.replace("\"\"", "\""))
}

fn preview_target_sheet(
    package: &Package<'_>,
    sheets: &[FormulaSheet],
    shared_strings: &[String],
    loaded_inputs: &mut [Option<CoordinateMap<InputValue>>],
    source_sheet: u32,
    row: i32,
    column: i32,
) -> Result<Option<u32>, String> {
    let Some(name) = resolve_preview_static_text(
        package,
        sheets,
        shared_strings,
        loaded_inputs,
        CellReferenceIndex {
            sheet: source_sheet,
            row,
            column,
        },
        &mut HashSet::new(),
    )?
    else {
        return Ok(None);
    };
    Ok(sheets
        .iter()
        .position(|sheet| sheet.name.eq_ignore_ascii_case(&name))
        .and_then(|index| u32::try_from(index).ok()))
}

fn queue_indirect_address_inputs(
    package: &Package<'_>,
    sheets: &[FormulaSheet],
    shared_strings: &[String],
    bounds: &WorkbookBounds,
    loaded_inputs: &mut [Option<CoordinateMap<InputValue>>],
    source: (u32, IndirectMatchAddress),
    queue: (&mut Vec<CellReferenceIndex>, &mut CellSet),
) -> Result<(), String> {
    let (source_sheet, address) = source;
    let (pending, scheduled) = queue;
    let Some(target_sheet) = preview_target_sheet(
        package,
        sheets,
        shared_strings,
        loaded_inputs,
        source_sheet,
        address.sheet_name_row,
        address.sheet_name_column,
    )?
    else {
        return Ok(());
    };
    load_preview_inputs(package, sheets, shared_strings, target_sheet, loaded_inputs)?;
    let last_row = bounds
        .by_index
        .get(target_sheet as usize)
        .copied()
        .unwrap_or(1) as i32;
    if let Some(inputs) = loaded_inputs[target_sheet as usize].as_ref() {
        for &(row, column) in inputs.keys() {
            if row <= last_row
                && (column == address.range_column || column == address.result_column)
            {
                queue_preview_cell(
                    pending,
                    scheduled,
                    CellReferenceIndex {
                        sheet: target_sheet,
                        row,
                        column,
                    },
                );
            }
        }
    }
    Ok(())
}

fn queue_static_address_input(
    package: &Package<'_>,
    sheets: &[FormulaSheet],
    shared_strings: &[String],
    loaded_inputs: &mut [Option<CoordinateMap<InputValue>>],
    source: (u32, StaticAddress),
    queue: (&mut Vec<CellReferenceIndex>, &mut CellSet),
) -> Result<(), String> {
    let (source_sheet, address) = source;
    let (pending, scheduled) = queue;
    let Some(target_sheet) = preview_target_sheet(
        package,
        sheets,
        shared_strings,
        loaded_inputs,
        source_sheet,
        address.sheet_name_row,
        address.sheet_name_column,
    )?
    else {
        return Ok(());
    };
    queue_preview_cell(
        pending,
        scheduled,
        CellReferenceIndex {
            sheet: target_sheet,
            row: address.row,
            column: address.column,
        },
    );
    Ok(())
}

fn exact_formula_call_arguments<'a>(formula: &'a str, name: &str) -> Option<Vec<&'a str>> {
    let formula = formula.trim();
    if formula.len() <= name.len()
        || !formula[..name.len()].eq_ignore_ascii_case(name)
        || formula.as_bytes().get(name.len()) != Some(&b'(')
        || matching_parenthesis(formula, name.len()) != Some(formula.len() - 1)
    {
        return None;
    }
    split_formula_arguments(&formula[name.len() + 1..formula.len() - 1])
}

fn parse_local_a1(reference: &str) -> Option<(i32, i32)> {
    let reference = reference.trim().replace('$', "");
    (!reference.contains('!'))
        .then(|| parse_a1(&reference))
        .flatten()
}

fn parse_indirect_whole_column(expression: &str) -> Option<i32> {
    let after_bang = expression.rsplit_once('!')?.1;
    let normalized = after_bang
        .chars()
        .filter(|character| character.is_ascii_uppercase() || matches!(character, '$' | ':'))
        .collect::<String>();
    let (left, right) = normalized.split_once(':')?;
    let left = left.trim_matches('$');
    let right = right.trim_matches('$');
    if left != right {
        return None;
    }
    parse_a1(&format!("{left}1")).map(|(column, _)| column)
}

fn resolve_indirect_match_address(
    model: &Model<'_>,
    source_to_model_sheet: &[u32],
    sheets: &[FormulaSheet],
    bounds: &WorkbookBounds,
    source_sheet: u32,
    address: IndirectMatchAddress,
) -> Result<Option<String>, String> {
    let model_sheet = source_to_model_sheet[source_sheet as usize];
    let CellValue::String(target_sheet_name) = model.get_cell_value_by_index(
        model_sheet,
        address.sheet_name_row,
        address.sheet_name_column,
    )?
    else {
        return Ok(None);
    };
    let Some(target_source_sheet) = sheets
        .iter()
        .position(|sheet| sheet.name.eq_ignore_ascii_case(&target_sheet_name))
    else {
        return Ok(None);
    };
    let lookup =
        model.get_cell_value_by_index(model_sheet, address.lookup_row, address.lookup_column)?;
    let target_model_sheet = source_to_model_sheet[target_source_sheet];
    let last_row = bounds
        .by_index
        .get(target_source_sheet)
        .copied()
        .unwrap_or(1);
    for row in 1..=last_row {
        let row = i32::try_from(row).map_err(|_| "formula row exceeds evaluator range")?;
        let candidate =
            model.get_cell_value_by_index(target_model_sheet, row, address.range_column)?;
        if excel_match_equal(&lookup, &candidate) {
            return Ok(Some(format_absolute_reference(
                &sheets[target_source_sheet].name,
                address.result_column,
                row,
            )));
        }
    }
    Ok(None)
}

fn excel_match_equal(left: &CellValue, right: &CellValue) -> bool {
    match (left, right) {
        (CellValue::Boolean(left), CellValue::Boolean(right)) => left == right,
        (CellValue::Number(left), CellValue::Number(right)) => left == right,
        (CellValue::String(left), CellValue::String(right)) => left.eq_ignore_ascii_case(right),
        (CellValue::None, CellValue::None)
        | (CellValue::None, _)
        | (_, CellValue::None)
        | (CellValue::Boolean(_), _)
        | (CellValue::Number(_), _)
        | (CellValue::String(_), _) => false,
    }
}

fn format_absolute_reference(sheet: &str, mut column: i32, row: i32) -> String {
    let mut column_name = String::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        column_name.insert(0, char::from(b'A' + remainder as u8));
        column = (column - 1) / 26;
    }
    let sheet = if sheet
        .chars()
        .all(|character| character.is_alphanumeric() || matches!(character, '_' | '.'))
    {
        sheet.to_owned()
    } else {
        format!("'{}'", sheet.replace('\'', "''"))
    };
    format!("{sheet}!${column_name}${row}")
}

fn parse_cross_sheet_reference(
    reference: &str,
    current_sheet: u32,
    sheets: &[FormulaSheet],
) -> Option<(u32, i32, i32)> {
    let reference = reference.trim();
    let (sheet, address) = match reference.rsplit_once('!') {
        Some((sheet, address)) => {
            let sheet = sheet.trim();
            let sheet = sheet
                .strip_prefix('\'')
                .and_then(|value| value.strip_suffix('\''))
                .unwrap_or(sheet)
                .replace("''", "'");
            let index = sheets
                .iter()
                .position(|candidate| candidate.name == sheet)?;
            (u32::try_from(index).ok()?, address)
        }
        None => (current_sheet, reference),
    };
    let (column, row) = parse_a1(&address.replace('$', ""))?;
    Some((sheet, row, column))
}

fn cell_value_input(value: CellValue, cell_type: CellType) -> String {
    match value {
        CellValue::None => String::new(),
        CellValue::Boolean(value) => if value { "TRUE" } else { "FALSE" }.to_owned(),
        CellValue::Number(value) => value.to_string(),
        CellValue::String(value) if cell_type == CellType::ErrorValue => value,
        CellValue::String(value) => text_literal(&value),
    }
}

fn formula_value(value: CellValue, cell_type: CellType) -> FormulaValue {
    match value {
        CellValue::None => FormulaValue::Blank,
        CellValue::Boolean(value) => FormulaValue::Boolean(value),
        CellValue::Number(value) => FormulaValue::Number(value),
        CellValue::String(value) if cell_type == CellType::ErrorValue => FormulaValue::Error(value),
        CellValue::String(value) => FormulaValue::Text(value),
    }
}

fn restore_sheet_name_quotes(formula: &str, sheet_names: &[String]) -> String {
    let mut quoted_names = sheet_names
        .iter()
        .filter(|name| {
            name.chars()
                .any(|character| !(character.is_alphanumeric() || matches!(character, '_' | '.')))
        })
        .collect::<Vec<_>>();
    quoted_names.sort_by_key(|name| std::cmp::Reverse(name.len()));
    let mut output = String::with_capacity(formula.len());
    let mut index = 0_usize;
    let mut in_string = false;
    while index < formula.len() {
        if formula.as_bytes()[index] == b'"' {
            output.push('"');
            if in_string && formula.as_bytes().get(index + 1) == Some(&b'"') {
                output.push('"');
                index += 2;
                continue;
            }
            in_string = !in_string;
            index += 1;
            continue;
        }
        if !in_string
            && let Some(name) = quoted_names.iter().find(|name| {
                formula[index..].starts_with(name.as_str())
                    && formula.as_bytes().get(index + name.len()) == Some(&b'!')
                    && (index == 0 || !is_reference_identifier_byte(formula.as_bytes()[index - 1]))
            })
        {
            output.push('\'');
            output.push_str(&name.replace('\'', "''"));
            output.push_str("'!");
            index += name.len() + 1;
            continue;
        }
        let character = formula[index..]
            .chars()
            .next()
            .expect("valid UTF-8 formula");
        output.push(character);
        index += character.len_utf8();
    }
    output
}

#[derive(Clone, Debug)]
struct CellInput {
    sheet: u32,
    row: i32,
    column: i32,
    value: InputValue,
}

#[derive(Clone, Debug)]
enum PreviewModelInput {
    Literal(String),
    Formula(Node),
}

#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
struct PreviewReference {
    sheet: u32,
    start_row: i32,
    start_column: i32,
    end_row: i32,
    end_column: i32,
}

impl PreviewReference {
    fn is_range(self) -> bool {
        self.start_row != self.end_row || self.start_column != self.end_column
    }
}

fn load_preview_inputs(
    package: &Package<'_>,
    sheets: &[FormulaSheet],
    shared_strings: &[String],
    sheet: u32,
    loaded: &mut [Option<CoordinateMap<InputValue>>],
) -> Result<(), String> {
    let index = usize::try_from(sheet).map_err(|_| "preview formula sheet is invalid")?;
    if loaded.get(index).is_none() {
        return Err("preview formula sheet exceeds the workbook range".to_owned());
    }
    if loaded[index].is_some() {
        return Ok(());
    }
    let Some(part) = sheets[index].part.as_deref() else {
        loaded[index] = Some(CoordinateMap::default());
        return Ok(());
    };
    let bytes = package
        .required_part(part)
        .map_err(|error| error.to_string())?;
    let inputs = parse_sheet_inputs(&bytes, package, sheet, shared_strings, part)?;
    loaded[index] = Some(
        inputs
            .into_iter()
            .map(|input| ((input.row, input.column), input.value))
            .collect(),
    );
    Ok(())
}

fn preview_formula(
    parser: &mut Parser<'_>,
    sheet_names: &[String],
    inputs: &Option<CoordinateMap<InputValue>>,
    cell: CellReferenceIndex,
    formula: String,
    shared: Option<u32>,
    shared_master_nodes: &mut SharedFormulaMap,
) -> Result<String, String> {
    if !formula.is_empty() {
        return Ok(formula);
    }
    let Some(shared) = shared else {
        return Ok(String::new());
    };
    let key = (cell.sheet, shared);
    let node = if let Some(node) = shared_master_nodes.get(&key) {
        node.clone()
    } else {
        let Some(inputs) = inputs.as_ref() else {
            return Err("preview shared-formula inputs are unavailable".to_owned());
        };
        let Some((&(master_row, master_column), master_formula)) =
            inputs.iter().find_map(|(position, value)| match value {
                InputValue::Formula {
                    formula,
                    shared: Some(candidate),
                    ..
                } if *candidate == shared && !formula.is_empty() => Some((position, formula)),
                InputValue::Literal(_) | InputValue::Formula { .. } => None,
            })
        else {
            return Err(format!(
                "shared formula {shared} has no master on preview sheet {}",
                cell.sheet
            ));
        };
        let master_context = CellReferenceRC {
            sheet: sheet_names[cell.sheet as usize].clone(),
            row: master_row,
            column: master_column,
        };
        let node = parser.parse(master_formula, &master_context);
        shared_master_nodes.insert(key, node.clone());
        node
    };
    let target_context = CellReferenceRC {
        sheet: sheet_names[cell.sheet as usize].clone(),
        row: cell.row,
        column: cell.column,
    };
    Ok(to_english_string(&node, &target_context))
}

fn collect_preview_references(
    node: &Node,
    parser: &mut Parser<'_>,
    context: &CellReferenceRC,
    sheet_names: &[String],
    remaining_depth: u8,
    output: &mut Vec<PreviewReference>,
) -> Result<(), String> {
    let relative = |absolute: bool, value: i32, base: i32| {
        if absolute { value } else { base + value }
    };
    match node {
        Node::ReferenceKind {
            sheet_index,
            absolute_row,
            absolute_column,
            row,
            column,
            ..
        } => output.push(PreviewReference {
            sheet: *sheet_index,
            start_row: relative(*absolute_row, *row, context.row),
            start_column: relative(*absolute_column, *column, context.column),
            end_row: relative(*absolute_row, *row, context.row),
            end_column: relative(*absolute_column, *column, context.column),
        }),
        Node::RangeKind {
            sheet_index,
            absolute_row1,
            absolute_column1,
            row1,
            column1,
            absolute_row2,
            absolute_column2,
            row2,
            column2,
            ..
        } => output.push(PreviewReference {
            sheet: *sheet_index,
            start_row: relative(*absolute_row1, *row1, context.row),
            start_column: relative(*absolute_column1, *column1, context.column),
            end_row: relative(*absolute_row2, *row2, context.row),
            end_column: relative(*absolute_column2, *column2, context.column),
        }),
        Node::DefinedNameKind((_, scope, formula)) if remaining_depth > 0 => {
            let sheet = scope
                .and_then(|sheet| sheet_names.get(sheet as usize))
                .cloned()
                .unwrap_or_else(|| context.sheet.clone());
            let defined_context = CellReferenceRC {
                sheet,
                row: context.row,
                column: context.column,
            };
            let defined = parser.parse(formula, &defined_context);
            collect_preview_references(
                &defined,
                parser,
                &defined_context,
                sheet_names,
                remaining_depth - 1,
                output,
            )?;
        }
        Node::OpRangeKind { left, right }
        | Node::OpConcatenateKind { left, right }
        | Node::OpSumKind { left, right, .. }
        | Node::OpProductKind { left, right, .. }
        | Node::OpPowerKind { left, right }
        | Node::CompareKind { left, right, .. } => {
            collect_preview_references(
                left,
                parser,
                context,
                sheet_names,
                remaining_depth,
                output,
            )?;
            collect_preview_references(
                right,
                parser,
                context,
                sheet_names,
                remaining_depth,
                output,
            )?;
        }
        Node::FunctionKind { args, .. } | Node::NamedFunctionKind { args, .. } => {
            for argument in args {
                collect_preview_references(
                    argument,
                    parser,
                    context,
                    sheet_names,
                    remaining_depth,
                    output,
                )?;
            }
        }
        Node::LambdaDefKind { body, .. } => {
            collect_preview_references(body, parser, context, sheet_names, remaining_depth, output)?
        }
        Node::LambdaCallKind { lambda, args } => {
            collect_preview_references(
                lambda,
                parser,
                context,
                sheet_names,
                remaining_depth,
                output,
            )?;
            for argument in args {
                collect_preview_references(
                    argument,
                    parser,
                    context,
                    sheet_names,
                    remaining_depth,
                    output,
                )?;
            }
        }
        Node::ImplicitIntersection { child, .. } | Node::SpillRangeOperator { child } => {
            collect_preview_references(
                child,
                parser,
                context,
                sheet_names,
                remaining_depth,
                output,
            )?;
        }
        Node::UnaryKind { right, .. } => collect_preview_references(
            right,
            parser,
            context,
            sheet_names,
            remaining_depth,
            output,
        )?,
        Node::WrongReferenceKind { .. }
        | Node::WrongRangeKind { .. }
        | Node::ParseErrorKind { .. }
        | Node::TableNameKind(_)
        | Node::DefinedNameKind(_) => {
            return Err("preview formula has an unresolved reference".to_owned());
        }
        Node::BooleanKind(_)
        | Node::NumberKind(_)
        | Node::StringKind(_)
        | Node::ArrayKind(_)
        | Node::NamedVariableKind { .. }
        | Node::ErrorKind(_)
        | Node::EmptyArgKind => {}
    }
    Ok(())
}

fn queue_preview_reference(
    inputs: &CoordinateMap<InputValue>,
    reference: PreviewReference,
    pending: &mut Vec<CellReferenceIndex>,
    scheduled: &mut CellSet,
) {
    let start_row = reference.start_row.min(reference.end_row);
    let end_row = reference.start_row.max(reference.end_row);
    let start_column = reference.start_column.min(reference.end_column);
    let end_column = reference.start_column.max(reference.end_column);
    let rows = i64::from(end_row) - i64::from(start_row) + 1;
    let columns = i64::from(end_column) - i64::from(start_column) + 1;
    let area = rows
        .checked_mul(columns)
        .and_then(|value| usize::try_from(value).ok())
        .unwrap_or(usize::MAX);
    if area <= inputs.len() {
        for row in start_row..=end_row {
            for column in start_column..=end_column {
                if inputs.contains_key(&(row, column)) {
                    queue_preview_cell(
                        pending,
                        scheduled,
                        CellReferenceIndex {
                            sheet: reference.sheet,
                            row,
                            column,
                        },
                    );
                }
            }
        }
        return;
    }
    for &(row, column) in inputs.keys() {
        if row >= start_row && row <= end_row && column >= start_column && column <= end_column {
            queue_preview_cell(
                pending,
                scheduled,
                CellReferenceIndex {
                    sheet: reference.sheet,
                    row,
                    column,
                },
            );
        }
    }
}

fn queue_preview_cell(
    pending: &mut Vec<CellReferenceIndex>,
    scheduled: &mut CellSet,
    cell: CellReferenceIndex,
) {
    if scheduled.insert(cell) {
        pending.push(cell);
    }
}

#[derive(Clone, Copy, Debug)]
struct HyperlinkIndirectCell {
    sheet: u32,
    row: i32,
    column: i32,
    reference_row: i32,
    reference_column: i32,
    match_address: Option<IndirectMatchAddress>,
}

#[derive(Clone, Copy, Debug)]
struct IndirectMatchAddress {
    lookup_row: i32,
    lookup_column: i32,
    sheet_name_row: i32,
    sheet_name_column: i32,
    range_column: i32,
    result_column: i32,
}

#[derive(Clone, Copy, Debug)]
struct StaticAddress {
    row: i32,
    column: i32,
    sheet_name_row: i32,
    sheet_name_column: i32,
}

#[derive(Clone, Debug)]
enum InputValue {
    Literal(String),
    Formula {
        formula: String,
        shared: Option<u32>,
        array_range: Option<ArrayFormulaRange>,
    },
}

#[derive(Clone, Copy, Debug)]
struct ArrayFormulaRange {
    start_row: i32,
    start_column: i32,
    end_row: i32,
    end_column: i32,
}

impl ArrayFormulaRange {
    fn width(self) -> Result<i32, String> {
        self.end_column
            .checked_sub(self.start_column)
            .and_then(|value| value.checked_add(1))
            .filter(|value| *value > 0)
            .ok_or_else(|| "invalid array-formula width".to_owned())
    }

    fn height(self) -> Result<i32, String> {
        self.end_row
            .checked_sub(self.start_row)
            .and_then(|value| value.checked_add(1))
            .filter(|value| *value > 0)
            .ok_or_else(|| "invalid array-formula height".to_owned())
    }
}

#[derive(Clone, Debug)]
struct SharedFormula {
    node: Node,
}

#[derive(Debug)]
struct ParsedCell {
    depth: usize,
    row: i32,
    column: i32,
    value_type: String,
    value_depth: Option<usize>,
    value: String,
    inline_text_depth: Option<usize>,
    inline_text: String,
    formula_depth: Option<usize>,
    formula: String,
    formula_shared: Option<u32>,
    formula_array_range: Option<ArrayFormulaRange>,
    has_formula: bool,
}

fn parse_sheet_inputs(
    bytes: &[u8],
    package: &Package<'_>,
    sheet: u32,
    shared_strings: &[String],
    part: &str,
) -> Result<Vec<CellInput>, String> {
    let mut depth = 0_usize;
    let mut current_row = 1_i32;
    let mut next_row = 1_i32;
    let mut next_column = 1_i32;
    let mut cell = None::<ParsedCell>;
    let mut inputs = Vec::new();
    parse_xml(bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } => {
                let local = local_name(name);
                if local == "row" {
                    current_row = attribute(&attributes, "r")
                        .map(|value| value.parse::<i32>())
                        .transpose()
                        .map_err(|_| formula_error(part, "invalid row index"))?
                        .unwrap_or(next_row);
                    next_row = current_row.saturating_add(1);
                    next_column = 1;
                } else if local == "c" {
                    let (column, row) = match attribute(&attributes, "r") {
                        Some(address) => parse_a1(address)
                            .ok_or_else(|| formula_error(part, "invalid formula cell address"))?,
                        None => (next_column, current_row),
                    };
                    next_column = column.saturating_add(1);
                    cell = Some(ParsedCell {
                        depth,
                        row,
                        column,
                        value_type: attribute(&attributes, "t").unwrap_or("n").to_owned(),
                        value_depth: None,
                        value: String::new(),
                        inline_text_depth: None,
                        inline_text: String::new(),
                        formula_depth: None,
                        formula: String::new(),
                        formula_shared: None,
                        formula_array_range: None,
                        has_formula: false,
                    });
                } else if let Some(cell) = cell.as_mut() {
                    match local {
                        "v" => cell.value_depth = (!empty).then_some(depth),
                        "t" if cell.value_type == "inlineStr" => {
                            cell.inline_text_depth = (!empty).then_some(depth)
                        }
                        "f" => {
                            cell.has_formula = true;
                            cell.formula_depth = (!empty).then_some(depth);
                            if attribute(&attributes, "t") == Some("array") {
                                let reference = attribute(&attributes, "ref").ok_or_else(|| {
                                    formula_error(part, "array formula has no range")
                                })?;
                                cell.formula_array_range =
                                    Some(parse_array_formula_range(reference).ok_or_else(
                                        || formula_error(part, "invalid array-formula range"),
                                    )?);
                            }
                            if attribute(&attributes, "t") == Some("shared") {
                                cell.formula_shared = Some(
                                    attribute(&attributes, "si")
                                        .ok_or_else(|| {
                                            formula_error(part, "shared formula has no index")
                                        })?
                                        .parse::<u32>()
                                        .map_err(|_| {
                                            formula_error(part, "invalid shared-formula index")
                                        })?,
                                );
                            }
                        }
                        _ => {}
                    }
                }
                if !empty {
                    depth = depth.saturating_add(1);
                }
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                let local = local_name(name);
                if let Some(cell) = cell.as_mut() {
                    if local == "v" {
                        cell.value_depth = None;
                    } else if local == "t" {
                        cell.inline_text_depth = None;
                    } else if local == "f" {
                        cell.formula_depth = None;
                    }
                }
                if local == "c" && cell.as_ref().is_some_and(|cell| cell.depth == depth) {
                    let cell = cell
                        .take()
                        .ok_or_else(|| formula_error(part, "formula cell state was lost"))?;
                    let value = if cell.has_formula {
                        InputValue::Formula {
                            formula: cell.formula,
                            shared: cell.formula_shared,
                            array_range: cell.formula_array_range,
                        }
                    } else {
                        InputValue::Literal(cell_literal(&cell, shared_strings, part)?)
                    };
                    inputs.push(CellInput {
                        sheet,
                        row: cell.row,
                        column: cell.column,
                        value,
                    });
                }
            }
            XmlEvent::Text(text) => append_cell_text(cell.as_mut(), text, true, part)?,
            XmlEvent::Cdata(text) => append_cell_text(cell.as_mut(), text, false, part)?,
        }
        Ok(())
    })
    .map_err(|error| error.to_string())?;
    Ok(inputs)
}

fn append_cell_text(
    cell: Option<&mut ParsedCell>,
    text: &str,
    decode: bool,
    part: &str,
) -> Result<(), crate::diagnostic::Diagnostic> {
    let Some(cell) = cell else {
        return Ok(());
    };
    let text = if decode {
        decode_xml_text(text)
            .map_err(|error| error.in_part(part))?
            .into_owned()
    } else {
        text.to_owned()
    };
    if cell.value_depth.is_some() {
        cell.value.push_str(&text);
    } else if cell.inline_text_depth.is_some() {
        cell.inline_text.push_str(&text);
    } else if cell.formula_depth.is_some() {
        cell.formula.push_str(&text);
    }
    Ok(())
}

fn cell_literal(
    cell: &ParsedCell,
    shared_strings: &[String],
    part: &str,
) -> Result<String, crate::diagnostic::Diagnostic> {
    let value = match cell.value_type.as_str() {
        "s" => cell
            .value
            .parse::<usize>()
            .ok()
            .and_then(|index| shared_strings.get(index))
            .cloned()
            .map(|value| text_literal(&value))
            .ok_or_else(|| formula_error(part, "invalid shared-string formula input"))?,
        "inlineStr" => text_literal(&cell.inline_text),
        "b" => match cell.value.as_str() {
            "0" | "false" => "FALSE".to_owned(),
            "1" | "true" => "TRUE".to_owned(),
            _ => return Err(formula_error(part, "invalid boolean formula input")),
        },
        "n" | "e" => cell.value.clone(),
        "str" | "d" => text_literal(&cell.value),
        _ => return Err(formula_error(part, "unsupported formula input type")),
    };
    Ok(value)
}

fn text_literal(value: &str) -> String {
    format!("'{value}")
}

#[derive(Debug)]
struct DefinedName {
    name: String,
    scope: Option<u32>,
    formula: String,
}

fn formula_sheet_order(
    sheets: &[FormulaSheet],
    inputs: &[CellInput],
    defined_names: &[DefinedName],
) -> Vec<usize> {
    let array_sheets = inputs
        .iter()
        .filter_map(|input| match &input.value {
            InputValue::Formula {
                array_range: Some(_),
                ..
            } => Some(input.sheet as usize),
            InputValue::Literal(_)
            | InputValue::Formula {
                array_range: None, ..
            } => None,
        })
        .collect::<HashSet<_>>();
    let mut dependencies = vec![HashSet::<usize>::new(); sheets.len()];
    for input in inputs {
        let InputValue::Formula {
            formula,
            array_range: Some(_),
            ..
        } = &input.value
        else {
            continue;
        };
        let source = input.sheet as usize;
        for dependency in &array_sheets {
            if *dependency != source && formula_references_sheet(formula, &sheets[*dependency].name)
            {
                dependencies[source].insert(*dependency);
            }
        }
        for defined_name in defined_names {
            if !formula_contains_identifier(formula, &defined_name.name) {
                continue;
            }
            for dependency in &array_sheets {
                if *dependency != source
                    && formula_references_sheet(&defined_name.formula, &sheets[*dependency].name)
                {
                    dependencies[source].insert(*dependency);
                }
            }
        }
    }

    fn visit(
        sheet: usize,
        dependencies: &[HashSet<usize>],
        state: &mut [u8],
        order: &mut Vec<usize>,
    ) {
        if state[sheet] == 2 {
            return;
        }
        if state[sheet] == 1 {
            return;
        }
        state[sheet] = 1;
        let mut sorted = dependencies[sheet].iter().copied().collect::<Vec<_>>();
        sorted.sort_unstable();
        for dependency in sorted {
            visit(dependency, dependencies, state, order);
        }
        state[sheet] = 2;
        order.push(sheet);
    }

    let mut state = vec![0_u8; sheets.len()];
    let mut order = Vec::with_capacity(sheets.len());
    let mut sorted_array_sheets = array_sheets.into_iter().collect::<Vec<_>>();
    sorted_array_sheets.sort_unstable();
    for sheet in sorted_array_sheets {
        visit(sheet, &dependencies, &mut state, &mut order);
    }
    for (sheet, &sheet_state) in state.iter().enumerate() {
        if sheet_state == 0 {
            order.push(sheet);
        }
    }
    order
}

fn formula_references_sheet(formula: &str, sheet_name: &str) -> bool {
    formula.contains(&format!("{sheet_name}!"))
        || formula.contains(&format!("'{}'!", sheet_name.replace('\'', "''")))
}

fn formula_contains_identifier(formula: &str, identifier: &str) -> bool {
    formula.match_indices(identifier).any(|(start, matched)| {
        let end = start + matched.len();
        let bytes = formula.as_bytes();
        (start == 0 || !is_reference_identifier_byte(bytes[start - 1]))
            && (end == bytes.len() || !is_reference_identifier_byte(bytes[end]))
    })
}

struct WorkbookBounds {
    by_index: Vec<u32>,
    by_name: HashMap<String, u32>,
    maximum: u32,
}

impl WorkbookBounds {
    fn new(sheets: &[FormulaSheet], by_index: Vec<u32>) -> Self {
        let maximum = by_index.iter().copied().max().unwrap_or(1);
        let by_name = sheets
            .iter()
            .zip(by_index.iter().copied())
            .map(|(sheet, rows)| (sheet.name.to_uppercase(), rows))
            .collect();
        Self {
            by_index,
            by_name,
            maximum,
        }
    }

    fn rows_for_reference(&self, formula: &str, start: usize, current: Option<u32>) -> u32 {
        reference_sheet_name(formula, start)
            .and_then(|name| self.by_name.get(&name.to_uppercase()).copied())
            .or_else(|| current.and_then(|sheet| self.by_index.get(sheet as usize).copied()))
            .unwrap_or(self.maximum)
            .max(1)
    }
}

fn parse_defined_names(
    package: &Package<'_>,
    workbook_part: &str,
) -> Result<Vec<DefinedName>, String> {
    #[derive(Debug)]
    struct CurrentName {
        depth: usize,
        value: DefinedName,
    }

    let bytes = package
        .required_part(workbook_part)
        .map_err(|error| error.to_string())?;
    let mut depth = 0_usize;
    let mut current = None::<CurrentName>;
    let mut names = Vec::new();
    parse_xml(&bytes, package.limits(), |event| {
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                empty,
            } if local_name(name) == "definedName" => {
                let defined_name = DefinedName {
                    name: attribute(&attributes, "name")
                        .ok_or_else(|| formula_error(workbook_part, "defined name is unnamed"))?
                        .to_owned(),
                    scope: attribute(&attributes, "localSheetId")
                        .map(|scope| scope.parse::<u32>())
                        .transpose()
                        .map_err(|_| formula_error(workbook_part, "invalid defined-name scope"))?,
                    formula: String::new(),
                };
                if empty {
                    names.push(defined_name);
                } else {
                    current = Some(CurrentName {
                        depth,
                        value: defined_name,
                    });
                }
                depth = depth.saturating_add(usize::from(!empty));
            }
            XmlEvent::StartElement { empty, .. } => {
                depth = depth.saturating_add(usize::from(!empty));
            }
            XmlEvent::EndElement { name } => {
                depth = depth.saturating_sub(1);
                if local_name(name) == "definedName"
                    && current.as_ref().is_some_and(|name| name.depth == depth)
                {
                    names.push(current.take().expect("defined-name state is present").value);
                }
            }
            XmlEvent::Text(text) => {
                if let Some(current) = current.as_mut() {
                    current.value.formula.push_str(
                        &decode_xml_text(text).map_err(|error| error.in_part(workbook_part))?,
                    );
                }
            }
            XmlEvent::Cdata(text) => {
                if let Some(current) = current.as_mut() {
                    current.value.formula.push_str(text);
                }
            }
        }
        Ok(())
    })
    .map_err(|error| error.to_string())?;
    Ok(names)
}

fn bound_whole_column_references(
    formula: &str,
    current_sheet: Option<u32>,
    bounds: &WorkbookBounds,
) -> String {
    let bytes = formula.as_bytes();
    let mut output = String::with_capacity(formula.len());
    let mut index = 0_usize;
    let mut in_string = false;
    while index < bytes.len() {
        if bytes[index] == b'"' {
            output.push('"');
            if in_string && bytes.get(index + 1) == Some(&b'"') {
                output.push('"');
                index += 2;
                continue;
            }
            in_string = !in_string;
            index += 1;
            continue;
        }
        let indirect_string_reference = in_string && index > 0 && bytes[index - 1] == b'!';
        if (!in_string || indirect_string_reference)
            && let Some((end, left, right)) = whole_column_range_at(formula, index)
            && !dimension_sensitive_reference(formula, index)
        {
            let used_row_count = if indirect_string_reference {
                bounds.maximum
            } else {
                bounds.rows_for_reference(formula, index, current_sheet)
            };
            output.push_str(left);
            output.push_str("$1:");
            output.push_str(right);
            output.push('$');
            output.push_str(&used_row_count.to_string());
            index = end;
            continue;
        }
        let character = formula[index..]
            .chars()
            .next()
            .expect("valid UTF-8 formula");
        output.push(character);
        index += character.len_utf8();
    }
    output
}

fn has_dynamic_array_function(formula: &str) -> bool {
    let upper = formula.to_ascii_uppercase();
    [
        "FILTER(",
        "UNIQUE(",
        "SORT(",
        "SORTBY(",
        "SEQUENCE(",
        "RANDARRAY(",
        "TOCOL(",
        "TOROW(",
        "WRAPROWS(",
        "WRAPCOLS(",
        "TAKE(",
        "DROP(",
        "CHOOSECOLS(",
        "CHOOSEROWS(",
        "EXPAND(",
        "HSTACK(",
        "VSTACK(",
    ]
    .iter()
    .any(|name| {
        upper.match_indices(name).any(|(start, _)| {
            start == 0
                || !upper.as_bytes()[start - 1].is_ascii_alphanumeric()
                    && upper.as_bytes()[start - 1] != b'_'
        })
    })
}

fn has_nondeterministic_function(formula: &str) -> bool {
    formula_calls_any(formula, &["RAND", "RANDBETWEEN", "RANDARRAY"])
}

fn rewrite_hyperlink_display(formula: &str) -> String {
    let mut output = String::with_capacity(formula.len());
    let mut index = 0_usize;
    let mut in_string = false;
    while index < formula.len() {
        let byte = formula.as_bytes()[index];
        if byte == b'"' {
            output.push('"');
            if in_string && formula.as_bytes().get(index + 1) == Some(&b'"') {
                output.push('"');
                index += 2;
                continue;
            }
            in_string = !in_string;
            index += 1;
            continue;
        }
        if !in_string && let Some((end, replacement)) = hyperlink_display_call_at(formula, index) {
            output.push_str(&replacement);
            index = end;
            continue;
        }
        let character = formula[index..]
            .chars()
            .next()
            .expect("valid UTF-8 formula");
        output.push(character);
        index += character.len_utf8();
    }
    output
}

fn rewrite_address_defaults(formula: &str) -> String {
    rewrite_formula_calls(formula, "ADDRESS", |arguments| {
        if !(2..=5).contains(&arguments.len()) {
            return None;
        }
        let mut arguments = arguments
            .into_iter()
            .map(|value| rewrite_address_defaults(value.trim()))
            .collect::<Vec<_>>();
        if arguments.len() >= 3 && arguments[2].is_empty() {
            arguments[2] = "1".to_owned();
        }
        if arguments.len() >= 4 && arguments[3].is_empty() {
            arguments[3] = "TRUE".to_owned();
        }
        Some(format!("ADDRESS({})", arguments.join(",")))
    })
}

fn rewrite_formula_calls<F>(formula: &str, name: &str, mut rewrite: F) -> String
where
    F: FnMut(Vec<&str>) -> Option<String>,
{
    let mut output = String::with_capacity(formula.len());
    let mut index = 0_usize;
    let mut in_string = false;
    while index < formula.len() {
        let byte = formula.as_bytes()[index];
        if byte == b'"' {
            output.push('"');
            if in_string && formula.as_bytes().get(index + 1) == Some(&b'"') {
                output.push('"');
                index += 2;
                continue;
            }
            in_string = !in_string;
            index += 1;
            continue;
        }
        let call = !in_string
            && (index == 0 || !is_reference_identifier_byte(formula.as_bytes()[index - 1]))
            && formula[index..].len() > name.len()
            && formula[index..index + name.len()].eq_ignore_ascii_case(name)
            && formula.as_bytes().get(index + name.len()) == Some(&b'(');
        if call {
            let open = index + name.len();
            if let Some(close) = matching_parenthesis(formula, open)
                && let Some(arguments) = split_formula_arguments(&formula[open + 1..close])
                && let Some(replacement) = rewrite(arguments)
            {
                output.push_str(&replacement);
                index = close + 1;
                continue;
            }
        }
        let character = formula[index..]
            .chars()
            .next()
            .expect("valid UTF-8 formula");
        output.push(character);
        index += character.len_utf8();
    }
    output
}

fn hyperlink_display_call_at(formula: &str, start: usize) -> Option<(usize, String)> {
    const NAME: &str = "HYPERLINK";
    if start > 0 && is_reference_identifier_byte(formula.as_bytes()[start - 1]) {
        return None;
    }
    if formula[start..].len() <= NAME.len()
        || !formula[start..start + NAME.len()].eq_ignore_ascii_case(NAME)
        || formula.as_bytes().get(start + NAME.len()) != Some(&b'(')
    {
        return None;
    }
    let open = start + NAME.len();
    let close = matching_parenthesis(formula, open)?;
    let arguments = split_formula_arguments(&formula[open + 1..close])?;
    let display = match arguments.as_slice() {
        [target] => target,
        [_, display] => display,
        _ => return None,
    };
    Some((close + 1, rewrite_hyperlink_display(display.trim())))
}

fn matching_parenthesis(formula: &str, open: usize) -> Option<usize> {
    let mut depth = 0_usize;
    let mut in_string = false;
    let bytes = formula.as_bytes();
    let mut index = open;
    while index < bytes.len() {
        match bytes[index] {
            b'"' => {
                if in_string && bytes.get(index + 1) == Some(&b'"') {
                    index += 2;
                    continue;
                }
                in_string = !in_string;
            }
            b'(' if !in_string => depth += 1,
            b')' if !in_string => {
                depth = depth.checked_sub(1)?;
                if depth == 0 {
                    return Some(index);
                }
            }
            _ => {}
        }
        index += 1;
    }
    None
}

fn split_formula_arguments(arguments: &str) -> Option<Vec<&str>> {
    let mut result = Vec::new();
    let mut depth = 0_usize;
    let mut in_string = false;
    let mut start = 0_usize;
    let bytes = arguments.as_bytes();
    let mut index = 0_usize;
    while index < bytes.len() {
        match bytes[index] {
            b'"' => {
                if in_string && bytes.get(index + 1) == Some(&b'"') {
                    index += 2;
                    continue;
                }
                in_string = !in_string;
            }
            b'(' if !in_string => depth += 1,
            b')' if !in_string => depth = depth.checked_sub(1)?,
            b',' if !in_string && depth == 0 => {
                result.push(&arguments[start..index]);
                start = index + 1;
            }
            _ => {}
        }
        index += 1;
    }
    if in_string || depth != 0 {
        return None;
    }
    result.push(&arguments[start..]);
    Some(result)
}

fn reference_sheet_name(formula: &str, column_start: usize) -> Option<String> {
    let prefix = formula.get(..column_start)?;
    let bang = prefix.len().checked_sub(1)?;
    if prefix.as_bytes().get(bang) != Some(&b'!') {
        return None;
    }
    let before_bang = &prefix[..bang];
    if let Some(quoted_name) = before_bang.strip_suffix('\'') {
        let mut search_end = quoted_name.len();
        let opening = loop {
            let quote = quoted_name[..search_end].rfind('\'')?;
            if quote > 0 && quoted_name.as_bytes().get(quote - 1) == Some(&b'\'') {
                search_end = quote - 1;
                continue;
            }
            break quote;
        };
        return Some(quoted_name[opening + 1..].replace("''", "'"));
    }
    let start = before_bang
        .rfind(|character: char| {
            !(character.is_alphanumeric() || matches!(character, '_' | '&' | '.'))
        })
        .map_or(0, |index| index + 1);
    (start < before_bang.len()).then(|| before_bang[start..].to_owned())
}

fn dimension_sensitive_reference(formula: &str, start: usize) -> bool {
    let prefix = formula[..start].trim_end().to_ascii_uppercase();
    prefix.ends_with("ROWS(") || prefix.ends_with("COLUMNS(")
}

fn whole_column_range_at(formula: &str, start: usize) -> Option<(usize, &str, &str)> {
    let bytes = formula.as_bytes();
    if start > 0 && is_reference_identifier_byte(bytes[start - 1]) {
        return None;
    }
    let mut cursor = start;
    if bytes.get(cursor) == Some(&b'$') {
        cursor += 1;
    }
    let left_start = start;
    let letters_start = cursor;
    while bytes.get(cursor).is_some_and(u8::is_ascii_uppercase) {
        cursor += 1;
    }
    if cursor == letters_start || cursor - letters_start > 3 || bytes.get(cursor) != Some(&b':') {
        return None;
    }
    let left = &formula[left_start..cursor];
    cursor += 1;
    let right_start = cursor;
    if bytes.get(cursor) == Some(&b'$') {
        cursor += 1;
    }
    let right_letters_start = cursor;
    while bytes.get(cursor).is_some_and(u8::is_ascii_uppercase) {
        cursor += 1;
    }
    if cursor == right_letters_start
        || cursor - right_letters_start > 3
        || bytes
            .get(cursor)
            .is_some_and(|byte| is_reference_identifier_byte(*byte))
    {
        return None;
    }
    Some((cursor, left, &formula[right_start..cursor]))
}

fn is_reference_identifier_byte(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'.')
}

fn parse_a1(address: &str) -> Option<(i32, i32)> {
    let split = address.find(|character: char| character.is_ascii_digit())?;
    let (column, row) = address.split_at(split);
    if column.is_empty() || row.is_empty() {
        return None;
    }
    let column = column.bytes().try_fold(0_i32, |value, byte| {
        byte.is_ascii_uppercase().then(|| {
            value
                .checked_mul(26)?
                .checked_add(i32::from(byte - b'A' + 1))
        })?
    })?;
    let row = row.parse::<i32>().ok()?;
    (row > 0 && column > 0).then_some((column, row))
}

fn parse_array_formula_range(reference: &str) -> Option<ArrayFormulaRange> {
    let (start, end) = reference.split_once(':').unwrap_or((reference, reference));
    let (start_column, start_row) = parse_a1(start)?;
    let (end_column, end_row) = parse_a1(end)?;
    (start_row <= end_row && start_column <= end_column).then_some(ArrayFormulaRange {
        start_row,
        start_column,
        end_row,
        end_column,
    })
}

fn to_one_based_i32(value: u32, label: &str) -> Result<i32, String> {
    i32::try_from(value)
        .map_err(|_| format!("{label} exceeds the evaluator range"))?
        .checked_add(1)
        .ok_or_else(|| format!("{label} exceeds the evaluator range"))
}

fn attribute<'a>(attributes: &'a [crate::xml::XmlAttribute<'a>], name: &str) -> Option<&'a str> {
    attributes
        .iter()
        .find(|attribute| local_name(attribute.name) == name)
        .map(|attribute| attribute.value)
}

fn formula_error(part: &str, message: &str) -> crate::diagnostic::Diagnostic {
    crate::diagnostic::Diagnostic::fatal(
        crate::diagnostic::DiagnosticCode::XmlInvalid,
        crate::diagnostic::Phase::Parse,
        None,
        message,
    )
    .in_part(part)
}

#[cfg(test)]
mod tests {
    use super::{
        CellInput, DefinedName, FormulaSheet, InputValue, WorkbookBounds,
        bound_whole_column_references, formula_sheet_order, has_dynamic_array_function,
        has_dynamic_reference_function, has_nondeterministic_function, import_cell_formula,
        node_may_reference_cells, reference_sheet_name, rewrite_address_defaults,
        rewrite_hyperlink_display, whole_column_range_at,
    };
    use ironcalc_base::expressions::{
        parser::new_parser_english,
        types::{CellReferenceIndex, CellReferenceRC},
    };
    use std::collections::{HashMap, HashSet};

    #[test]
    fn real_ooxml_conditional_formulas_are_evaluated() {
        let results = super::calculate(
            include_bytes!("../../tests/fixtures/ooxml-conditional-priority.xlsx"),
            crate::limits::Limits::default(),
        )
        .unwrap();
        assert_eq!(
            results.conditional_matches(0, 4),
            Some(&HashSet::from([(0, 0)]))
        );
        assert_eq!(
            results.conditional_matches(0, 1),
            Some(&HashSet::from([(0, 2)]))
        );
        let decoded =
            crate::calculation::CalculationResults::decode(&results.encode().unwrap(), 1000)
                .unwrap();
        assert_eq!(
            decoded.conditional_matches(0, 4),
            results.conditional_matches(0, 4)
        );
    }

    #[test]
    fn bounds_whole_columns_without_touching_strings_or_cell_ranges() {
        let bounds = WorkbookBounds::new(
            &[
                FormulaSheet {
                    name: "Sheet1".to_owned(),
                    part: None,
                },
                FormulaSheet {
                    name: "Data".to_owned(),
                    part: None,
                },
            ],
            vec![12, 321],
        );
        assert_eq!(
            bound_whole_column_references(
                r#"SUMIF(Data!R:R,A1,Data!$L:$L)+MATCH("A:A",A1:A9,0)"#,
                Some(0),
                &bounds,
            ),
            r#"SUMIF(Data!R$1:R$321,A1,Data!$L$1:$L$321)+MATCH("A:A",A1:A9,0)"#
        );
        assert_eq!(
            bound_whole_column_references("ROWS(A:A)+SUM(B:B)", Some(0), &bounds),
            "ROWS(A:A)+SUM(B$1:B$12)"
        );
        assert_eq!(
            bound_whole_column_references(
                r#"MATCH(A1,INDIRECT("'"&B1&"'!A:A"),0)"#,
                Some(0),
                &bounds,
            ),
            r#"MATCH(A1,INDIRECT("'"&B1&"'!A$1:A$321"),0)"#
        );
        assert!(whole_column_range_at("NAME_A:A", 5).is_none());
    }

    #[test]
    fn detects_direct_range_and_defined_name_dependencies() {
        let sheet_names = vec!["Contents".to_owned(), "Other".to_owned()];
        let defined_names = vec![("LinkLabel".to_owned(), None, "Contents!$D$8".to_owned())];
        let mut parser = new_parser_english(sheet_names.clone(), defined_names, HashMap::new());
        let context = CellReferenceRC {
            sheet: "Other".to_owned(),
            row: 1,
            column: 1,
        };
        let targets = HashSet::from([CellReferenceIndex {
            sheet: 0,
            row: 8,
            column: 4,
        }]);
        let target_sheets = HashSet::from([0]);
        for formula in ["Contents!D8", "SUM(Contents!D1:D10)", "LinkLabel"] {
            let node = parser.parse(formula, &context);
            assert!(
                node_may_reference_cells(
                    &node,
                    &mut parser,
                    &context,
                    &targets,
                    &target_sheets,
                    &sheet_names,
                    16,
                ),
                "{formula}"
            );
        }
        let node = parser.parse("Contents!E8", &context);
        assert!(!node_may_reference_cells(
            &node,
            &mut parser,
            &context,
            &targets,
            &target_sheets,
            &sheet_names,
            16,
        ));
        assert!(has_dynamic_reference_function("INDIRECT(A1)"));
        assert!(has_dynamic_reference_function("OFFSET(A1,1,0)"));
    }

    #[test]
    fn orders_array_formula_dependencies_before_dependents() {
        let sheets = [
            FormulaSheet {
                name: "Tools".to_owned(),
                part: None,
            },
            FormulaSheet {
                name: "Matrix".to_owned(),
                part: None,
            },
            FormulaSheet {
                name: "Data".to_owned(),
                part: None,
            },
        ];
        let inputs = [
            CellInput {
                sheet: 0,
                row: 1,
                column: 1,
                value: InputValue::Formula {
                    formula: "MMULT(MatrixInverse,A1:A3)".to_owned(),
                    shared: None,
                    array_range: Some(super::ArrayFormulaRange {
                        start_row: 1,
                        start_column: 1,
                        end_row: 3,
                        end_column: 1,
                    }),
                },
            },
            CellInput {
                sheet: 1,
                row: 1,
                column: 1,
                value: InputValue::Formula {
                    formula: "MINVERSE(Data!A1:C3)".to_owned(),
                    shared: None,
                    array_range: Some(super::ArrayFormulaRange {
                        start_row: 1,
                        start_column: 1,
                        end_row: 3,
                        end_column: 3,
                    }),
                },
            },
        ];
        let names = [DefinedName {
            name: "MatrixInverse".to_owned(),
            scope: None,
            formula: "Matrix!$A$1:$C$3".to_owned(),
        }];
        assert_eq!(formula_sheet_order(&sheets, &inputs, &names), [1, 0, 2]);
    }

    #[test]
    fn detects_only_explicit_dynamic_array_functions() {
        assert!(has_dynamic_array_function("_xlfn._xlws.FILTER(A:A,B:B=1)"));
        assert!(has_dynamic_array_function("UNIQUE(A1:A9)"));
        assert!(!has_dynamic_array_function("SUMIF(A:A,B1,C:C)"));
        assert!(!has_dynamic_array_function("MYFILTER(A1:A9)"));
    }

    #[test]
    fn refuses_random_but_recalculates_time_dependent_formula_results() {
        assert!(!has_nondeterministic_function("NOW()"));
        assert!(!has_nondeterministic_function("IF(A1,TODAY(),1)"));
        assert!(has_nondeterministic_function("IF(A1,TODAY(),RAND())"));
        assert!(!has_nondeterministic_function("SNOW(A1)+RANDOM_VALUE"));
    }

    #[test]
    fn uses_the_authored_hyperlink_display_expression() {
        assert_eq!(
            rewrite_hyperlink_display(
                r##"IFERROR(HYPERLINK("#"&ADDRESS(2,3),INDEX($R:$R,MATCH(R3,$A:$A,0))),"")"##
            ),
            r#"IFERROR(INDEX($R:$R,MATCH(R3,$A:$A,0)),"")"#
        );
    }

    #[test]
    fn supplies_excel_address_defaults_for_omitted_arguments() {
        assert_eq!(
            rewrite_address_defaults("ADDRESS(ROW(B1),COLUMN(),,,$B$85)"),
            "ADDRESS(ROW(B1),COLUMN(),1,TRUE,$B$85)"
        );
        assert_eq!(
            rewrite_address_defaults("ADDRESS(1,2,,FALSE)"),
            "ADDRESS(1,2,1,FALSE)"
        );
    }

    #[test]
    fn imports_quoted_cross_sheet_references() {
        let mut parser = new_parser_english(
            vec![
                "C_Emissions&Energy".to_owned(),
                "Summary_Communication".to_owned(),
            ],
            Vec::new(),
            HashMap::new(),
        );
        assert_eq!(
            import_cell_formula(
                &mut parser,
                "'C_Emissions&Energy'!H32",
                "Summary_Communication",
                13,
                27,
                false,
                &[
                    "C_Emissions&Energy".to_owned(),
                    "Summary_Communication".to_owned()
                ],
            ),
            "='C_Emissions&Energy'!H32"
        );
        assert_eq!(
            reference_sheet_name("'Manager''s Summary'!A:A", 21).as_deref(),
            Some("Manager's Summary")
        );
    }
}
