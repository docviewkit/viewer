//! OfficeViewer's dependency-minimal WebAssembly core and handwritten ABI.

pub mod calculation;
pub mod deflate;
pub mod diagnostic;
pub mod font_metrics;
pub mod format;
pub mod limits;
pub mod model;
pub mod package;
pub mod protocol;
#[cfg(any(
    feature = "native-formats",
    feature = "odf-formats",
    feature = "legacy-office-formats",
    feature = "iwork-formats"
))]
mod text_layout;
pub mod xml;
pub mod zip;

use std::alloc::{Layout, alloc_zeroed, dealloc};
#[cfg(feature = "native-formats")]
use std::collections::BTreeMap;
#[cfg(any(
    feature = "native-formats",
    feature = "pdf-formats",
    feature = "xps-formats",
    feature = "ofd-formats"
))]
use std::collections::BTreeSet;
#[cfg(feature = "xps-formats")]
use std::collections::HashSet;
use std::ops::Deref;
use std::sync::{Arc, Mutex};

use diagnostic::{Diagnostic, DiagnosticCode, Phase};
use limits::Limits;
use model::Document;

const ABI_VERSION: u32 = 11;
#[cfg(feature = "calculation-service")]
const CALC_ABI_VERSION: u32 = 3;
const UNIT_ALREADY_LOADED: u32 = u32::MAX;
#[cfg(feature = "native-formats")]
const UNIT_LOAD_ALL_REQUIRED: u32 = u32::MAX - 1;
#[cfg(feature = "native-formats")]
const UNIT_REPLACE_SNAPSHOT_REQUIRED: u32 = u32::MAX - 2;
const LIMIT_WORDS: usize = 12;
const ABI_ALIGNMENT: usize = 8;

struct StoredDocument {
    document: Document,
    #[cfg(feature = "pdf-formats")]
    lazy_pdf: Option<LazyPdf>,
    #[cfg(feature = "xps-formats")]
    lazy_xps: Option<LazyXps>,
    #[cfg(feature = "ofd-formats")]
    lazy_ofd: Option<LazyOfd>,
    #[cfg(feature = "native-formats")]
    lazy_xlsx: Option<LazyXlsx>,
    #[cfg(feature = "native-formats")]
    lazy_pptx: Option<LazyPptx>,
}

#[cfg(feature = "pdf-formats")]
struct LazyPdf {
    prepared: format::pdf::PreparedPdf,
    limits: Limits,
    loaded_units: BTreeSet<u32>,
    emitted_font_ids: BTreeSet<format::pdf::ObjectId>,
    emitted_font_bytes: usize,
}

#[cfg(feature = "xps-formats")]
struct LazyXps {
    prepared: format::xps::PreparedXps,
    limits: Limits,
    loaded_units: BTreeSet<u32>,
    emitted_fonts: HashSet<String>,
    emitted_font_bytes: usize,
}

#[cfg(feature = "ofd-formats")]
struct LazyOfd {
    prepared: format::ofd::PreparedOfd,
    limits: Limits,
    loaded_units: BTreeSet<u32>,
}

#[cfg(feature = "native-formats")]
struct LazyXlsx {
    prepared: format::PreparedXlsx,
    loaded_units: BTreeSet<u32>,
    partial_units: BTreeMap<u32, (i32, i32)>,
}

#[cfg(feature = "native-formats")]
struct LazyPptx {
    prepared: format::PreparedPptx,
    loaded_units: BTreeSet<u32>,
    materialized_image_bytes: usize,
}

#[cfg(any(
    feature = "native-formats",
    feature = "xps-formats",
    feature = "ofd-formats"
))]
fn rebase_objects(objects: &mut [model::Object], base: u32, preserve_stable_ids: bool) -> bool {
    for object in objects {
        let Some(numeric_id) = object.numeric_id.checked_add(base) else {
            return false;
        };
        object.numeric_id = numeric_id;
        if !preserve_stable_ids {
            object.stable_id = format!("object:{numeric_id}");
        }
        object.parent_numeric_id = match object.parent_numeric_id {
            Some(parent) => match parent.checked_add(base) {
                Some(parent) => Some(parent),
                None => return false,
            },
            None => None,
        };
        if !preserve_stable_ids {
            object.parent_stable_id = object
                .parent_numeric_id
                .map(|parent| format!("object:{parent}"));
        }
    }
    true
}

#[cfg(feature = "native-formats")]
fn replace_xlsx_unit(
    document: &mut Document,
    mut parsed: Document,
    unit_index: u32,
    object_limit: usize,
) -> bool {
    let retained_count = document
        .objects
        .iter()
        .filter(|object| object.unit_index != unit_index)
        .count();
    if retained_count
        .checked_add(parsed.objects.len())
        .is_none_or(|count| count > object_limit)
    {
        return false;
    }
    let base = match document
        .objects
        .iter()
        .filter(|object| object.unit_index != unit_index)
        .map(|object| object.numeric_id)
        .max()
    {
        Some(numeric_id) => match numeric_id.checked_add(1) {
            Some(base) => base,
            None => return false,
        },
        None => 0,
    };
    if !rebase_objects(&mut parsed.objects, base, false) {
        return false;
    }
    document
        .objects
        .retain(|object| object.unit_index != unit_index);
    document.objects.extend(parsed.objects);
    for diagnostic in parsed.diagnostics {
        if !document.diagnostics.contains(&diagnostic) {
            document.diagnostics.push(diagnostic);
        }
    }
    true
}

static DOCUMENTS: Mutex<Vec<Option<StoredDocument>>> = Mutex::new(Vec::new());
static RESULT: Mutex<Vec<u8>> = Mutex::new(Vec::new());
static ALLOCATIONS: Mutex<Vec<(usize, usize)>> = Mutex::new(Vec::new());

#[derive(Debug)]
pub(crate) struct RetainedInput {
    pointer: usize,
    length: usize,
}

// SAFETY: the allocation is immutable after `open_document` takes ownership.
unsafe impl Send for RetainedInput {}
// SAFETY: shared access only exposes immutable bytes.
unsafe impl Sync for RetainedInput {}

impl Deref for RetainedInput {
    type Target = [u8];

    fn deref(&self) -> &Self::Target {
        // SAFETY: ownership was removed from `ALLOCATIONS` and remains with self.
        unsafe { core::slice::from_raw_parts(self.pointer as *const u8, self.length) }
    }
}

impl Drop for RetainedInput {
    fn drop(&mut self) {
        let Ok(layout) = Layout::from_size_align(self.length.max(1), ABI_ALIGNMENT) else {
            return;
        };
        // SAFETY: this is the sole owner of the allocation taken from `ALLOCATIONS`.
        unsafe { dealloc(self.pointer as *mut u8, layout) };
    }
}

fn take_allocation(pointer: *mut u8, length: usize) -> Option<Arc<RetainedInput>> {
    let allocation_length = length.max(1);
    let mut allocations = ALLOCATIONS.lock().ok()?;
    let index = allocations
        .iter()
        .position(|(allocated_pointer, allocated_length)| {
            *allocated_pointer == pointer as usize && *allocated_length == allocation_length
        })?;
    allocations.swap_remove(index);
    Some(Arc::new(RetainedInput {
        pointer: pointer as usize,
        length,
    }))
}

#[unsafe(no_mangle)]
pub extern "C" fn ov_core_abi_version() -> u32 {
    ABI_VERSION
}

#[cfg(feature = "calculation-service")]
#[unsafe(no_mangle)]
pub extern "C" fn ov_calc_abi_version() -> u32 {
    CALC_ABI_VERSION
}

#[cfg(feature = "calculation-service")]
#[unsafe(no_mangle)]
/// Calculates missing spreadsheet formula results into the shared result buffer.
///
/// # Safety
///
/// Pointer ranges must refer to live allocations returned by [`ov_alloc`].
pub unsafe extern "C" fn ov_calculate(
    input_pointer: *const u8,
    input_length: u32,
    limits_pointer: *const u32,
    limits_count: u32,
    local_time_millis: f64,
) -> u32 {
    let input_length = input_length as usize;
    if !valid_allocation(input_pointer.cast_mut(), input_length.max(1), 1) {
        return 0;
    }
    let LimitsResult::Valid(limits) = read_limits(limits_pointer, limits_count) else {
        return 0;
    };
    if input_length > limits.max_input_bytes {
        return 0;
    }
    if !local_time_millis.is_finite()
        || local_time_millis < i64::MIN as f64
        || local_time_millis > i64::MAX as f64
    {
        return 0;
    }
    ironcalc_base::mock_time::set_mock_time(local_time_millis as i64);
    // SAFETY: the allocation registry proves this readable range is live.
    let input = unsafe { core::slice::from_raw_parts(input_pointer, input_length) };
    let Ok(results) = format::calculate_spreadsheet(input, limits) else {
        return 0;
    };
    let Ok(bytes) = results.encode() else {
        return 0;
    };
    let Ok(length) = u32::try_from(bytes.len()) else {
        return 0;
    };
    let Ok(mut result) = RESULT.lock() else {
        return 0;
    };
    *result = bytes;
    length
}

/// Allocates zeroed, eight-byte-aligned linear memory owned by the caller.
#[unsafe(no_mangle)]
pub extern "C" fn ov_alloc(length: u32) -> *mut u8 {
    let length = usize::try_from(length).unwrap_or(usize::MAX).max(1);
    if length > Limits::HARD_MAX.max_input_bytes {
        return core::ptr::null_mut();
    }
    let Ok(layout) = Layout::from_size_align(length, ABI_ALIGNMENT) else {
        return core::ptr::null_mut();
    };
    // SAFETY: `layout` is non-zero and valid.
    let pointer = unsafe { alloc_zeroed(layout) };
    if pointer.is_null() {
        return pointer;
    }
    let Ok(mut allocations) = ALLOCATIONS.lock() else {
        // SAFETY: this function just allocated `pointer` with `layout`.
        unsafe { dealloc(pointer, layout) };
        return core::ptr::null_mut();
    };
    if allocations.try_reserve(1).is_err() {
        // SAFETY: this function just allocated `pointer` with `layout`.
        unsafe { dealloc(pointer, layout) };
        return core::ptr::null_mut();
    }
    allocations.push((pointer as usize, length));
    pointer
}

#[unsafe(no_mangle)]
/// # Safety
///
/// `pointer` and `length` must exactly match a live allocation returned by
/// [`ov_alloc`]. Invalid or repeated frees are ignored defensively.
pub unsafe extern "C" fn ov_free(pointer: *mut u8, length: u32) {
    if pointer.is_null() {
        return;
    }
    let length = usize::try_from(length).unwrap_or(usize::MAX).max(1);
    let Ok(mut allocations) = ALLOCATIONS.lock() else {
        return;
    };
    let Some(index) = allocations
        .iter()
        .position(|(allocated_pointer, allocated_length)| {
            *allocated_pointer == pointer as usize && *allocated_length == length
        })
    else {
        return;
    };
    allocations.swap_remove(index);
    drop(allocations);
    let Ok(layout) = Layout::from_size_align(length, ABI_ALIGNMENT) else {
        return;
    };
    // SAFETY: the allocation registry proves ownership, exact layout, and that
    // this pointer has not already been freed.
    unsafe { dealloc(pointer, layout) };
}

/// Parses a complete document synchronously inside the document's Worker.
/// The returned non-zero handle owns the parsed model, including fatal diagnostics.
///
/// # Safety
///
/// All pointer ranges must refer to live allocations returned by [`ov_alloc`].
/// This call consumes the input allocation; limits and password remain caller-owned.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn ov_document_open(
    input_pointer: *const u8,
    input_length: u32,
    limits_pointer: *const u32,
    limits_count: u32,
    password_pointer: *const u8,
    password_length: u32,
    field_date: u32,
    field_time: u32,
) -> u32 {
    // SAFETY: this function forwards the ranges with the ownership documented above.
    unsafe {
        open_document(
            input_pointer,
            input_length,
            limits_pointer,
            limits_count,
            None,
            (password_pointer, password_length),
            decode_field_date_time(field_date, field_time),
        )
    }
}

/// Parses a document using a bounded table of browser-measured font advances.
///
/// # Safety
///
/// Every non-empty pointer range must refer to a live allocation returned by
/// [`ov_alloc`]. This call consumes input; all other ranges remain caller-owned.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn ov_document_open_with_font_metrics(
    input_pointer: *const u8,
    input_length: u32,
    limits_pointer: *const u32,
    limits_count: u32,
    metrics_pointer: *const u8,
    metrics_length: u32,
    password_pointer: *const u8,
    password_length: u32,
    field_date: u32,
    field_time: u32,
) -> u32 {
    // SAFETY: this function forwards the ranges with the ownership documented above.
    unsafe {
        open_document(
            input_pointer,
            input_length,
            limits_pointer,
            limits_count,
            Some((metrics_pointer, metrics_length)),
            (password_pointer, password_length),
            decode_field_date_time(field_date, field_time),
        )
    }
}

unsafe fn open_document(
    input_pointer: *const u8,
    input_length: u32,
    limits_pointer: *const u32,
    limits_count: u32,
    font_metrics: Option<(*const u8, u32)>,
    password: (*const u8, u32),
    field_date_time: Option<format::FieldDateTime>,
) -> u32 {
    let input_length = input_length as usize;
    let Some(input) = take_allocation(input_pointer.cast_mut(), input_length) else {
        return store_document(Document::fatal(invalid_input(
            "input pointer is not an owned ABI allocation",
        )));
    };
    let limits = read_limits(limits_pointer, limits_count);
    let limits = match limits {
        LimitsResult::Valid(limits) => limits,
        LimitsResult::Fatal(diagnostic) => return store_document(Document::fatal(diagnostic)),
    };
    if input_length > limits.max_input_bytes {
        return store_document(Document::fatal(Diagnostic::fatal(
            DiagnosticCode::InputTooLarge,
            Phase::Input,
            None,
            "input exceeds the configured byte limit",
        )));
    }
    let bytes: &[u8] = &input;
    let metrics = match font_metrics {
        Some((pointer, length)) if length != 0 => {
            let length = length as usize;
            if !valid_allocation(pointer.cast_mut(), length, 1) {
                return store_document(Document::fatal(invalid_input(
                    "font metric pointer is not an owned ABI allocation",
                )));
            }
            if length > limits.max_font_bytes {
                return store_document(Document::fatal(Diagnostic::fatal(
                    DiagnosticCode::FontBytesLimit,
                    Phase::Layout,
                    None,
                    "font metric table exceeds the configured font byte limit",
                )));
            }
            // SAFETY: the adapter allocated at least `length` bytes and keeps
            // the allocation alive for this synchronous call.
            let bytes = unsafe { core::slice::from_raw_parts(pointer, length) };
            match font_metrics::FontMetricTable::decode(bytes, limits) {
                Ok(metrics) => metrics,
                Err(diagnostic) => return store_document(Document::fatal(diagnostic)),
            }
        }
        _ => font_metrics::FontMetricTable::default(),
    };
    #[cfg(feature = "pdf-formats")]
    let password = match password {
        (pointer, _) if pointer.is_null() => None,
        (pointer, length) => {
            let length = length as usize;
            if !valid_allocation(pointer.cast_mut(), length.max(1), 1) {
                return store_document(Document::fatal(invalid_input(
                    "password pointer is not an owned ABI allocation",
                )));
            }
            // SAFETY: the adapter owns this allocation for the synchronous call.
            Some(unsafe { core::slice::from_raw_parts(pointer, length) })
        }
    };
    #[cfg(not(feature = "pdf-formats"))]
    let _ = password;
    // Accept iWork before OPC probes reject its unflagged UTF-8 paths.
    // Unrecognized or invalid inputs retain the existing dispatch and diagnostics.
    #[cfg(feature = "iwork-formats")]
    if let Ok(Some(document)) =
        format::iwork::detect_and_parse_with_font_metrics(bytes, limits, &metrics)
    {
        return store_document(document);
    }
    #[cfg(feature = "native-formats")]
    let mut lazy_xlsx = None;
    #[cfg(feature = "native-formats")]
    let mut lazy_pptx = None;
    #[cfg(feature = "pdf-formats")]
    let mut lazy_pdf = None;
    #[cfg(feature = "xps-formats")]
    let mut lazy_xps = None;
    #[cfg(feature = "ofd-formats")]
    let mut lazy_ofd = None;
    #[cfg(feature = "xps-formats")]
    let parsed_xps = match format::xps::prepare_retained(Arc::clone(&input), limits) {
        Ok(Some(prepared)) => match prepared.materialize(Some(0), &HashSet::new()) {
            Ok((document, emitted_fonts)) => {
                let emitted_font_bytes = document
                    .embedded_fonts
                    .iter()
                    .map(|font| font.bytes.len())
                    .sum();
                lazy_xps = Some(LazyXps {
                    prepared,
                    limits,
                    loaded_units: BTreeSet::from([0]),
                    emitted_fonts,
                    emitted_font_bytes,
                });
                Some(Ok(Some(document)))
            }
            Err(diagnostic) => Some(Err(diagnostic)),
        },
        Ok(None) => None,
        Err(diagnostic) => Some(Err(diagnostic)),
    };
    #[cfg(feature = "ofd-formats")]
    let parsed_ofd = match format::ofd::prepare_retained(Arc::clone(&input), limits) {
        Ok(Some(prepared)) => match prepared.materialize(Some(0)) {
            Ok(document) => {
                lazy_ofd = Some(LazyOfd {
                    prepared,
                    limits,
                    loaded_units: BTreeSet::from([0]),
                });
                Some(Ok(Some(document)))
            }
            Err(diagnostic) => Some(Err(diagnostic)),
        },
        Ok(None) => None,
        Err(diagnostic) => Some(Err(diagnostic)),
    };
    let parsed = (|| {
        // A prepared fixed-layout document already contains the first page.
        // Do not run the eager all-page fallback only to discard its result.
        #[cfg(feature = "ofd-formats")]
        if let Some(parsed) = parsed_ofd {
            return parsed;
        }
        #[cfg(feature = "xps-formats")]
        if let Some(parsed) = parsed_xps {
            return parsed;
        }
        #[cfg(feature = "native-formats")]
        let parsed_native =
            match format::prepare_native_retained(Arc::clone(&input), limits, &metrics) {
                Ok(Some(format::PreparedNative::Pptx {
                    document,
                    loaded_units,
                    materialized_image_bytes,
                    prepared,
                })) => {
                    lazy_pptx = Some(LazyPptx {
                        prepared,
                        loaded_units,
                        materialized_image_bytes,
                    });
                    Some(Ok(Some(document)))
                }
                Ok(Some(format::PreparedNative::Xlsx { document, prepared })) => {
                    lazy_xlsx = Some(LazyXlsx {
                        prepared,
                        loaded_units: BTreeSet::new(),
                        partial_units: BTreeMap::new(),
                    });
                    Some(Ok(Some(document)))
                }
                Ok(None) => None,
                Err(diagnostic) => Some(Err(diagnostic)),
            };
        #[cfg(all(feature = "native-formats", feature = "pdf-formats"))]
        let parsed = match parsed_native {
            Some(parsed) => parsed,
            None => match format::pdf::prepare(bytes, limits, password) {
                Ok(Some(prepared)) => {
                    match prepared.materialize(Some(0), true, &BTreeSet::new(), 0) {
                        Ok((document, emitted_font_ids)) => {
                            let emitted_font_bytes = document
                                .embedded_fonts
                                .iter()
                                .map(|font| font.bytes.len())
                                .sum();
                            lazy_pdf = Some(LazyPdf {
                                prepared,
                                limits,
                                loaded_units: BTreeSet::from([0]),
                                emitted_font_ids,
                                emitted_font_bytes,
                            });
                            Ok(Some(document))
                        }
                        Err(diagnostic) => Err(diagnostic),
                    }
                }
                Ok(None) => format::detect_and_parse_with_font_metrics_at(
                    bytes,
                    limits,
                    &metrics,
                    field_date_time,
                ),
                Err(diagnostic) => Err(diagnostic),
            },
        };
        #[cfg(all(feature = "native-formats", not(feature = "pdf-formats")))]
        let parsed = parsed_native.unwrap_or_else(|| {
            format::detect_and_parse_with_font_metrics_at(bytes, limits, &metrics, field_date_time)
        });
        #[cfg(all(not(feature = "native-formats"), feature = "pdf-formats"))]
        let parsed = match format::pdf::prepare(bytes, limits, password) {
            Ok(Some(prepared)) => match prepared.materialize(Some(0), true, &BTreeSet::new(), 0) {
                Ok((document, emitted_font_ids)) => {
                    let emitted_font_bytes = document
                        .embedded_fonts
                        .iter()
                        .map(|font| font.bytes.len())
                        .sum();
                    lazy_pdf = Some(LazyPdf {
                        prepared,
                        limits,
                        loaded_units: BTreeSet::from([0]),
                        emitted_font_ids,
                        emitted_font_bytes,
                    });
                    Ok(Some(document))
                }
                Err(diagnostic) => Err(diagnostic),
            },
            Ok(None) => format::detect_and_parse_with_font_metrics_at(
                bytes,
                limits,
                &metrics,
                field_date_time,
            ),
            Err(diagnostic) => Err(diagnostic),
        };
        #[cfg(all(not(feature = "native-formats"), not(feature = "pdf-formats")))]
        let parsed =
            format::detect_and_parse_with_font_metrics_at(bytes, limits, &metrics, field_date_time);
        parsed
    })();
    let document = match parsed {
        Ok(Some(document)) => document,
        Ok(None) => Document::fatal(Diagnostic::fatal(
            DiagnosticCode::UnsupportedFormat,
            Phase::Identify,
            None,
            "input is not a supported Office package or safe flat document",
        )),
        Err(diagnostic) => Document::fatal(diagnostic),
    };
    store_stored_document(StoredDocument {
        document,
        #[cfg(feature = "pdf-formats")]
        lazy_pdf,
        #[cfg(feature = "xps-formats")]
        lazy_xps,
        #[cfg(feature = "ofd-formats")]
        lazy_ofd,
        #[cfg(feature = "native-formats")]
        lazy_xlsx,
        #[cfg(feature = "native-formats")]
        lazy_pptx,
    })
}

fn decode_field_date_time(date: u32, time: u32) -> Option<format::FieldDateTime> {
    let year = date / 10_000;
    let month = date / 100 % 100;
    let day = date % 100;
    let hour = time / 100;
    let minute = time % 100;
    (year != 0 && (1..=12).contains(&month) && (1..=31).contains(&day) && hour < 24 && minute < 60)
        .then_some(format::FieldDateTime {
            year,
            month,
            day,
            hour,
            minute,
        })
}

/// Serializes a document into the versioned binary snapshot buffer.
#[unsafe(no_mangle)]
pub extern "C" fn ov_document_snapshot(handle: u32) -> u32 {
    let encoded = {
        let Ok(documents) = DOCUMENTS.lock() else {
            return 0;
        };
        let Some(Some(document)) = handle
            .checked_sub(1)
            .and_then(|index| documents.get(index as usize))
        else {
            return 0;
        };
        match protocol::encode(&document.document) {
            Ok(encoded) => encoded,
            Err(error) => {
                let fallback = Document::fatal(snapshot_error(error));
                let Ok(encoded) = protocol::encode(&fallback) else {
                    return 0;
                };
                encoded
            }
        }
    };
    let Ok(length) = u32::try_from(encoded.len()) else {
        return 0;
    };
    let Ok(mut result) = RESULT.lock() else {
        return 0;
    };
    *result = encoded;
    length
}

#[unsafe(no_mangle)]
pub extern "C" fn ov_result_pointer() -> *const u8 {
    let Ok(result) = RESULT.lock() else {
        return core::ptr::null();
    };
    if result.is_empty() {
        core::ptr::null()
    } else {
        result.as_ptr()
    }
}

#[unsafe(no_mangle)]
pub extern "C" fn ov_result_clear() {
    if let Ok(mut result) = RESULT.lock() {
        *result = Vec::new();
    }
}

#[unsafe(no_mangle)]
/// # Safety
///
/// `output_pointer` must refer to a live [`ov_alloc`] allocation containing at
/// least `capacity` `u32` slots and remain valid for the duration of this call.
pub unsafe extern "C" fn ov_document_hit_test(
    handle: u32,
    unit_index: u32,
    x: f32,
    y: f32,
    output_pointer: *mut u32,
    capacity: u32,
) -> u32 {
    if !x.is_finite()
        || !y.is_finite()
        || !valid_allocation(
            output_pointer.cast(),
            (capacity as usize).saturating_mul(4),
            4,
        )
    {
        return 0;
    }
    let capacity = capacity.min(256) as usize;
    let Ok(documents) = DOCUMENTS.lock() else {
        return 0;
    };
    let Some(Some(document)) = handle
        .checked_sub(1)
        .and_then(|index| documents.get(index as usize))
    else {
        return 0;
    };
    let hits = document.document.hit_test(unit_index, x, y, capacity);
    if capacity != 0 {
        // SAFETY: the adapter allocated `capacity * 4` bytes and the hit list is
        // capped to that capacity before copying.
        unsafe { core::ptr::copy_nonoverlapping(hits.as_ptr(), output_pointer, hits.len()) };
    }
    hits.len() as u32
}

/// Materializes one lazily parsed PDF page. Returns the encoded delta snapshot
/// length, [`UNIT_ALREADY_LOADED`] when no update is needed, or zero on error.
#[unsafe(no_mangle)]
pub extern "C" fn ov_document_load_unit(handle: u32, unit_index: u32) -> u32 {
    let Ok(mut documents) = DOCUMENTS.lock() else {
        return 0;
    };
    let Some(Some(stored)) = handle
        .checked_sub(1)
        .and_then(|index| documents.get_mut(index as usize))
    else {
        return 0;
    };
    #[cfg(not(any(
        feature = "native-formats",
        feature = "pdf-formats",
        feature = "xps-formats",
        feature = "ofd-formats"
    )))]
    let _ = (stored, unit_index);
    #[cfg(feature = "native-formats")]
    if let Some(lazy) = stored.lazy_pptx.as_mut() {
        if lazy.loaded_units.contains(&unit_index) {
            return UNIT_ALREADY_LOADED;
        }
        if unit_index as usize >= stored.document.units.len() {
            return 0;
        }
        let (mut parsed, materialized_image_bytes) = match lazy
            .prepared
            .materialize_unit_with_budget(unit_index, lazy.materialized_image_bytes)
        {
            Ok((document, materialized_image_bytes)) if !document.fatal => {
                (document, materialized_image_bytes)
            }
            _ => return 0,
        };
        let Some(next_count) = stored
            .document
            .objects
            .len()
            .checked_add(parsed.objects.len())
        else {
            return 0;
        };
        if next_count > lazy.prepared.limits().max_document_objects {
            return 0;
        }
        let Ok(base) = u32::try_from(stored.document.objects.len()) else {
            return 0;
        };
        if !rebase_objects(&mut parsed.objects, base, false) {
            return 0;
        }
        let encoded = match protocol::encode(&parsed) {
            Ok(encoded) => encoded,
            Err(_) => return 0,
        };
        let Ok(length) = u32::try_from(encoded.len()) else {
            return 0;
        };
        if length == 0 || length >= UNIT_LOAD_ALL_REQUIRED {
            return 0;
        }
        let Ok(mut result) = RESULT.lock() else {
            return 0;
        };
        *result = encoded;
        stored.document.objects.extend(parsed.objects);
        stored.document.diagnostics.extend(parsed.diagnostics);
        lazy.loaded_units.insert(unit_index);
        lazy.materialized_image_bytes = materialized_image_bytes;
        return length;
    }
    #[cfg(feature = "native-formats")]
    if let Some(lazy) = stored.lazy_xlsx.as_mut() {
        if lazy.loaded_units.contains(&unit_index) {
            return UNIT_ALREADY_LOADED;
        }
        if lazy.partial_units.contains_key(&unit_index) {
            let parsed = match lazy.prepared.materialize_unit(unit_index) {
                Ok(document) if !document.fatal => document,
                _ => return 0,
            };
            if !replace_xlsx_unit(
                &mut stored.document,
                parsed,
                unit_index,
                lazy.prepared.limits().max_document_objects,
            ) {
                return 0;
            }
            lazy.partial_units.remove(&unit_index);
            lazy.loaded_units.insert(unit_index);
            return UNIT_REPLACE_SNAPSHOT_REQUIRED;
        }
        if unit_index as usize >= stored.document.units.len() {
            return 0;
        }
        let mut parsed = match lazy.prepared.materialize_unit(unit_index) {
            Ok(document) if !document.fatal => document,
            _ => return 0,
        };
        let Some(next_count) = stored
            .document
            .objects
            .len()
            .checked_add(parsed.objects.len())
        else {
            return 0;
        };
        if next_count > lazy.prepared.limits().max_document_objects {
            return 0;
        }
        let Ok(base) = u32::try_from(stored.document.objects.len()) else {
            return 0;
        };
        if !rebase_objects(&mut parsed.objects, base, false) {
            return 0;
        }
        let encoded = match protocol::encode(&parsed) {
            Ok(encoded) => encoded,
            Err(_) => return 0,
        };
        let Ok(length) = u32::try_from(encoded.len()) else {
            return 0;
        };
        if length == 0 || length == UNIT_ALREADY_LOADED {
            return 0;
        }
        let Ok(mut result) = RESULT.lock() else {
            return 0;
        };
        *result = encoded;
        stored.document.objects.extend(parsed.objects);
        stored.document.diagnostics.extend(parsed.diagnostics);
        lazy.loaded_units.insert(unit_index);
        return length;
    }
    #[cfg(feature = "pdf-formats")]
    if let Some(lazy) = stored.lazy_pdf.as_mut() {
        if lazy.loaded_units.contains(&unit_index) {
            return UNIT_ALREADY_LOADED;
        }
        if unit_index as usize >= stored.document.units.len() {
            return 0;
        }
        let (mut parsed, emitted_font_ids) = match lazy.prepared.materialize(
            Some(unit_index as usize),
            true,
            &lazy.emitted_font_ids,
            lazy.emitted_font_bytes,
        ) {
            Ok((document, emitted_font_ids)) if !document.fatal => (document, emitted_font_ids),
            _ => return 0,
        };
        let Some(next_count) = stored
            .document
            .objects
            .len()
            .checked_add(parsed.objects.len())
        else {
            return 0;
        };
        if next_count > lazy.limits.max_document_objects {
            return 0;
        }
        let Ok(base) = u32::try_from(stored.document.objects.len()) else {
            return 0;
        };
        for object in &mut parsed.objects {
            let Some(numeric_id) = object.numeric_id.checked_add(base) else {
                return 0;
            };
            object.numeric_id = numeric_id;
            object.parent_numeric_id = match object.parent_numeric_id {
                Some(parent) => match parent.checked_add(base) {
                    Some(parent) => Some(parent),
                    None => return 0,
                },
                None => None,
            };
        }
        let encoded = match protocol::encode(&parsed) {
            Ok(encoded) => encoded,
            Err(_) => return 0,
        };
        let Ok(length) = u32::try_from(encoded.len()) else {
            return 0;
        };
        if length == 0 || length == UNIT_ALREADY_LOADED {
            return 0;
        }
        let Ok(mut result) = RESULT.lock() else {
            return 0;
        };
        let Some(next_font_bytes) = parsed
            .embedded_fonts
            .iter()
            .try_fold(lazy.emitted_font_bytes, |total, font| {
                total.checked_add(font.bytes.len())
            })
        else {
            return 0;
        };
        if next_font_bytes > lazy.limits.max_font_bytes {
            return 0;
        }
        *result = encoded;
        stored.document.embedded_fonts.extend(parsed.embedded_fonts);
        stored.document.objects.extend(parsed.objects);
        stored.document.diagnostics.extend(parsed.diagnostics);
        lazy.loaded_units.insert(unit_index);
        lazy.emitted_font_ids.extend(emitted_font_ids);
        lazy.emitted_font_bytes = next_font_bytes;
        return length;
    }
    #[cfg(feature = "xps-formats")]
    if let Some(lazy) = stored.lazy_xps.as_mut() {
        if lazy.loaded_units.contains(&unit_index) {
            return UNIT_ALREADY_LOADED;
        }
        if unit_index as usize >= stored.document.units.len() {
            return 0;
        }
        let (mut parsed, emitted_fonts) = match lazy
            .prepared
            .materialize(Some(unit_index as usize), &lazy.emitted_fonts)
        {
            Ok((document, emitted_fonts)) if !document.fatal => (document, emitted_fonts),
            _ => return 0,
        };
        let Some(next_count) = stored
            .document
            .objects
            .len()
            .checked_add(parsed.objects.len())
        else {
            return 0;
        };
        if next_count > lazy.limits.max_document_objects {
            return 0;
        }
        let Ok(base) = u32::try_from(stored.document.objects.len()) else {
            return 0;
        };
        if !rebase_objects(&mut parsed.objects, base, false) {
            return 0;
        }
        let encoded = match protocol::encode(&parsed) {
            Ok(encoded) => encoded,
            Err(_) => return 0,
        };
        let Ok(length) = u32::try_from(encoded.len()) else {
            return 0;
        };
        if length == 0 || length == UNIT_ALREADY_LOADED {
            return 0;
        }
        let Some(next_font_bytes) = parsed
            .embedded_fonts
            .iter()
            .try_fold(lazy.emitted_font_bytes, |total, font| {
                total.checked_add(font.bytes.len())
            })
        else {
            return 0;
        };
        if next_font_bytes > lazy.limits.max_font_bytes {
            return 0;
        }
        let Ok(mut result) = RESULT.lock() else {
            return 0;
        };
        *result = encoded;
        stored.document.embedded_fonts.extend(parsed.embedded_fonts);
        stored.document.objects.extend(parsed.objects);
        stored.document.diagnostics.extend(parsed.diagnostics);
        lazy.loaded_units.insert(unit_index);
        lazy.emitted_fonts = emitted_fonts;
        lazy.emitted_font_bytes = next_font_bytes;
        return length;
    }
    #[cfg(feature = "ofd-formats")]
    if let Some(lazy) = stored.lazy_ofd.as_mut() {
        if lazy.loaded_units.contains(&unit_index) {
            return UNIT_ALREADY_LOADED;
        }
        if unit_index as usize >= stored.document.units.len() {
            return 0;
        }
        let mut parsed = match lazy.prepared.materialize(Some(unit_index as usize)) {
            Ok(document) if !document.fatal => document,
            _ => return 0,
        };
        let Some(next_count) = stored
            .document
            .objects
            .len()
            .checked_add(parsed.objects.len())
        else {
            return 0;
        };
        if next_count > lazy.limits.max_document_objects {
            return 0;
        }
        let Ok(base) = u32::try_from(stored.document.objects.len()) else {
            return 0;
        };
        if !rebase_objects(&mut parsed.objects, base, true) {
            return 0;
        }
        let encoded = match protocol::encode(&parsed) {
            Ok(encoded) => encoded,
            Err(_) => return 0,
        };
        let Ok(length) = u32::try_from(encoded.len()) else {
            return 0;
        };
        if length == 0 || length == UNIT_ALREADY_LOADED {
            return 0;
        }
        let Ok(mut result) = RESULT.lock() else {
            return 0;
        };
        *result = encoded;
        stored.document.objects.extend(parsed.objects);
        stored.document.diagnostics.extend(parsed.diagnostics);
        lazy.loaded_units.insert(unit_index);
        return length;
    }
    UNIT_ALREADY_LOADED
}

/// Materializes the formula dependency closure needed by the initial visible
/// rectangle of one XLSX worksheet. A later request outside that rectangle, or
/// any full-unit query, upgrades the document to its exact complete model.
#[unsafe(no_mangle)]
pub extern "C" fn ov_document_load_unit_region(
    handle: u32,
    unit_index: u32,
    max_row: i32,
    max_column: i32,
) -> u32 {
    if max_row < 1 || max_column < 1 {
        return 0;
    }
    let Ok(mut documents) = DOCUMENTS.lock() else {
        return 0;
    };
    let Some(Some(stored)) = handle
        .checked_sub(1)
        .and_then(|index| documents.get_mut(index as usize))
    else {
        return 0;
    };
    #[cfg(not(feature = "native-formats"))]
    let _ = (stored, unit_index, max_row, max_column);
    #[cfg(feature = "native-formats")]
    if let Some(lazy) = stored.lazy_xlsx.as_mut() {
        if lazy.loaded_units.contains(&unit_index) {
            return UNIT_ALREADY_LOADED;
        }
        if let Some(&(loaded_row, loaded_column)) = lazy.partial_units.get(&unit_index) {
            if max_row <= loaded_row && max_column <= loaded_column {
                return UNIT_ALREADY_LOADED;
            }
            let next_row = max_row.max(loaded_row);
            let next_column = max_column.max(loaded_column);
            let parsed =
                match lazy
                    .prepared
                    .materialize_unit_region(unit_index, next_row, next_column)
                {
                    Ok(document) if !document.fatal => document,
                    _ => return 0,
                };
            if !replace_xlsx_unit(
                &mut stored.document,
                parsed,
                unit_index,
                lazy.prepared.limits().max_document_objects,
            ) {
                return 0;
            }
            lazy.partial_units
                .insert(unit_index, (next_row, next_column));
            return UNIT_REPLACE_SNAPSHOT_REQUIRED;
        }
        if unit_index as usize >= stored.document.units.len() {
            return 0;
        }
        let mut parsed = match lazy
            .prepared
            .materialize_unit_region(unit_index, max_row, max_column)
        {
            Ok(document) if !document.fatal => document,
            _ => return 0,
        };
        let Some(next_count) = stored
            .document
            .objects
            .len()
            .checked_add(parsed.objects.len())
        else {
            return 0;
        };
        if next_count > lazy.prepared.limits().max_document_objects {
            return 0;
        }
        let Ok(base) = u32::try_from(stored.document.objects.len()) else {
            return 0;
        };
        if !rebase_objects(&mut parsed.objects, base, false) {
            return 0;
        }
        let encoded = match protocol::encode(&parsed) {
            Ok(encoded) => encoded,
            Err(_) => return 0,
        };
        let Ok(length) = u32::try_from(encoded.len()) else {
            return 0;
        };
        if length == 0 || length >= UNIT_LOAD_ALL_REQUIRED {
            return 0;
        }
        let Ok(mut result) = RESULT.lock() else {
            return 0;
        };
        *result = encoded;
        stored.document.objects.extend(parsed.objects);
        stored.document.diagnostics.extend(parsed.diagnostics);
        lazy.partial_units.insert(unit_index, (max_row, max_column));
        return length;
    }
    UNIT_ALREADY_LOADED
}

/// Materializes every remaining page for document-wide queries.
#[unsafe(no_mangle)]
pub extern "C" fn ov_document_load_all(handle: u32) -> u32 {
    let Ok(mut documents) = DOCUMENTS.lock() else {
        return 0;
    };
    let Some(Some(stored)) = handle
        .checked_sub(1)
        .and_then(|index| documents.get_mut(index as usize))
    else {
        return 0;
    };
    #[cfg(not(any(
        feature = "native-formats",
        feature = "pdf-formats",
        feature = "xps-formats",
        feature = "ofd-formats"
    )))]
    let _ = stored;
    #[cfg(feature = "native-formats")]
    if let Some(lazy) = stored.lazy_pptx.as_ref() {
        let parsed = match lazy.prepared.materialize() {
            Ok(document) if !document.fatal => document,
            _ => return 0,
        };
        stored.document = parsed;
        stored.lazy_pptx = None;
        return 1;
    }
    #[cfg(feature = "native-formats")]
    if let Some(lazy) = stored.lazy_xlsx.as_ref() {
        let parsed = match lazy.prepared.materialize() {
            Ok(document) if !document.fatal => document,
            _ => return 0,
        };
        stored.document = parsed;
        stored.lazy_xlsx = None;
        return 1;
    }
    #[cfg(feature = "pdf-formats")]
    if let Some(lazy) = stored.lazy_pdf.as_ref() {
        let parsed = match lazy.prepared.materialize(None, true, &BTreeSet::new(), 0) {
            Ok((document, _)) if !document.fatal => document,
            _ => return 0,
        };
        stored.document = parsed;
        stored.lazy_pdf = None;
        return 1;
    }
    #[cfg(feature = "xps-formats")]
    if let Some(lazy) = stored.lazy_xps.as_ref() {
        let parsed = match lazy.prepared.materialize(None, &HashSet::new()) {
            Ok((document, _)) if !document.fatal => document,
            _ => return 0,
        };
        stored.document = parsed;
        stored.lazy_xps = None;
        return 1;
    }
    #[cfg(feature = "ofd-formats")]
    if let Some(lazy) = stored.lazy_ofd.as_ref() {
        let parsed = match lazy.prepared.materialize(None) {
            Ok(document) if !document.fatal => document,
            _ => return 0,
        };
        stored.document = parsed;
        stored.lazy_ofd = None;
        return 1;
    }
    2
}

#[unsafe(no_mangle)]
pub extern "C" fn ov_document_requires_calculation(handle: u32) -> u32 {
    #[cfg(feature = "native-formats")]
    {
        let Ok(documents) = DOCUMENTS.lock() else {
            return 0;
        };
        let Some(Some(stored)) = handle
            .checked_sub(1)
            .and_then(|index| documents.get(index as usize))
        else {
            return 0;
        };
        return u32::from(
            stored
                .lazy_xlsx
                .as_ref()
                .is_some_and(|lazy| lazy.prepared.requires_calculation()),
        );
    }
    #[cfg(not(feature = "native-formats"))]
    {
        let _ = handle;
        0
    }
}

#[unsafe(no_mangle)]
/// Applies opaque results produced by the optional calculation Wasm.
///
/// # Safety
///
/// The result range must refer to a live allocation returned by [`ov_alloc`].
pub unsafe extern "C" fn ov_document_apply_calculation(
    handle: u32,
    result_pointer: *const u8,
    result_length: u32,
) -> u32 {
    #[cfg(not(feature = "native-formats"))]
    {
        let _ = (handle, result_pointer, result_length);
        return 0;
    }
    #[cfg(feature = "native-formats")]
    {
        let result_length = result_length as usize;
        if !valid_allocation(result_pointer.cast_mut(), result_length.max(1), 1) {
            return 0;
        }
        let Ok(mut documents) = DOCUMENTS.lock() else {
            return 0;
        };
        let Some(Some(stored)) = handle
            .checked_sub(1)
            .and_then(|index| documents.get_mut(index as usize))
        else {
            return 0;
        };
        if let Some(lazy) = stored.lazy_xlsx.as_mut() {
            // SAFETY: the allocation registry proves this readable range is live.
            let bytes = unsafe { core::slice::from_raw_parts(result_pointer, result_length) };
            let Ok(calculation) = calculation::CalculationResults::decode(
                bytes,
                lazy.prepared.limits().max_document_objects,
            ) else {
                return 0;
            };
            lazy.prepared.apply_calculation(calculation);
            stored
                .document
                .diagnostics
                .retain(|diagnostic| diagnostic.message != calculation::REQUIRED_MESSAGE);
            return 1;
        }
        0
    }
}

#[unsafe(no_mangle)]
pub extern "C" fn ov_document_close(handle: u32) {
    let Ok(mut documents) = DOCUMENTS.lock() else {
        return;
    };
    let Some(index) = handle.checked_sub(1).map(|index| index as usize) else {
        return;
    };
    if let Some(slot) = documents.get_mut(index) {
        *slot = None;
    }
    while documents.last().is_some_and(Option::is_none) {
        documents.pop();
    }
}

enum LimitsResult {
    Valid(Limits),
    Fatal(Diagnostic),
}

fn read_limits(pointer: *const u32, count: u32) -> LimitsResult {
    if count as usize != LIMIT_WORDS
        || !valid_allocation(pointer.cast::<u8>().cast_mut(), LIMIT_WORDS * 4, 4)
    {
        return LimitsResult::Fatal(Diagnostic::fatal(
            DiagnosticCode::InvalidLimits,
            Phase::Input,
            None,
            "resource limit block has the wrong size",
        ));
    }
    // SAFETY: the adapter allocates an aligned block containing LIMIT_WORDS u32 values.
    let words = unsafe { core::slice::from_raw_parts(pointer, LIMIT_WORDS) };
    let limits = Limits {
        max_input_bytes: words[0] as usize,
        max_zip_entries: words[1] as usize,
        max_entry_uncompressed_bytes: words[2] as usize,
        max_total_uncompressed_bytes: words[3] as usize,
        max_compression_ratio: words[4],
        max_xml_bytes: words[5] as usize,
        max_xml_depth: words[6] as usize,
        max_xml_nodes: words[7] as usize,
        max_relationship_edges: words[8] as usize,
        max_document_objects: words[9] as usize,
        max_render_pixels: words[10] as usize,
        max_font_bytes: words[11] as usize,
        ..Limits::default()
    };
    if let Err(error) = limits.validate() {
        return LimitsResult::Fatal(Diagnostic::fatal(
            DiagnosticCode::InvalidLimits,
            Phase::Input,
            None,
            error.to_string(),
        ));
    }
    LimitsResult::Valid(limits)
}

fn valid_allocation(pointer: *mut u8, required_length: usize, alignment: usize) -> bool {
    if pointer.is_null() || !(pointer as usize).is_multiple_of(alignment) {
        return false;
    }
    let Ok(allocations) = ALLOCATIONS.lock() else {
        return false;
    };
    allocations
        .iter()
        .any(|(allocated_pointer, allocated_length)| {
            *allocated_pointer == pointer as usize && required_length <= *allocated_length
        })
}

fn invalid_input(message: &str) -> Diagnostic {
    Diagnostic::fatal(DiagnosticCode::FormatInvalid, Phase::Input, None, message)
}

fn snapshot_error(error: protocol::ProtocolError) -> Diagnostic {
    match error {
        protocol::ProtocolError::AllocationFailed => Diagnostic::fatal(
            DiagnosticCode::AllocationFailed,
            Phase::Parse,
            None,
            "memory allocation failed while preparing the document snapshot",
        ),
        protocol::ProtocolError::SnapshotTooLarge => Diagnostic::fatal(
            DiagnosticCode::ObjectLimit,
            Phase::Parse,
            None,
            "document snapshot exceeds the 256 MiB transfer limit",
        ),
        protocol::ProtocolError::FieldTooLong | protocol::ProtocolError::TooManyFields => {
            Diagnostic::fatal(
                DiagnosticCode::FormatInvalid,
                Phase::Parse,
                None,
                "document metadata exceeds the binary protocol field limit",
            )
        }
        protocol::ProtocolError::InvalidValue => Diagnostic::fatal(
            DiagnosticCode::FormatInvalid,
            Phase::Parse,
            None,
            "document contains a non-finite or out-of-range rendering value",
        ),
    }
}

fn store_document(document: Document) -> u32 {
    store_stored_document(StoredDocument {
        document,
        #[cfg(feature = "pdf-formats")]
        lazy_pdf: None,
        #[cfg(feature = "xps-formats")]
        lazy_xps: None,
        #[cfg(feature = "ofd-formats")]
        lazy_ofd: None,
        #[cfg(feature = "native-formats")]
        lazy_xlsx: None,
        #[cfg(feature = "native-formats")]
        lazy_pptx: None,
    })
}

fn store_stored_document(document: StoredDocument) -> u32 {
    let Ok(mut documents) = DOCUMENTS.lock() else {
        return 0;
    };
    if let Some((index, slot)) = documents
        .iter_mut()
        .enumerate()
        .find(|(_, slot)| slot.is_none())
    {
        *slot = Some(document);
        return u32::try_from(index + 1).unwrap_or(0);
    }
    if documents.len() >= u32::MAX as usize {
        return 0;
    }
    documents.push(Some(document));
    documents.len() as u32
}

#[cfg(test)]
mod tests {
    use super::{ov_alloc, ov_core_abi_version, ov_free, valid_allocation};
    #[cfg(feature = "pdf-formats")]
    use super::{ov_document_close, ov_document_load_unit, ov_document_open};
    #[cfg(feature = "pdf-formats")]
    use crate::limits::Limits;

    #[test]
    fn allocation_abi_is_aligned_and_round_trips() {
        let pointer = ov_alloc(17);
        assert!(!pointer.is_null());
        assert_eq!((pointer as usize) % 8, 0);
        // SAFETY: this test owns the 17-byte allocation until the matching free.
        unsafe { core::ptr::write_bytes(pointer, 0x5a, 17) };
        assert!(valid_allocation(pointer, 17, 8));
        // SAFETY: `pointer` is the exact allocation returned by `ov_alloc`.
        unsafe { ov_free(pointer, 17) };
        assert!(!valid_allocation(pointer, 17, 8));
        // A repeated or incorrectly sized free is ignored instead of reaching
        // the allocator with an invalid layout.
        // SAFETY: exercising the ABI's defensive repeated-free handling.
        unsafe { ov_free(pointer, 17) };
        assert_eq!(ov_core_abi_version(), super::ABI_VERSION);
    }

    #[cfg(feature = "pdf-formats")]
    fn open_pdf_for_test(pdf: &[u8]) -> u32 {
        let limits = Limits::default();
        let limit_words = [
            limits.max_input_bytes as u32,
            limits.max_zip_entries as u32,
            limits.max_entry_uncompressed_bytes as u32,
            limits.max_total_uncompressed_bytes as u32,
            limits.max_compression_ratio,
            limits.max_xml_bytes as u32,
            limits.max_xml_depth as u32,
            limits.max_xml_nodes as u32,
            limits.max_relationship_edges as u32,
            limits.max_document_objects as u32,
            limits.max_render_pixels as u32,
            limits.max_font_bytes as u32,
        ];
        let input = ov_alloc(pdf.len() as u32);
        let limit_bytes = (limit_words.len() * size_of::<u32>()) as u32;
        let limit_pointer = ov_alloc(limit_bytes).cast::<u32>();
        assert!(!input.is_null());
        assert!(!limit_pointer.is_null());
        // SAFETY: both destinations are the exact live ABI allocations above.
        unsafe {
            core::ptr::copy_nonoverlapping(pdf.as_ptr(), input, pdf.len());
            core::ptr::copy_nonoverlapping(limit_words.as_ptr(), limit_pointer, limit_words.len());
        }
        // SAFETY: the input and limit ranges are live ABI allocations.
        let handle = unsafe {
            ov_document_open(
                input,
                pdf.len() as u32,
                limit_pointer,
                limit_words.len() as u32,
                core::ptr::null(),
                0,
                0,
                0,
            )
        };
        assert_ne!(handle, 0);
        assert!(!valid_allocation(input, pdf.len(), 1));
        // SAFETY: `open` consumed input; the limits allocation remains caller-owned.
        unsafe { ov_free(limit_pointer.cast(), limit_bytes) };
        handle
    }

    #[cfg(all(feature = "iwork-formats", feature = "pdf-formats"))]
    #[test]
    fn iwork_abi_preserves_parser_snapshot() {
        let mut inputs =
            vec![include_bytes!("../tests/fixtures/keynote-first-line-indent.key").to_vec()];
        let mut unicode = inputs[0].clone();
        let original = b"Metadata/BuildVersionHistory.plist";
        let replacement = "Metadata/BuildVersion历史x.plist".as_bytes();
        assert_eq!(original.len(), replacement.len());
        let mut replacements = 0;
        while let Some(offset) = unicode
            .windows(original.len())
            .position(|name| name == original)
        {
            unicode[offset..offset + original.len()].copy_from_slice(replacement);
            replacements += 1;
        }
        assert_eq!(replacements, 2);
        inputs.push(unicode);
        if let Some(path) = std::env::var_os("IWORK_REGRESSION_INPUT") {
            inputs.insert(0, std::fs::read(path).unwrap());
        }
        for input in inputs {
            let expected = crate::format::iwork::detect_and_parse(&input, Limits::default())
                .unwrap()
                .unwrap();
            assert!(!expected.fatal);
            let handle = open_pdf_for_test(&input);
            let actual = {
                let documents = super::DOCUMENTS.lock().unwrap();
                crate::protocol::encode(&documents[handle as usize - 1].as_ref().unwrap().document)
                    .unwrap()
            };
            ov_document_close(handle);
            assert!(
                actual == crate::protocol::encode(&expected).unwrap(),
                "ABI changed the iWork scene snapshot"
            );
        }
    }

    #[cfg(feature = "pdf-formats")]
    #[test]
    fn lazy_pdf_load_ignores_zero_size_text_in_a_transparent_form() {
        let pdf = b"%PDF-1.7
1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj
2 0 obj << /Type /Pages /Kids [3 0 R 8 0 R] /Count 2 /MediaBox [0 0 100 100] >> endobj
3 0 obj << /Type /Page /Parent 2 0 R /Resources << /Font << /F0 6 0 R >> >> /Contents 7 0 R >> endobj
4 0 obj << /Type /ExtGState /CA 1 /ca 0 >> endobj
5 0 obj << /Type /XObject /Subtype /Form /FormType 1 /BBox [0 0 100 100] /Resources << /Font << /F0 6 0 R >> >> /Length 42 >> stream
BT /F0 0 Tf 1 0 0 1 0 0 Tm (hidden) Tj ET
endstream endobj
6 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj
7 0 obj << /Length 46 >> stream
BT /F0 12 Tf 1 0 0 1 10 80 Tm (Visible) Tj ET
endstream endobj
8 0 obj << /Type /Page /Parent 2 0 R /Resources << /ExtGState << /GS0 4 0 R >> /XObject << /Fm0 5 0 R >> /Font << /F0 6 0 R >> >> /Contents 9 0 R >> endobj
9 0 obj << /Length 66 >> stream
BT /F0 12 Tf 1 0 0 1 10 80 Tm (Visible) Tj ET
q /GS0 gs /Fm0 Do Q
endstream endobj
trailer << /Root 1 0 R >>
%%EOF";
        let handle = open_pdf_for_test(pdf);
        let delta = ov_document_load_unit(handle, 1);
        ov_document_close(handle);
        assert_ne!(delta, 0);
    }

    #[cfg(feature = "pdf-formats")]
    #[test]
    fn lazy_pdf_deltas_send_a_shared_embedded_font_once() {
        let cff = [
            1, 0, 4, 1, // header
            0, 1, 1, 1, 2, b'A', // Name INDEX
            0, 1, 1, 1, 7, // Top DICT INDEX header
            173, 15, // charset offset 34
            170, 16, // encoding offset 31
            176, 17, // CharStrings offset 37
            0, 1, 1, 1, 4, b'G', b'2', b'1', // String INDEX: SID 391 = G21
            0, 0, // Global Subr INDEX
            0, 1, 33, // format 0 encoding: character code 33 -> GID 1
            0, 1, 135, // format 0 charset: GID 1 -> SID 391
            0, 2, 1, 1, 2, 3, 14, 14, // two empty Type 2 charstrings
        ];
        let mut pdf = format!(
            "%PDF-1.7
1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj
2 0 obj << /Type /Pages /Kids [3 0 R 4 0 R 5 0 R] /Count 3 /MediaBox [0 0 100 100] >> endobj
3 0 obj << /Type /Page /Parent 2 0 R /Resources << /Font << /F0 9 0 R >> >> /Contents 6 0 R >> endobj
4 0 obj << /Type /Page /Parent 2 0 R /Resources << /Font << /F1 10 0 R >> >> /Contents 7 0 R >> endobj
5 0 obj << /Type /Page /Parent 2 0 R /Resources << /Font << /F1 10 0 R >> >> /Contents 8 0 R >> endobj
6 0 obj << >> stream
BT /F0 10 Tf 1 0 0 1 10 80 Tm (First) Tj ET
endstream endobj
7 0 obj << >> stream
BT /F1 10 Tf 1 0 0 1 10 80 Tm (!) Tj ET
endstream endobj
8 0 obj << >> stream
BT /F1 10 Tf 1 0 0 1 10 80 Tm (!) Tj ET
endstream endobj
9 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj
10 0 obj << /Type /Font /Subtype /Type1 /BaseFont /SharedSubset /Encoding << /Differences [33 /G21] >> /FirstChar 33 /LastChar 33 /Widths [600] /FontDescriptor 11 0 R >> endobj
11 0 obj << /Type /FontDescriptor /FontName /SharedSubset /Ascent 800 /Descent -200 /FontBBox [0 -200 1000 800] /FontFile3 12 0 R >> endobj
12 0 obj << /Subtype /Type1C /Length {} >> stream
",
            cff.len(),
        )
        .into_bytes();
        pdf.extend_from_slice(&cff);
        pdf.extend_from_slice(
            b"
endstream endobj
trailer << /Root 1 0 R >>
%%EOF",
        );

        let handle = open_pdf_for_test(&pdf);
        let first_shared_font_delta = ov_document_load_unit(handle, 1);
        let reused_font_delta = ov_document_load_unit(handle, 2);
        ov_document_close(handle);
        assert_ne!(first_shared_font_delta, 0);
        assert_ne!(reused_font_delta, 0);
        assert!(
            reused_font_delta < first_shared_font_delta,
            "the second page delta must reuse the font registered by the first page"
        );
    }
}
