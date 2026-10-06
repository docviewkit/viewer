//! Strict, bounded ZIP reader for OOXML and ODF packages.

use std::cell::Cell;
use std::collections::HashSet;

#[cfg(feature = "xps-formats")]
use crate::deflate::decompress_output_prefix;
use crate::deflate::{DeflateErrorKind, decompress_with_limit};
use crate::diagnostic::{Diagnostic, DiagnosticCode, Phase};
use crate::limits::Limits;

const LOCAL_FILE_HEADER: u32 = 0x0403_4b50;
const CENTRAL_DIRECTORY_HEADER: u32 = 0x0201_4b50;
const END_OF_CENTRAL_DIRECTORY: u32 = 0x0605_4b50;
const ZIP64_END_OF_CENTRAL_DIRECTORY: u32 = 0x0606_4b50;
const ZIP64_END_OF_CENTRAL_DIRECTORY_LOCATOR: u32 = 0x0706_4b50;
const DATA_DESCRIPTOR: u32 = 0x0807_4b50;

const FLAG_ENCRYPTED: u16 = 1 << 0;
const FLAG_DATA_DESCRIPTOR: u16 = 1 << 3;
const FLAG_UTF8: u16 = 1 << 11;
const ALLOWED_FLAGS: u16 = (1 << 1) | (1 << 2) | FLAG_DATA_DESCRIPTOR | FLAG_UTF8;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum CompressionMethod {
    Stored,
    Deflate,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ZipEntry {
    name: String,
    compression: CompressionMethod,
    crc32: u32,
    compressed_size: usize,
    uncompressed_size: usize,
    data_start: usize,
    is_directory: bool,
}

impl ZipEntry {
    pub fn name(&self) -> &str {
        &self.name
    }

    pub const fn compression(&self) -> CompressionMethod {
        self.compression
    }

    pub const fn compressed_size(&self) -> usize {
        self.compressed_size
    }

    pub const fn uncompressed_size(&self) -> usize {
        self.uncompressed_size
    }

    pub const fn is_directory(&self) -> bool {
        self.is_directory
    }
}

#[derive(Debug)]
pub struct ZipArchive<'a> {
    bytes: &'a [u8],
    entries: Vec<ZipEntry>,
    limits: Limits,
    iwork_expanded_bytes_left: Option<Cell<usize>>,
    deflate_operations_left: Cell<usize>,
}

impl<'a> ZipArchive<'a> {
    /// Parses and validates ZIP metadata without eagerly inflating entries.
    pub fn parse(bytes: &'a [u8], limits: Limits) -> Result<Self, Diagnostic> {
        Self::parse_with_iwork_compatibility(bytes, limits, false)
    }

    /// Uses iWork-specific path compatibility and expanded-byte accounting.
    /// ZIP64 container handling itself is shared with ordinary packages.
    #[cfg(any(feature = "iwork-formats", test))]
    pub(crate) fn parse_iwork(bytes: &'a [u8], limits: Limits) -> Result<Self, Diagnostic> {
        Self::parse_with_iwork_compatibility(bytes, limits, true)
    }

    fn parse_with_iwork_compatibility(
        bytes: &'a [u8],
        limits: Limits,
        iwork_compatibility: bool,
    ) -> Result<Self, Diagnostic> {
        limits.validate().map_err(|error| {
            Diagnostic::fatal(
                DiagnosticCode::InvalidLimits,
                Phase::Input,
                None,
                error.to_string(),
            )
        })?;
        if bytes.len() > limits.max_input_bytes {
            return Err(zip_error(
                DiagnosticCode::InputTooLarge,
                None,
                "ZIP input exceeds the configured byte limit",
            ));
        }

        let eocd = find_eocd(bytes)?;
        let directory = parse_directory_record(bytes, eocd)?;
        let entry_count = directory.entry_count;
        if entry_count > limits.max_zip_entries {
            return Err(zip_error(
                DiagnosticCode::ZipEntryLimit,
                Some(eocd),
                "ZIP contains too many entries",
            ));
        }
        let directory_offset = directory.offset;
        let directory_end = directory.end;

        let mut cursor = directory_offset;
        let mut total_uncompressed = 0_usize;
        let mut names = HashSet::new();
        let mut entries = Vec::new();
        entries.try_reserve_exact(entry_count).map_err(|_| {
            zip_error(
                DiagnosticCode::AllocationFailed,
                Some(directory_offset),
                "unable to allocate ZIP entry table",
            )
        })?;
        let mut local_ranges = Vec::new();
        local_ranges.try_reserve_exact(entry_count).map_err(|_| {
            zip_error(
                DiagnosticCode::AllocationFailed,
                Some(directory_offset),
                "unable to allocate ZIP range table",
            )
        })?;

        for _ in 0..entry_count {
            ensure_range(bytes, cursor, 46)?;
            if read_u32(bytes, cursor)? != CENTRAL_DIRECTORY_HEADER {
                return Err(zip_error(
                    DiagnosticCode::ZipInvalid,
                    Some(cursor),
                    "invalid central-directory entry signature",
                ));
            }

            let flags = read_u16(bytes, cursor + 8)?;
            validate_flags(flags, cursor + 8)?;
            let compression = compression_method(read_u16(bytes, cursor + 10)?, cursor + 10)?;
            if compression == CompressionMethod::Stored && flags & ((1 << 1) | (1 << 2)) != 0 {
                return Err(zip_error(
                    DiagnosticCode::ZipInvalid,
                    Some(cursor + 8),
                    "stored ZIP entry uses Deflate-only option flags",
                ));
            }
            let expected_crc = read_u32(bytes, cursor + 16)?;
            let compressed_size_u32 = read_u32(bytes, cursor + 20)?;
            let uncompressed_size_u32 = read_u32(bytes, cursor + 24)?;
            let name_length = usize::from(read_u16(bytes, cursor + 28)?);
            let extra_length = usize::from(read_u16(bytes, cursor + 30)?);
            let comment_length = usize::from(read_u16(bytes, cursor + 32)?);
            let start_disk = read_u16(bytes, cursor + 34)?;
            let local_offset_u32 = read_u32(bytes, cursor + 42)?;

            let variable_size = name_length
                .checked_add(extra_length)
                .and_then(|value| value.checked_add(comment_length))
                .ok_or_else(|| {
                    zip_error(
                        DiagnosticCode::ZipInvalid,
                        Some(cursor),
                        "central-directory field lengths overflow",
                    )
                })?;
            let next_cursor = checked_add(cursor, 46 + variable_size, cursor)?;
            if next_cursor > directory_end {
                return Err(zip_error(
                    DiagnosticCode::ZipInvalid,
                    Some(cursor),
                    "central-directory entry is truncated",
                ));
            }

            let raw_name = &bytes[cursor + 46..cursor + 46 + name_length];
            let extra_start = cursor + 46 + name_length;
            let zip64 = parse_central_extra(
                &bytes[extra_start..extra_start + extra_length],
                extra_start,
                uncompressed_size_u32 == u32::MAX,
                compressed_size_u32 == u32::MAX,
                local_offset_u32 == u32::MAX,
                start_disk == u16::MAX,
            )?;
            let compressed_size_u64 = resolve_zip64_u32(
                compressed_size_u32,
                zip64.compressed_size,
                cursor + 20,
                "compressed size",
            )?;
            let uncompressed_size_u64 = resolve_zip64_u32(
                uncompressed_size_u32,
                zip64.uncompressed_size,
                cursor + 24,
                "uncompressed size",
            )?;
            let local_offset_u64 = resolve_zip64_u32(
                local_offset_u32,
                zip64.local_offset,
                cursor + 42,
                "local-header offset",
            )?;
            let start_disk = if start_disk == u16::MAX {
                zip64.start_disk.ok_or_else(|| {
                    zip_error(
                        DiagnosticCode::ZipInvalid,
                        Some(extra_start),
                        "ZIP64 start-disk field is missing",
                    )
                })?
            } else {
                u32::from(start_disk)
            };
            if start_disk != 0 {
                return Err(zip_error(
                    DiagnosticCode::ZipMultiDiskForbidden,
                    Some(cursor + 34),
                    "entry points to another ZIP disk",
                ));
            }
            let (name, is_directory) = validate_path(
                raw_name,
                flags,
                limits.max_zip_path_bytes,
                cursor + 46,
                iwork_compatibility,
            )?;
            if !names.insert(name.clone()) {
                return Err(zip_error(
                    DiagnosticCode::ZipDuplicateEntry,
                    Some(cursor + 46),
                    "duplicate ZIP entry name",
                )
                .in_part(name));
            }

            let compressed_size = size_with_limit(
                compressed_size_u64,
                bytes.len(),
                DiagnosticCode::ZipInvalid,
                cursor + 20,
                "compressed size exceeds the ZIP input",
            )?;
            let uncompressed_size = size_with_limit(
                uncompressed_size_u64,
                limits.max_entry_uncompressed_bytes,
                DiagnosticCode::ZipEntryTooLarge,
                cursor + 24,
                "ZIP entry exceeds the uncompressed-size limit",
            )?;
            validate_sizes(
                compression,
                compressed_size,
                uncompressed_size,
                is_directory,
                &limits,
                cursor,
            )?;
            total_uncompressed = total_uncompressed
                .checked_add(uncompressed_size)
                .ok_or_else(|| {
                    zip_error(
                        DiagnosticCode::ZipTotalSizeLimit,
                        Some(cursor),
                        "cumulative uncompressed size overflow",
                    )
                })?;
            if total_uncompressed > limits.max_total_uncompressed_bytes {
                return Err(zip_error(
                    DiagnosticCode::ZipTotalSizeLimit,
                    Some(cursor),
                    "cumulative uncompressed size exceeds the configured limit",
                ));
            }

            let local_offset = size_with_limit(
                local_offset_u64,
                directory_offset,
                DiagnosticCode::ZipInvalid,
                cursor + 42,
                "local-header offset is outside the file-data region",
            )?;
            let local = validate_local_header(
                bytes,
                local_offset,
                raw_name,
                flags,
                compression,
                expected_crc,
                compressed_size,
                uncompressed_size,
                directory_offset,
                compressed_size_u32 == u32::MAX || uncompressed_size_u32 == u32::MAX,
            )?;
            local_ranges.push((local_offset, local.range_end));
            entries.push(ZipEntry {
                name,
                compression,
                crc32: expected_crc,
                compressed_size,
                uncompressed_size,
                data_start: local.data_start,
                is_directory,
            });
            cursor = next_cursor;
        }

        if cursor != directory_end {
            return Err(zip_error(
                DiagnosticCode::ZipInvalid,
                Some(cursor),
                "central-directory size or entry count is inconsistent",
            ));
        }
        local_ranges.sort_unstable_by_key(|&(start, _)| start);
        for ranges in local_ranges.windows(2) {
            if ranges[1].0 < ranges[0].1 {
                return Err(zip_error(
                    DiagnosticCode::ZipInvalid,
                    Some(ranges[1].0),
                    "ZIP local entry ranges overlap",
                ));
            }
        }

        Ok(Self {
            bytes,
            entries,
            limits,
            iwork_expanded_bytes_left: iwork_compatibility
                .then(|| Cell::new(limits.max_total_uncompressed_bytes)),
            deflate_operations_left: Cell::new(limits.max_deflate_operations),
        })
    }

    pub fn entries(&self) -> &[ZipEntry] {
        &self.entries
    }

    pub fn entry(&self, name: &str) -> Option<&ZipEntry> {
        self.entries.iter().find(|entry| entry.name == name)
    }

    pub fn extract(&self, entry: &ZipEntry) -> Result<Vec<u8>, Diagnostic> {
        let data_end = checked_add(entry.data_start, entry.compressed_size, entry.data_start)?;
        let compressed = self.bytes.get(entry.data_start..data_end).ok_or_else(|| {
            zip_error(
                DiagnosticCode::ZipInvalid,
                Some(entry.data_start),
                "entry data is outside the ZIP input",
            )
            .in_part(&entry.name)
        })?;
        if let Some(expanded_bytes_left) = &self.iwork_expanded_bytes_left {
            let available = expanded_bytes_left.get();
            if entry.uncompressed_size > available {
                return Err(zip_error(
                    DiagnosticCode::ZipTotalSizeLimit,
                    Some(entry.data_start),
                    "document-wide expanded-byte budget exhausted",
                )
                .in_part(&entry.name));
            }
            expanded_bytes_left.set(available - entry.uncompressed_size);
        }

        let output = match entry.compression {
            CompressionMethod::Stored => {
                let mut output = Vec::new();
                output.try_reserve_exact(compressed.len()).map_err(|_| {
                    zip_error(
                        DiagnosticCode::AllocationFailed,
                        Some(entry.data_start),
                        "unable to allocate stored ZIP entry",
                    )
                    .in_part(&entry.name)
                })?;
                output.extend_from_slice(compressed);
                output
            }
            CompressionMethod::Deflate => {
                let available = self.deflate_operations_left.get();
                match decompress_with_limit(
                    compressed,
                    self.limits.max_entry_uncompressed_bytes,
                    Some(entry.uncompressed_size),
                    available,
                ) {
                    Ok((output, used)) => {
                        self.deflate_operations_left.set(available - used);
                        output
                    }
                    Err(error) => {
                        if error.kind == DeflateErrorKind::OperationLimit {
                            self.deflate_operations_left.set(0);
                        }
                        let code = match error.kind {
                            DeflateErrorKind::AllocationFailed => DiagnosticCode::AllocationFailed,
                            DeflateErrorKind::OutputLimit => DiagnosticCode::ZipEntryTooLarge,
                            _ => DiagnosticCode::ZipDeflateInvalid,
                        };
                        let message = if error.kind == DeflateErrorKind::OperationLimit {
                            "document-wide Deflate operation budget exhausted".to_owned()
                        } else {
                            error.to_string()
                        };
                        return Err(zip_error(
                            code,
                            Some(entry.data_start + error.byte_offset),
                            message,
                        )
                        .in_part(&entry.name));
                    }
                }
            }
        };

        if output.len() != entry.uncompressed_size {
            return Err(zip_error(
                DiagnosticCode::ZipInvalid,
                Some(entry.data_start),
                "entry output size does not match the central directory",
            )
            .in_part(&entry.name));
        }
        if crc32(&output) != entry.crc32 {
            return Err(zip_error(
                DiagnosticCode::ZipCrcMismatch,
                Some(entry.data_start),
                "entry CRC-32 does not match the central directory",
            )
            .in_part(&entry.name));
        }
        Ok(output)
    }

    #[cfg(feature = "xps-formats")]
    pub(crate) fn extract_prefix(
        &self,
        entry: &ZipEntry,
        max_output: usize,
    ) -> Result<Vec<u8>, Diagnostic> {
        let data_end = checked_add(entry.data_start, entry.compressed_size, entry.data_start)?;
        let compressed = self.bytes.get(entry.data_start..data_end).ok_or_else(|| {
            zip_error(
                DiagnosticCode::ZipInvalid,
                Some(entry.data_start),
                "entry data is outside the ZIP input",
            )
            .in_part(&entry.name)
        })?;
        match entry.compression {
            CompressionMethod::Stored => {
                Ok(compressed[..compressed.len().min(max_output)].to_vec())
            }
            CompressionMethod::Deflate => {
                decompress_output_prefix(compressed, max_output).map_err(|error| {
                    zip_error(
                        match error.kind {
                            DeflateErrorKind::AllocationFailed => DiagnosticCode::AllocationFailed,
                            DeflateErrorKind::OutputLimit => DiagnosticCode::ZipEntryTooLarge,
                            _ => DiagnosticCode::ZipDeflateInvalid,
                        },
                        Some(entry.data_start + error.byte_offset),
                        error.to_string(),
                    )
                    .in_part(&entry.name)
                })
            }
        }
    }

    #[cfg(feature = "iwork-formats")]
    pub(crate) fn remaining_nested_budget(&self) -> (usize, usize) {
        (
            self.iwork_expanded_bytes_left.as_ref().map_or(0, Cell::get),
            self.deflate_operations_left.get(),
        )
    }

    #[cfg(feature = "iwork-formats")]
    pub(crate) fn consume_nested_budget(&self, expanded_bytes: usize, operations: usize) -> bool {
        let Some(expanded_bytes_left) = &self.iwork_expanded_bytes_left else {
            return false;
        };
        let bytes_left = expanded_bytes_left.get();
        let operations_left = self.deflate_operations_left.get();
        if expanded_bytes > bytes_left || operations > operations_left {
            return false;
        }
        expanded_bytes_left.set(bytes_left - expanded_bytes);
        self.deflate_operations_left
            .set(operations_left - operations);
        true
    }

    /// Extracts every entry once, enforcing Deflate output size and CRC checks.
    pub fn validate_all(&self) -> Result<(), Diagnostic> {
        for entry in &self.entries {
            self.extract(entry)?;
        }
        Ok(())
    }
}

struct LocalEntry {
    data_start: usize,
    range_end: usize,
}

#[derive(Clone, Copy)]
struct DirectoryRecord {
    entry_count: usize,
    offset: usize,
    end: usize,
}

#[derive(Clone, Copy, Default)]
struct Zip64Fields {
    uncompressed_size: Option<u64>,
    compressed_size: Option<u64>,
    local_offset: Option<u64>,
    start_disk: Option<u32>,
}

#[allow(clippy::too_many_arguments)]
fn validate_local_header(
    bytes: &[u8],
    offset: usize,
    central_name: &[u8],
    central_flags: u16,
    central_method: CompressionMethod,
    central_crc: u32,
    compressed_size: usize,
    uncompressed_size: usize,
    directory_offset: usize,
    central_uses_zip64_sizes: bool,
) -> Result<LocalEntry, Diagnostic> {
    ensure_range(bytes, offset, 30)?;
    if read_u32(bytes, offset)? != LOCAL_FILE_HEADER {
        return Err(zip_error(
            DiagnosticCode::ZipInvalid,
            Some(offset),
            "central-directory entry does not point to a local header",
        ));
    }
    let flags = read_u16(bytes, offset + 6)?;
    validate_flags(flags, offset + 6)?;
    let flag_difference = flags ^ central_flags;
    let compatible_difference = flag_difference & !(FLAG_DATA_DESCRIPTOR | FLAG_UTF8) == 0
        && (flag_difference & FLAG_UTF8 == 0 || std::str::from_utf8(central_name).is_ok());
    if flags != central_flags && !compatible_difference {
        return Err(zip_error(
            DiagnosticCode::ZipInvalid,
            Some(offset + 6),
            "local and central flags differ beyond the data-descriptor bit",
        ));
    }
    let method = compression_method(read_u16(bytes, offset + 8)?, offset + 8)?;
    if method != central_method {
        return Err(zip_error(
            DiagnosticCode::ZipInvalid,
            Some(offset + 8),
            "local and central compression methods differ",
        ));
    }

    let local_crc = read_u32(bytes, offset + 14)?;
    let local_compressed = read_u32(bytes, offset + 18)?;
    let local_uncompressed = read_u32(bytes, offset + 22)?;
    let name_length = usize::from(read_u16(bytes, offset + 26)?);
    let extra_length = usize::from(read_u16(bytes, offset + 28)?);
    let name_start = checked_add(offset, 30, offset)?;
    let extra_start = checked_add(name_start, name_length, offset)?;
    let data_start = checked_add(extra_start, extra_length, offset)?;
    ensure_range(bytes, name_start, name_length)?;
    ensure_range(bytes, extra_start, extra_length)?;
    if &bytes[name_start..name_start + name_length] != central_name {
        return Err(zip_error(
            DiagnosticCode::ZipInvalid,
            Some(name_start),
            "local and central entry names differ",
        ));
    }
    let zip64 = parse_local_extra(
        &bytes[extra_start..extra_start + extra_length],
        extra_start,
        local_uncompressed == u32::MAX,
        local_compressed == u32::MAX,
        (flags & FLAG_DATA_DESCRIPTOR == 0)
            .then_some((uncompressed_size as u64, compressed_size as u64)),
    )?;
    let local_compressed = resolve_zip64_u32(
        local_compressed,
        zip64.compressed_size,
        offset + 18,
        "local compressed size",
    )?;
    let local_uncompressed = resolve_zip64_u32(
        local_uncompressed,
        zip64.uncompressed_size,
        offset + 22,
        "local uncompressed size",
    )?;
    if flags & FLAG_DATA_DESCRIPTOR == 0 {
        if local_compressed != compressed_size as u64
            || local_uncompressed != uncompressed_size as u64
        {
            return Err(zip_error(
                DiagnosticCode::ZipInvalid,
                Some(offset + 14),
                "local and central CRC or sizes differ",
            ));
        }
    } else if (local_crc != 0 && local_crc != central_crc)
        || (local_compressed != 0 && local_compressed != compressed_size as u64)
        || (local_uncompressed != 0 && local_uncompressed != uncompressed_size as u64)
    {
        return Err(zip_error(
            DiagnosticCode::ZipInvalid,
            Some(offset + 14),
            "local data-descriptor placeholders conflict with central values",
        ));
    }
    let data_end = checked_add(data_start, compressed_size, data_start)?;
    if data_end > directory_offset {
        return Err(zip_error(
            DiagnosticCode::ZipInvalid,
            Some(data_start),
            "local entry data overlaps the central directory",
        ));
    }

    let range_end = if flags & FLAG_DATA_DESCRIPTOR != 0 {
        validate_data_descriptor(
            bytes,
            data_end,
            central_crc,
            compressed_size,
            uncompressed_size,
            directory_offset,
            central_uses_zip64_sizes
                || local_compressed > u64::from(u32::MAX)
                || local_uncompressed > u64::from(u32::MAX),
        )?
    } else {
        data_end
    };
    Ok(LocalEntry {
        data_start,
        range_end,
    })
}

fn validate_data_descriptor(
    bytes: &[u8],
    offset: usize,
    crc: u32,
    compressed_size: usize,
    uncompressed_size: usize,
    directory_offset: usize,
    zip64: bool,
) -> Result<usize, Diagnostic> {
    if zip64 {
        let compressed_size = compressed_size as u64;
        let uncompressed_size = uncompressed_size as u64;
        let signed_matches = offset
            .checked_add(24)
            .is_some_and(|end| end <= directory_offset && end <= bytes.len())
            && read_u32(bytes, offset).ok() == Some(DATA_DESCRIPTOR)
            && read_u32(bytes, offset + 4).ok() == Some(crc)
            && read_u64(bytes, offset + 8).ok() == Some(compressed_size)
            && read_u64(bytes, offset + 16).ok() == Some(uncompressed_size);
        if signed_matches {
            return Ok(offset + 24);
        }

        let unsigned_matches = offset
            .checked_add(20)
            .is_some_and(|end| end <= directory_offset && end <= bytes.len())
            && read_u32(bytes, offset).ok() == Some(crc)
            && read_u64(bytes, offset + 4).ok() == Some(compressed_size)
            && read_u64(bytes, offset + 12).ok() == Some(uncompressed_size);
        if unsigned_matches {
            return Ok(offset + 20);
        }
        return Err(zip_error(
            DiagnosticCode::ZipInvalid,
            Some(offset),
            "missing or inconsistent ZIP64 data descriptor",
        ));
    }

    let compressed_size = compressed_size as u32;
    let uncompressed_size = uncompressed_size as u32;

    let signed_matches = offset
        .checked_add(16)
        .is_some_and(|end| end <= directory_offset && end <= bytes.len())
        && read_u32(bytes, offset).ok() == Some(DATA_DESCRIPTOR)
        && read_u32(bytes, offset + 4).ok() == Some(crc)
        && read_u32(bytes, offset + 8).ok() == Some(compressed_size)
        && read_u32(bytes, offset + 12).ok() == Some(uncompressed_size);
    if signed_matches {
        return Ok(offset + 16);
    }

    let unsigned_matches = offset
        .checked_add(12)
        .is_some_and(|end| end <= directory_offset && end <= bytes.len())
        && read_u32(bytes, offset).ok() == Some(crc)
        && read_u32(bytes, offset + 4).ok() == Some(compressed_size)
        && read_u32(bytes, offset + 8).ok() == Some(uncompressed_size);
    if unsigned_matches {
        return Ok(offset + 12);
    }

    Err(zip_error(
        DiagnosticCode::ZipInvalid,
        Some(offset),
        "missing or inconsistent ZIP data descriptor",
    ))
}

fn validate_sizes(
    method: CompressionMethod,
    compressed: usize,
    uncompressed: usize,
    is_directory: bool,
    limits: &Limits,
    offset: usize,
) -> Result<(), Diagnostic> {
    if uncompressed > limits.max_entry_uncompressed_bytes {
        return Err(zip_error(
            DiagnosticCode::ZipEntryTooLarge,
            Some(offset),
            "ZIP entry exceeds the uncompressed-size limit",
        ));
    }
    if is_directory && uncompressed != 0 {
        return Err(zip_error(
            DiagnosticCode::ZipInvalid,
            Some(offset),
            "directory entry expands to data",
        ));
    }
    if method == CompressionMethod::Stored && compressed != uncompressed {
        return Err(zip_error(
            DiagnosticCode::ZipInvalid,
            Some(offset),
            "stored ZIP entry has unequal compressed and uncompressed sizes",
        ));
    }
    if uncompressed >= limits.compression_ratio_min_bytes
        && (compressed == 0
            || uncompressed > compressed.saturating_mul(limits.max_compression_ratio as usize))
    {
        return Err(zip_error(
            DiagnosticCode::ZipCompressionRatioLimit,
            Some(offset),
            "ZIP entry exceeds the compression-ratio limit",
        ));
    }
    Ok(())
}

fn validate_flags(flags: u16, offset: usize) -> Result<(), Diagnostic> {
    if flags & FLAG_ENCRYPTED != 0 || flags & (1 << 6) != 0 || flags & (1 << 13) != 0 {
        return Err(zip_error(
            DiagnosticCode::ZipEncrypted,
            Some(offset),
            "encrypted ZIP entries are forbidden",
        ));
    }
    if flags & !ALLOWED_FLAGS != 0 {
        return Err(zip_error(
            DiagnosticCode::ZipInvalid,
            Some(offset),
            "ZIP entry uses unsupported general-purpose flags",
        ));
    }
    Ok(())
}

fn compression_method(value: u16, offset: usize) -> Result<CompressionMethod, Diagnostic> {
    match value {
        0 => Ok(CompressionMethod::Stored),
        8 => Ok(CompressionMethod::Deflate),
        _ => Err(zip_error(
            DiagnosticCode::ZipUnsupportedCompression,
            Some(offset),
            "ZIP compression method is unsupported",
        )),
    }
}

fn validate_path(
    raw: &[u8],
    flags: u16,
    max_bytes: usize,
    offset: usize,
    allow_unflagged_utf8: bool,
) -> Result<(String, bool), Diagnostic> {
    if raw.is_empty() || raw.len() > max_bytes {
        return Err(zip_error(
            DiagnosticCode::ZipPathTraversal,
            Some(offset),
            "ZIP entry path is empty or too long",
        ));
    }
    if raw.contains(&0) {
        return Err(zip_error(
            DiagnosticCode::ZipPathTraversal,
            Some(offset),
            "ZIP entry path contains a null byte",
        ));
    }
    let normalized;
    let raw = if raw.contains(&b'\\') {
        normalized = raw
            .iter()
            .map(|byte| if *byte == b'\\' { b'/' } else { *byte })
            .collect::<Vec<_>>();
        normalized.as_slice()
    } else {
        raw
    };
    if raw[0] == b'/' {
        return Err(zip_error(
            DiagnosticCode::ZipPathTraversal,
            Some(offset),
            "ZIP entry path is absolute",
        ));
    }

    let path = if flags & FLAG_UTF8 != 0 {
        std::str::from_utf8(raw).map_err(|_| {
            zip_error(
                DiagnosticCode::ZipInvalid,
                Some(offset),
                "ZIP entry declares invalid UTF-8",
            )
        })?
    } else if raw.is_ascii() {
        std::str::from_utf8(raw).map_err(|_| {
            zip_error(
                DiagnosticCode::ZipInvalid,
                Some(offset),
                "ZIP entry path is not valid UTF-8",
            )
        })?
    } else if allow_unflagged_utf8 {
        std::str::from_utf8(raw).map_err(|_| {
            zip_error(
                DiagnosticCode::ZipInvalid,
                Some(offset),
                "iWork ZIP path without the UTF-8 flag is not valid UTF-8",
            )
        })?
    } else {
        return Err(zip_error(
            DiagnosticCode::ZipInvalid,
            Some(offset),
            "non-ASCII ZIP path requires the UTF-8 flag",
        ));
    };
    if path.chars().any(char::is_control) {
        return Err(zip_error(
            DiagnosticCode::ZipPathTraversal,
            Some(offset),
            "ZIP entry path contains control characters",
        ));
    }

    let is_directory = path.ends_with('/');
    let canonical = if is_directory {
        &path[..path.len() - 1]
    } else {
        path
    };
    if canonical.is_empty()
        || canonical
            .split('/')
            .any(|part| part.is_empty() || part == "." || part == "..")
    {
        return Err(zip_error(
            DiagnosticCode::ZipPathTraversal,
            Some(offset),
            "ZIP entry path contains an empty, current, or parent segment",
        ));
    }
    let first = canonical.split('/').next().unwrap_or_default().as_bytes();
    if first.len() == 2 && first[0].is_ascii_alphabetic() && first[1] == b':' {
        return Err(zip_error(
            DiagnosticCode::ZipPathTraversal,
            Some(offset),
            "ZIP entry path contains a drive prefix",
        ));
    }
    Ok((canonical.to_owned(), is_directory))
}

fn parse_central_extra(
    extra: &[u8],
    base_offset: usize,
    needs_uncompressed: bool,
    needs_compressed: bool,
    needs_offset: bool,
    needs_disk: bool,
) -> Result<Zip64Fields, Diagnostic> {
    parse_zip64_extra(
        extra,
        base_offset,
        needs_uncompressed,
        needs_compressed,
        needs_offset,
        needs_disk,
        None,
    )
}

fn parse_local_extra(
    extra: &[u8],
    base_offset: usize,
    needs_uncompressed: bool,
    needs_compressed: bool,
    redundant_sizes: Option<(u64, u64)>,
) -> Result<Zip64Fields, Diagnostic> {
    parse_zip64_extra(
        extra,
        base_offset,
        needs_uncompressed,
        needs_compressed,
        false,
        false,
        redundant_sizes,
    )
}

#[allow(clippy::too_many_arguments)]
fn parse_zip64_extra(
    extra: &[u8],
    base_offset: usize,
    needs_uncompressed: bool,
    needs_compressed: bool,
    needs_offset: bool,
    needs_disk: bool,
    redundant_sizes: Option<(u64, u64)>,
) -> Result<Zip64Fields, Diagnostic> {
    let mut cursor = 0;
    let mut result = Zip64Fields::default();
    let mut saw_zip64 = false;
    while cursor < extra.len() {
        if extra.len() - cursor < 4 {
            return Err(zip_error(
                DiagnosticCode::ZipInvalid,
                Some(base_offset + cursor),
                "truncated ZIP extra-field header",
            ));
        }
        let kind = u16::from_le_bytes([extra[cursor], extra[cursor + 1]]);
        let length = usize::from(u16::from_le_bytes([extra[cursor + 2], extra[cursor + 3]]));
        let end = cursor
            .checked_add(4)
            .and_then(|value| value.checked_add(length))
            .ok_or_else(|| {
                zip_error(
                    DiagnosticCode::ZipInvalid,
                    Some(base_offset + cursor),
                    "ZIP extra-field length overflow",
                )
            })?;
        if end > extra.len() {
            return Err(zip_error(
                DiagnosticCode::ZipInvalid,
                Some(base_offset + cursor),
                "truncated ZIP extra-field value",
            ));
        }
        if kind == 0x0001 {
            if saw_zip64 {
                return Err(zip_error(
                    DiagnosticCode::ZipInvalid,
                    Some(base_offset + cursor),
                    "duplicate ZIP64 extra field",
                ));
            }
            saw_zip64 = true;
            let value = &extra[cursor + 4..end];
            let expected_length = usize::from(needs_uncompressed) * 8
                + usize::from(needs_compressed) * 8
                + usize::from(needs_offset) * 8
                + usize::from(needs_disk) * 4;
            if expected_length == 0 {
                let compatible = redundant_sizes.is_some_and(|(uncompressed, compressed)| {
                    if length != 16 {
                        return false;
                    }
                    let Ok(extra_uncompressed) = read_u64(value, 0) else {
                        return false;
                    };
                    let Ok(extra_compressed) = read_u64(value, 8) else {
                        return false;
                    };
                    (extra_uncompressed == 0 && extra_compressed == 0)
                        || (extra_uncompressed == uncompressed && extra_compressed == compressed)
                });
                if !compatible {
                    return Err(zip_error(
                        DiagnosticCode::ZipInvalid,
                        Some(base_offset + cursor),
                        "unexpected redundant ZIP64 extra field",
                    ));
                }
            } else {
                if length != expected_length {
                    return Err(zip_error(
                        DiagnosticCode::ZipInvalid,
                        Some(base_offset + cursor),
                        "ZIP64 extra field has an inconsistent length",
                    ));
                }
                let mut value_cursor = 0;
                if needs_uncompressed {
                    result.uncompressed_size = Some(read_u64(value, value_cursor)?);
                    value_cursor += 8;
                }
                if needs_compressed {
                    result.compressed_size = Some(read_u64(value, value_cursor)?);
                    value_cursor += 8;
                }
                if needs_offset {
                    result.local_offset = Some(read_u64(value, value_cursor)?);
                    value_cursor += 8;
                }
                if needs_disk {
                    result.start_disk = Some(read_u32(value, value_cursor)?);
                }
            }
        }
        cursor = end;
    }
    if (needs_uncompressed || needs_compressed || needs_offset || needs_disk) && !saw_zip64 {
        return Err(zip_error(
            DiagnosticCode::ZipInvalid,
            Some(base_offset),
            "required ZIP64 extra field is missing",
        ));
    }
    Ok(result)
}

fn parse_directory_record(bytes: &[u8], eocd: usize) -> Result<DirectoryRecord, Diagnostic> {
    let disk = read_u16(bytes, eocd + 4)?;
    let directory_disk = read_u16(bytes, eocd + 6)?;
    let entries_on_disk = read_u16(bytes, eocd + 8)?;
    let entry_count = read_u16(bytes, eocd + 10)?;
    let directory_size = read_u32(bytes, eocd + 12)?;
    let directory_offset = read_u32(bytes, eocd + 16)?;
    if disk != 0 || directory_disk != 0 {
        return Err(zip_error(
            DiagnosticCode::ZipMultiDiskForbidden,
            Some(eocd + 4),
            "multi-disk ZIP archives are not supported",
        ));
    }

    let locator = eocd.checked_sub(20).filter(|&offset| {
        read_u32(bytes, offset).ok() == Some(ZIP64_END_OF_CENTRAL_DIRECTORY_LOCATOR)
    });
    let uses_sentinels = entries_on_disk == u16::MAX
        || entry_count == u16::MAX
        || directory_size == u32::MAX
        || directory_offset == u32::MAX;

    let (entry_count_u64, directory_size_u64, directory_offset_u64, directory_end) =
        if let Some(locator) = locator {
            let zip64_disk = read_u32(bytes, locator + 4)?;
            let zip64_offset_u64 = read_u64(bytes, locator + 8)?;
            let disk_count = read_u32(bytes, locator + 16)?;
            if zip64_disk != 0 || disk_count != 1 {
                return Err(zip_error(
                    DiagnosticCode::ZipMultiDiskForbidden,
                    Some(locator + 4),
                    "multi-disk ZIP64 archives are not supported",
                ));
            }
            let zip64_offset = size_with_limit(
                zip64_offset_u64,
                locator,
                DiagnosticCode::ZipInvalid,
                locator + 8,
                "ZIP64 end record offset is outside the archive",
            )?;
            ensure_range(bytes, zip64_offset, 56)?;
            if read_u32(bytes, zip64_offset)? != ZIP64_END_OF_CENTRAL_DIRECTORY {
                return Err(zip_error(
                    DiagnosticCode::ZipInvalid,
                    Some(zip64_offset),
                    "ZIP64 locator does not point to a ZIP64 end record",
                ));
            }
            let record_size = read_u64(bytes, zip64_offset + 4)?;
            if record_size < 44 {
                return Err(zip_error(
                    DiagnosticCode::ZipInvalid,
                    Some(zip64_offset + 4),
                    "ZIP64 end record is too short",
                ));
            }
            let record_size = size_with_limit(
                record_size,
                locator,
                DiagnosticCode::ZipInvalid,
                zip64_offset + 4,
                "ZIP64 end record is too large",
            )?;
            let record_end = checked_add(zip64_offset, 12 + record_size, zip64_offset)?;
            if record_end != locator {
                return Err(zip_error(
                    DiagnosticCode::ZipInvalid,
                    Some(zip64_offset),
                    "ZIP64 end record does not end at its locator",
                ));
            }
            if read_u32(bytes, zip64_offset + 16)? != 0 || read_u32(bytes, zip64_offset + 20)? != 0
            {
                return Err(zip_error(
                    DiagnosticCode::ZipMultiDiskForbidden,
                    Some(zip64_offset + 16),
                    "multi-disk ZIP64 archives are not supported",
                ));
            }
            let zip64_entries_on_disk = read_u64(bytes, zip64_offset + 24)?;
            let zip64_entry_count = read_u64(bytes, zip64_offset + 32)?;
            if zip64_entries_on_disk != zip64_entry_count {
                return Err(zip_error(
                    DiagnosticCode::ZipMultiDiskForbidden,
                    Some(zip64_offset + 24),
                    "ZIP64 entry counts indicate multiple disks",
                ));
            }
            let zip64_directory_size = read_u64(bytes, zip64_offset + 40)?;
            let zip64_directory_offset = read_u64(bytes, zip64_offset + 48)?;
            cross_check_legacy_u16(entries_on_disk, zip64_entries_on_disk, eocd + 8)?;
            cross_check_legacy_u16(entry_count, zip64_entry_count, eocd + 10)?;
            cross_check_legacy_u32(directory_size, zip64_directory_size, eocd + 12)?;
            cross_check_legacy_u32(directory_offset, zip64_directory_offset, eocd + 16)?;
            (
                zip64_entry_count,
                zip64_directory_size,
                zip64_directory_offset,
                zip64_offset,
            )
        } else {
            if uses_sentinels {
                return Err(zip_error(
                    DiagnosticCode::ZipInvalid,
                    Some(eocd),
                    "ZIP64 sentinel values require a ZIP64 end record and locator",
                ));
            }
            if entries_on_disk != entry_count {
                return Err(zip_error(
                    DiagnosticCode::ZipMultiDiskForbidden,
                    Some(eocd + 8),
                    "ZIP entry counts indicate multiple disks",
                ));
            }
            (
                u64::from(entry_count),
                u64::from(directory_size),
                u64::from(directory_offset),
                eocd,
            )
        };

    let entry_count = usize::try_from(entry_count_u64).map_err(|_| {
        zip_error(
            DiagnosticCode::ZipEntryLimit,
            Some(eocd + 10),
            "ZIP entry count is not addressable",
        )
    })?;
    let offset = size_with_limit(
        directory_offset_u64,
        directory_end,
        DiagnosticCode::ZipInvalid,
        eocd + 16,
        "central-directory offset is outside the archive",
    )?;
    let size = size_with_limit(
        directory_size_u64,
        directory_end,
        DiagnosticCode::ZipInvalid,
        eocd + 12,
        "central-directory size is outside the archive",
    )?;
    if checked_add(offset, size, offset)? != directory_end {
        return Err(zip_error(
            DiagnosticCode::ZipInvalid,
            Some(offset),
            "central directory does not end at the end record",
        ));
    }
    Ok(DirectoryRecord {
        entry_count,
        offset,
        end: directory_end,
    })
}

fn cross_check_legacy_u16(value: u16, zip64: u64, offset: usize) -> Result<(), Diagnostic> {
    if value != u16::MAX && u64::from(value) != zip64 {
        return Err(zip_error(
            DiagnosticCode::ZipInvalid,
            Some(offset),
            "legacy and ZIP64 end records disagree",
        ));
    }
    Ok(())
}

fn cross_check_legacy_u32(value: u32, zip64: u64, offset: usize) -> Result<(), Diagnostic> {
    if value != u32::MAX && u64::from(value) != zip64 {
        return Err(zip_error(
            DiagnosticCode::ZipInvalid,
            Some(offset),
            "legacy and ZIP64 end records disagree",
        ));
    }
    Ok(())
}

fn resolve_zip64_u32(
    legacy: u32,
    zip64: Option<u64>,
    offset: usize,
    field: &str,
) -> Result<u64, Diagnostic> {
    if legacy == u32::MAX {
        zip64.ok_or_else(|| {
            zip_error(
                DiagnosticCode::ZipInvalid,
                Some(offset),
                format!("ZIP64 {field} is missing"),
            )
        })
    } else {
        if zip64.is_some() {
            return Err(zip_error(
                DiagnosticCode::ZipInvalid,
                Some(offset),
                format!("unexpected ZIP64 {field}"),
            ));
        }
        Ok(u64::from(legacy))
    }
}

fn size_with_limit(
    value: u64,
    limit: usize,
    code: DiagnosticCode,
    offset: usize,
    message: &str,
) -> Result<usize, Diagnostic> {
    if value > limit as u64 {
        return Err(zip_error(code, Some(offset), message));
    }
    usize::try_from(value).map_err(|_| zip_error(code, Some(offset), message))
}

fn find_eocd(bytes: &[u8]) -> Result<usize, Diagnostic> {
    if bytes.len() < 22 {
        return Err(zip_error(
            DiagnosticCode::ZipInvalid,
            None,
            "ZIP is too short to contain EOCD",
        ));
    }
    let minimum = bytes.len().saturating_sub(22 + usize::from(u16::MAX));
    for offset in (minimum..=bytes.len() - 22).rev() {
        if read_u32(bytes, offset).ok() != Some(END_OF_CENTRAL_DIRECTORY) {
            continue;
        }
        let comment_length = usize::from(read_u16(bytes, offset + 20)?);
        if offset
            .checked_add(22 + comment_length)
            .is_some_and(|end| end == bytes.len())
        {
            return Ok(offset);
        }
    }
    Err(zip_error(
        DiagnosticCode::ZipInvalid,
        None,
        "valid end-of-central-directory record not found",
    ))
}

fn ensure_range(bytes: &[u8], offset: usize, length: usize) -> Result<(), Diagnostic> {
    if offset
        .checked_add(length)
        .is_none_or(|end| end > bytes.len())
    {
        return Err(zip_error(
            DiagnosticCode::ZipInvalid,
            Some(offset),
            "ZIP structure is truncated or has an overflowing range",
        ));
    }
    Ok(())
}

fn checked_add(left: usize, right: usize, offset: usize) -> Result<usize, Diagnostic> {
    left.checked_add(right).ok_or_else(|| {
        zip_error(
            DiagnosticCode::ZipInvalid,
            Some(offset),
            "ZIP offset arithmetic overflow",
        )
    })
}

fn read_u16(bytes: &[u8], offset: usize) -> Result<u16, Diagnostic> {
    ensure_range(bytes, offset, 2)?;
    Ok(u16::from_le_bytes([bytes[offset], bytes[offset + 1]]))
}

fn read_u32(bytes: &[u8], offset: usize) -> Result<u32, Diagnostic> {
    ensure_range(bytes, offset, 4)?;
    Ok(u32::from_le_bytes([
        bytes[offset],
        bytes[offset + 1],
        bytes[offset + 2],
        bytes[offset + 3],
    ]))
}

fn read_u64(bytes: &[u8], offset: usize) -> Result<u64, Diagnostic> {
    ensure_range(bytes, offset, 8)?;
    Ok(u64::from_le_bytes([
        bytes[offset],
        bytes[offset + 1],
        bytes[offset + 2],
        bytes[offset + 3],
        bytes[offset + 4],
        bytes[offset + 5],
        bytes[offset + 6],
        bytes[offset + 7],
    ]))
}

fn zip_error(
    code: DiagnosticCode,
    offset: Option<usize>,
    message: impl Into<String>,
) -> Diagnostic {
    Diagnostic::fatal(code, Phase::Container, offset, message)
}

/// CRC-32/ISO-HDLC used by ZIP (reflected polynomial `0xedb88320`).
pub fn crc32(bytes: &[u8]) -> u32 {
    const TABLE: [u32; 256] = {
        let mut table = [0; 256];
        let mut index = 0;
        while index < table.len() {
            let mut crc = index as u32;
            let mut bit = 0;
            while bit < 8 {
                crc = (crc >> 1) ^ (0xedb8_8320 & 0_u32.wrapping_sub(crc & 1));
                bit += 1;
            }
            table[index] = crc;
            index += 1;
        }
        table
    };
    let mut crc = u32::MAX;
    for &byte in bytes {
        crc = (crc >> 8) ^ TABLE[((crc ^ u32::from(byte)) & 0xff) as usize];
    }
    !crc
}

#[cfg(test)]
mod tests {
    use super::{CompressionMethod, ZipArchive, crc32};
    use crate::diagnostic::DiagnosticCode;
    use crate::limits::Limits;

    fn u16le(output: &mut Vec<u8>, value: u16) {
        output.extend_from_slice(&value.to_le_bytes());
    }

    fn u32le(output: &mut Vec<u8>, value: u32) {
        output.extend_from_slice(&value.to_le_bytes());
    }

    fn u64le(output: &mut Vec<u8>, value: u64) {
        output.extend_from_slice(&value.to_le_bytes());
    }

    fn one_entry_zip(
        name: &[u8],
        data: &[u8],
        compressed: &[u8],
        method: u16,
        flags: u16,
    ) -> Vec<u8> {
        let crc = crc32(data);
        let mut output = Vec::new();
        u32le(&mut output, super::LOCAL_FILE_HEADER);
        u16le(&mut output, 20);
        u16le(&mut output, flags);
        u16le(&mut output, method);
        u16le(&mut output, 0);
        u16le(&mut output, 0);
        u32le(&mut output, crc);
        u32le(&mut output, compressed.len() as u32);
        u32le(&mut output, data.len() as u32);
        u16le(&mut output, name.len() as u16);
        u16le(&mut output, 0);
        output.extend_from_slice(name);
        output.extend_from_slice(compressed);
        if flags & super::FLAG_DATA_DESCRIPTOR != 0 {
            u32le(&mut output, super::DATA_DESCRIPTOR);
            u32le(&mut output, crc);
            u32le(&mut output, compressed.len() as u32);
            u32le(&mut output, data.len() as u32);
        }

        let directory_offset = output.len() as u32;
        u32le(&mut output, super::CENTRAL_DIRECTORY_HEADER);
        u16le(&mut output, 20);
        u16le(&mut output, 20);
        u16le(&mut output, flags);
        u16le(&mut output, method);
        u16le(&mut output, 0);
        u16le(&mut output, 0);
        u32le(&mut output, crc);
        u32le(&mut output, compressed.len() as u32);
        u32le(&mut output, data.len() as u32);
        u16le(&mut output, name.len() as u16);
        u16le(&mut output, 0);
        u16le(&mut output, 0);
        u16le(&mut output, 0);
        u16le(&mut output, 0);
        u32le(&mut output, 0);
        u32le(&mut output, 0);
        output.extend_from_slice(name);

        let directory_size = output.len() as u32 - directory_offset;
        u32le(&mut output, super::END_OF_CENTRAL_DIRECTORY);
        u16le(&mut output, 0);
        u16le(&mut output, 0);
        u16le(&mut output, 1);
        u16le(&mut output, 1);
        u32le(&mut output, directory_size);
        u32le(&mut output, directory_offset);
        u16le(&mut output, 0);
        output
    }

    fn with_redundant_local_zip64(mut zip: Vec<u8>, name: &[u8], size: u64) -> Vec<u8> {
        let mut extra = Vec::new();
        u16le(&mut extra, 0x0001);
        u16le(&mut extra, 16);
        extra.extend_from_slice(&size.to_le_bytes());
        extra.extend_from_slice(&size.to_le_bytes());
        zip.splice(30 + name.len()..30 + name.len(), extra);
        zip[28..30].copy_from_slice(&20_u16.to_le_bytes());
        let eocd = zip.len() - 22;
        let directory_offset = u32::from_le_bytes(zip[eocd + 16..eocd + 20].try_into().unwrap());
        zip[eocd + 16..eocd + 20].copy_from_slice(&(directory_offset + 20).to_le_bytes());
        zip
    }

    fn one_entry_zip64(name: &[u8], data: &[u8], descriptor: bool) -> Vec<u8> {
        let crc = crc32(data);
        let flags = if descriptor {
            super::FLAG_DATA_DESCRIPTOR
        } else {
            0
        };
        let mut output = Vec::new();
        u32le(&mut output, super::LOCAL_FILE_HEADER);
        u16le(&mut output, 45);
        u16le(&mut output, flags);
        u16le(&mut output, 0);
        u16le(&mut output, 0);
        u16le(&mut output, 0);
        u32le(&mut output, if descriptor { 0 } else { crc });
        u32le(&mut output, u32::MAX);
        u32le(&mut output, u32::MAX);
        u16le(&mut output, name.len() as u16);
        u16le(&mut output, 20);
        output.extend_from_slice(name);
        u16le(&mut output, 0x0001);
        u16le(&mut output, 16);
        u64le(&mut output, data.len() as u64);
        u64le(&mut output, data.len() as u64);
        output.extend_from_slice(data);
        if descriptor {
            u32le(&mut output, super::DATA_DESCRIPTOR);
            u32le(&mut output, crc);
            u64le(&mut output, data.len() as u64);
            u64le(&mut output, data.len() as u64);
        }

        let directory_offset = output.len() as u64;
        u32le(&mut output, super::CENTRAL_DIRECTORY_HEADER);
        u16le(&mut output, 45);
        u16le(&mut output, 45);
        u16le(&mut output, flags);
        u16le(&mut output, 0);
        u16le(&mut output, 0);
        u16le(&mut output, 0);
        u32le(&mut output, crc);
        u32le(&mut output, u32::MAX);
        u32le(&mut output, u32::MAX);
        u16le(&mut output, name.len() as u16);
        u16le(&mut output, 28);
        u16le(&mut output, 0);
        u16le(&mut output, 0);
        u16le(&mut output, 0);
        u32le(&mut output, 0);
        u32le(&mut output, u32::MAX);
        output.extend_from_slice(name);
        u16le(&mut output, 0x0001);
        u16le(&mut output, 24);
        u64le(&mut output, data.len() as u64);
        u64le(&mut output, data.len() as u64);
        u64le(&mut output, 0);

        let directory_size = output.len() as u64 - directory_offset;
        let zip64_eocd = output.len() as u64;
        u32le(&mut output, super::ZIP64_END_OF_CENTRAL_DIRECTORY);
        u64le(&mut output, 44);
        u16le(&mut output, 45);
        u16le(&mut output, 45);
        u32le(&mut output, 0);
        u32le(&mut output, 0);
        u64le(&mut output, 1);
        u64le(&mut output, 1);
        u64le(&mut output, directory_size);
        u64le(&mut output, directory_offset);
        u32le(&mut output, super::ZIP64_END_OF_CENTRAL_DIRECTORY_LOCATOR);
        u32le(&mut output, 0);
        u64le(&mut output, zip64_eocd);
        u32le(&mut output, 1);
        u32le(&mut output, super::END_OF_CENTRAL_DIRECTORY);
        u16le(&mut output, 0);
        u16le(&mut output, 0);
        u16le(&mut output, u16::MAX);
        u16le(&mut output, u16::MAX);
        u32le(&mut output, u32::MAX);
        u32le(&mut output, u32::MAX);
        u16le(&mut output, 0);
        output
    }

    #[test]
    fn crc_matches_standard_vector() {
        assert_eq!(crc32(b"123456789"), 0xcbf4_3926);
        assert_eq!(crc32(b""), 0);
        let bytes: Vec<u8> = (0..=u16::MAX).flat_map(u16::to_le_bytes).collect();
        let mut reference = u32::MAX;
        for (index, &byte) in bytes.iter().enumerate() {
            reference ^= u32::from(byte);
            for _ in 0..8 {
                reference = (reference >> 1) ^ (0xedb8_8320 & 0_u32.wrapping_sub(reference & 1));
            }
            if index < 256 || index.is_multiple_of(1024) || index + 1 == bytes.len() {
                assert_eq!(crc32(&bytes[..=index]), !reference);
            }
        }
    }

    #[test]
    fn reads_stored_and_deflated_entries() {
        let stored = one_entry_zip(b"word/document.xml", b"hello", b"hello", 0, 0);
        let archive = ZipArchive::parse(&stored, Limits::default()).unwrap();
        let entry = &archive.entries()[0];
        assert_eq!(entry.compression(), CompressionMethod::Stored);
        assert_eq!(archive.extract(entry).unwrap(), b"hello");

        let compressed = [0xcb, 0x48, 0xcd, 0xc9, 0xc9, 0x07, 0x00];
        let deflated = one_entry_zip(b"content.xml", b"hello", &compressed, 8, 0);
        let archive = ZipArchive::parse(&deflated, Limits::default()).unwrap();
        archive.validate_all().unwrap();
        assert_eq!(archive.extract(&archive.entries()[0]).unwrap(), b"hello");
    }

    #[test]
    fn accepts_only_safe_redundant_local_zip64_pairs() {
        let name = b"preview.jpg";
        let compatible =
            with_redundant_local_zip64(one_entry_zip(name, b"image", b"image", 0, 0), name, 5);
        let archive = ZipArchive::parse(&compatible, Limits::default()).unwrap();
        assert_eq!(archive.extract(&archive.entries()[0]).unwrap(), b"image");

        let zero =
            with_redundant_local_zip64(one_entry_zip(name, b"image", b"image", 0, 0), name, 0);
        let archive = ZipArchive::parse(&zero, Limits::default()).unwrap();
        assert_eq!(archive.extract(&archive.entries()[0]).unwrap(), b"image");

        let mut mixed = zero;
        let extra_compressed = 30 + name.len() + 4 + 8;
        mixed[extra_compressed..extra_compressed + 8].copy_from_slice(&5_u64.to_le_bytes());
        assert_eq!(
            ZipArchive::parse(&mixed, Limits::default())
                .unwrap_err()
                .code,
            DiagnosticCode::ZipInvalid,
        );

        let mismatched =
            with_redundant_local_zip64(one_entry_zip(name, b"image", b"image", 0, 0), name, 6);
        assert_eq!(
            ZipArchive::parse(&mismatched, Limits::default())
                .unwrap_err()
                .code,
            DiagnosticCode::ZipInvalid,
        );

        let mut descriptor =
            one_entry_zip(name, b"image", b"image", 0, super::FLAG_DATA_DESCRIPTOR);
        descriptor[14..26].fill(0);
        let descriptor = with_redundant_local_zip64(descriptor, name, 0);
        assert_eq!(
            ZipArchive::parse(&descriptor, Limits::default())
                .unwrap_err()
                .code,
            DiagnosticCode::ZipInvalid,
        );
    }

    #[test]
    fn reads_zip64_entries_and_descriptors() {
        for descriptor in [false, true] {
            let zip = one_entry_zip64(b"word/document.xml", b"hello", descriptor);
            let archive = ZipArchive::parse(&zip, Limits::default()).unwrap();
            assert_eq!(archive.entries()[0].compressed_size(), 5);
            assert_eq!(archive.extract(&archive.entries()[0]).unwrap(), b"hello");
        }
    }

    #[test]
    fn rejects_inconsistent_or_oversized_zip64_metadata() {
        let name = b"safe.xml";

        let mut mismatched_local = one_entry_zip64(name, b"hello", false);
        let local_uncompressed = 30 + name.len() + 4;
        mismatched_local[local_uncompressed..local_uncompressed + 8]
            .copy_from_slice(&6_u64.to_le_bytes());
        assert_eq!(
            ZipArchive::parse(&mismatched_local, Limits::default())
                .unwrap_err()
                .code,
            DiagnosticCode::ZipInvalid,
        );

        let mut oversized = one_entry_zip64(name, b"hello", false);
        let directory = 30 + name.len() + 20 + 5;
        let central_uncompressed = directory + 46 + name.len() + 4;
        oversized[central_uncompressed..central_uncompressed + 8].copy_from_slice(
            &(Limits::default().max_entry_uncompressed_bytes as u64 + 1).to_le_bytes(),
        );
        assert_eq!(
            ZipArchive::parse(&oversized, Limits::default())
                .unwrap_err()
                .code,
            DiagnosticCode::ZipEntryTooLarge,
        );

        let mut bad_locator = one_entry_zip64(name, b"hello", false);
        let eocd = bad_locator.len() - 22;
        let locator = eocd - 20;
        bad_locator[locator + 8..locator + 16].copy_from_slice(&0_u64.to_le_bytes());
        assert_eq!(
            ZipArchive::parse(&bad_locator, Limits::default())
                .unwrap_err()
                .code,
            DiagnosticCode::ZipInvalid,
        );

        let mut disagreeing_eocd = one_entry_zip64(name, b"hello", false);
        let eocd = disagreeing_eocd.len() - 22;
        disagreeing_eocd[eocd + 10..eocd + 12].copy_from_slice(&0_u16.to_le_bytes());
        assert_eq!(
            ZipArchive::parse(&disagreeing_eocd, Limits::default())
                .unwrap_err()
                .code,
            DiagnosticCode::ZipInvalid,
        );

        let mut multi_disk = one_entry_zip64(name, b"hello", false);
        let locator = multi_disk.len() - 22 - 20;
        multi_disk[locator + 16..locator + 20].copy_from_slice(&2_u32.to_le_bytes());
        assert_eq!(
            ZipArchive::parse(&multi_disk, Limits::default())
                .unwrap_err()
                .code,
            DiagnosticCode::ZipMultiDiskForbidden,
        );
    }

    #[test]
    fn iwork_compatibility_accepts_only_valid_unflagged_utf8_paths() {
        let name = "Data/unicode-骰子.png".as_bytes();
        let unflagged = one_entry_zip(name, b"image", b"image", 0, 0);
        assert_eq!(
            ZipArchive::parse(&unflagged, Limits::default())
                .unwrap_err()
                .code,
            DiagnosticCode::ZipInvalid,
        );
        let archive = ZipArchive::parse_iwork(&unflagged, Limits::default()).unwrap();
        assert_eq!(archive.entries()[0].name(), "Data/unicode-骰子.png");

        let invalid_name = b"Data/invalid-\xff.png";
        let invalid = one_entry_zip(invalid_name, b"image", b"image", 0, 0);
        assert_eq!(
            ZipArchive::parse_iwork(&invalid, Limits::default())
                .unwrap_err()
                .code,
            DiagnosticCode::ZipInvalid,
        );
    }

    #[test]
    fn accepts_valid_utf8_when_only_local_and_central_utf8_flags_differ() {
        let name = b"Doc_0/Document.xml";
        let mut zip = one_entry_zip(name, b"ofd", b"ofd", 0, super::FLAG_UTF8);
        zip[6..8].copy_from_slice(&0_u16.to_le_bytes());
        let archive = ZipArchive::parse(&zip, Limits::default()).unwrap();
        assert_eq!(archive.entries()[0].name(), "Doc_0/Document.xml");

        let invalid_name = b"Doc_0/invalid-\xff.xml";
        let mut invalid = one_entry_zip(invalid_name, b"ofd", b"ofd", 0, super::FLAG_UTF8);
        invalid[6..8].copy_from_slice(&0_u16.to_le_bytes());
        assert_eq!(
            ZipArchive::parse(&invalid, Limits::default())
                .unwrap_err()
                .code,
            DiagnosticCode::ZipInvalid,
        );
    }

    #[test]
    fn accepts_deflated_empty_directory_entries() {
        let zip = one_entry_zip(b"META-INF/", b"", &[0x03, 0x00], 8, 0);
        let archive = ZipArchive::parse(&zip, Limits::default())
            .expect("a Deflate stream may encode an empty directory payload");

        assert!(archive.entries()[0].is_directory());
        archive.validate_all().expect("the empty stream is valid");
    }

    #[test]
    fn deflate_operation_budget_is_shared_by_all_extractions() {
        let compressed = [0xcb, 0x48, 0xcd, 0xc9, 0xc9, 0x07, 0x00];
        let deflated = one_entry_zip(b"content.xml", b"hello", &compressed, 8, 0);
        let limits = Limits {
            max_deflate_operations: 500,
            ..Limits::default()
        };
        let archive = ZipArchive::parse(&deflated, limits).unwrap();
        let entry = &archive.entries()[0];

        assert_eq!(archive.extract(entry).unwrap(), b"hello");
        let error = archive.extract(entry).unwrap_err();
        assert_eq!(error.code, DiagnosticCode::ZipDeflateInvalid);
        assert!(error.message.contains("operation"));
    }

    #[test]
    fn iwork_expanded_byte_budget_does_not_change_standard_zip_behavior() {
        let stored = one_entry_zip(b"content.xml", b"hello", b"hello", 0, 0);
        let limits = Limits {
            max_entry_uncompressed_bytes: 5,
            max_total_uncompressed_bytes: 5,
            max_xml_bytes: 5,
            ..Limits::default()
        };

        let standard = ZipArchive::parse(&stored, limits).unwrap();
        let entry = &standard.entries()[0];
        assert_eq!(standard.extract(entry).unwrap(), b"hello");
        assert_eq!(standard.extract(entry).unwrap(), b"hello");

        let iwork = ZipArchive::parse_iwork(&stored, limits).unwrap();
        let entry = &iwork.entries()[0];
        assert_eq!(iwork.extract(entry).unwrap(), b"hello");
        assert_eq!(
            iwork.extract(entry).unwrap_err().code,
            DiagnosticCode::ZipTotalSizeLimit
        );
    }

    #[test]
    fn validates_signed_and_unsigned_data_descriptors() {
        let name = b"word/document.xml";
        let signed = one_entry_zip(name, b"hello", b"hello", 0, super::FLAG_DATA_DESCRIPTOR);
        ZipArchive::parse(&signed, Limits::default())
            .unwrap()
            .validate_all()
            .unwrap();

        let descriptor = 30 + name.len() + 5;
        let mut unsigned = signed;
        unsigned.drain(descriptor..descriptor + 4);
        let eocd = unsigned.len() - 22;
        let old_directory_offset =
            u32::from_le_bytes(unsigned[eocd + 16..eocd + 20].try_into().unwrap());
        unsigned[eocd + 16..eocd + 20].copy_from_slice(&(old_directory_offset - 4).to_le_bytes());
        ZipArchive::parse(&unsigned, Limits::default())
            .unwrap()
            .validate_all()
            .unwrap();

        let name = b"word/document.xml";
        let mut redundant_central_flag = one_entry_zip(name, b"hello", b"hello", 0, 0);
        let directory = 30 + name.len() + 5;
        redundant_central_flag[directory + 8..directory + 10]
            .copy_from_slice(&super::FLAG_DATA_DESCRIPTOR.to_le_bytes());
        ZipArchive::parse(&redundant_central_flag, Limits::default())
            .expect("an unused central descriptor bit is tolerated when local sizes are exact")
            .validate_all()
            .unwrap();

        let mut stale_local_crc = one_entry_zip(name, b"hello", b"hello", 0, 0);
        stale_local_crc[14..18].copy_from_slice(&0_u32.to_le_bytes());
        ZipArchive::parse(&stale_local_crc, Limits::default())
            .expect("the central CRC is authoritative")
            .validate_all()
            .unwrap();
    }

    #[test]
    fn rejects_traversal_encryption_incomplete_zip64_and_crc_mismatch() {
        let traversal = one_entry_zip(b"../evil.xml", b"x", b"x", 0, 0);
        assert_eq!(
            ZipArchive::parse(&traversal, Limits::default())
                .unwrap_err()
                .code,
            DiagnosticCode::ZipPathTraversal
        );

        let backslash = one_entry_zip(b"word\\document.xml", b"x", b"x", 0, 0);
        let archive = ZipArchive::parse(&backslash, Limits::default()).unwrap();
        assert_eq!(archive.entries()[0].name(), "word/document.xml");

        let backslash_traversal = one_entry_zip(b"..\\evil.xml", b"x", b"x", 0, 0);
        assert_eq!(
            ZipArchive::parse(&backslash_traversal, Limits::default())
                .unwrap_err()
                .code,
            DiagnosticCode::ZipPathTraversal
        );

        let encrypted = one_entry_zip(b"safe.xml", b"x", b"x", 0, super::FLAG_ENCRYPTED);
        assert_eq!(
            ZipArchive::parse(&encrypted, Limits::default())
                .unwrap_err()
                .code,
            DiagnosticCode::ZipEncrypted
        );

        let mut zip64 = one_entry_zip(b"safe.xml", b"x", b"x", 0, 0);
        let eocd = zip64.len() - 22;
        zip64[eocd + 10..eocd + 12].copy_from_slice(&u16::MAX.to_le_bytes());
        assert_eq!(
            ZipArchive::parse(&zip64, Limits::default())
                .unwrap_err()
                .code,
            DiagnosticCode::ZipInvalid
        );

        let mut bad_crc = one_entry_zip(b"safe.xml", b"x", b"x", 0, 0);
        bad_crc[14..18].copy_from_slice(&0_u32.to_le_bytes());
        let directory = 30 + b"safe.xml".len() + 1;
        bad_crc[directory + 16..directory + 20].copy_from_slice(&0_u32.to_le_bytes());
        let archive = ZipArchive::parse(&bad_crc, Limits::default()).unwrap();
        assert_eq!(
            archive.validate_all().unwrap_err().code,
            DiagnosticCode::ZipCrcMismatch
        );
    }

    #[test]
    fn eocd_signature_in_comment_is_not_accepted() {
        let mut zip = one_entry_zip(b"safe.xml", b"x", b"x", 0, 0);
        let eocd = zip.len() - 22;
        zip[eocd + 20..eocd + 22].copy_from_slice(&4_u16.to_le_bytes());
        zip.extend_from_slice(b"PK\x05\x06");
        let archive = ZipArchive::parse(&zip, Limits::default()).unwrap();
        archive.validate_all().unwrap();
    }
}
