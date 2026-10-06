//! Resource limits applied before and during untrusted document parsing.

/// Runtime limits. Callers may lower these values but may not exceed [`Limits::HARD_MAX`].
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct Limits {
    pub max_input_bytes: usize,
    pub max_zip_entries: usize,
    pub max_zip_path_bytes: usize,
    pub max_entry_uncompressed_bytes: usize,
    pub max_total_uncompressed_bytes: usize,
    pub max_compression_ratio: u32,
    pub compression_ratio_min_bytes: usize,
    pub max_deflate_operations: usize,
    pub max_xml_bytes: usize,
    pub max_xml_depth: usize,
    pub max_xml_nodes: usize,
    pub max_xml_attributes_per_element: usize,
    pub max_xml_name_bytes: usize,
    pub max_relationship_edges: usize,
    pub max_document_objects: usize,
    pub max_render_pixels: usize,
    pub max_font_bytes: usize,
}

impl Limits {
    pub const MIB: usize = 1024 * 1024;

    /// Absolute ceilings enforced even when a host supplies custom limits.
    pub const HARD_MAX: Self = Self {
        max_input_bytes: 512 * Self::MIB,
        max_zip_entries: 50_000,
        max_zip_path_bytes: 64 * 1024,
        max_entry_uncompressed_bytes: 256 * Self::MIB,
        max_total_uncompressed_bytes: 1024 * Self::MIB,
        max_compression_ratio: 1_000,
        compression_ratio_min_bytes: 16 * Self::MIB,
        max_deflate_operations: 1_000_000_000,
        max_xml_bytes: 256 * Self::MIB,
        max_xml_depth: 256,
        max_xml_nodes: 10_000_000,
        max_xml_attributes_per_element: 1_024,
        max_xml_name_bytes: 16 * 1024,
        max_relationship_edges: 200_000,
        max_document_objects: 4_000_000,
        max_render_pixels: 256_000_000,
        max_font_bytes: 256 * Self::MIB,
    };

    /// Rejects disabled limits and values above the hard security ceilings.
    pub fn validate(&self) -> Result<(), LimitConfigError> {
        macro_rules! bounded {
            ($field:ident) => {
                if self.$field == 0 || self.$field > Self::HARD_MAX.$field {
                    return Err(LimitConfigError::new(stringify!($field)));
                }
            };
        }

        bounded!(max_input_bytes);
        bounded!(max_zip_entries);
        bounded!(max_zip_path_bytes);
        bounded!(max_entry_uncompressed_bytes);
        bounded!(max_total_uncompressed_bytes);
        bounded!(max_compression_ratio);
        bounded!(compression_ratio_min_bytes);
        bounded!(max_deflate_operations);
        bounded!(max_xml_bytes);
        bounded!(max_xml_depth);
        bounded!(max_xml_nodes);
        bounded!(max_xml_attributes_per_element);
        bounded!(max_xml_name_bytes);
        bounded!(max_relationship_edges);
        bounded!(max_document_objects);
        bounded!(max_render_pixels);
        if self.max_font_bytes > Self::HARD_MAX.max_font_bytes {
            return Err(LimitConfigError::new("max_font_bytes"));
        }

        if self.max_entry_uncompressed_bytes > self.max_total_uncompressed_bytes {
            return Err(LimitConfigError::new(
                "max_entry_uncompressed_bytes exceeds max_total_uncompressed_bytes",
            ));
        }
        if self.max_xml_bytes > self.max_entry_uncompressed_bytes {
            return Err(LimitConfigError::new(
                "max_xml_bytes exceeds max_entry_uncompressed_bytes",
            ));
        }
        Ok(())
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_input_bytes: 128 * Self::MIB,
            max_zip_entries: 20_000,
            max_zip_path_bytes: 4 * 1024,
            max_entry_uncompressed_bytes: 128 * Self::MIB,
            max_total_uncompressed_bytes: 512 * Self::MIB,
            max_compression_ratio: 200,
            compression_ratio_min_bytes: Self::MIB,
            max_deflate_operations: 200_000_000,
            max_xml_bytes: 64 * Self::MIB,
            max_xml_depth: 128,
            max_xml_nodes: 2_000_000,
            max_xml_attributes_per_element: 256,
            max_xml_name_bytes: 4 * 1024,
            max_relationship_edges: 50_000,
            max_document_objects: 1_000_000,
            max_render_pixels: 32_000_000,
            max_font_bytes: 64 * Self::MIB,
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct LimitConfigError {
    field: &'static str,
}

impl LimitConfigError {
    const fn new(field: &'static str) -> Self {
        Self { field }
    }

    pub const fn field(self) -> &'static str {
        self.field
    }
}

impl core::fmt::Display for LimitConfigError {
    fn fmt(&self, formatter: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
        write!(formatter, "invalid resource limit: {}", self.field)
    }
}

impl std::error::Error for LimitConfigError {}

#[cfg(test)]
mod tests {
    use super::Limits;

    #[test]
    fn defaults_are_valid() {
        Limits::default().validate().unwrap();
    }

    #[test]
    fn zero_or_excessive_limits_are_rejected() {
        let limits = Limits {
            max_zip_entries: 0,
            ..Limits::default()
        };
        assert_eq!(limits.validate().unwrap_err().field(), "max_zip_entries");

        let limits = Limits {
            max_xml_depth: Limits::HARD_MAX.max_xml_depth + 1,
            ..Limits::default()
        };
        assert_eq!(limits.validate().unwrap_err().field(), "max_xml_depth");
    }
}
