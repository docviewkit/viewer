//! Stable diagnostics shared by container and XML parsing.

#[repr(u32)]
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum DiagnosticCode {
    None = 0,
    InvalidLimits = 1,
    InputTooLarge = 2,
    AllocationFailed = 3,
    UnsupportedFormat = 4,
    ZipInvalid = 100,
    ZipMultiDiskForbidden = 101,
    Zip64Forbidden = 102,
    ZipEncrypted = 103,
    ZipUnsupportedCompression = 104,
    ZipEntryLimit = 105,
    ZipEntryTooLarge = 106,
    ZipTotalSizeLimit = 107,
    ZipCompressionRatioLimit = 108,
    ZipPathTraversal = 109,
    ZipDuplicateEntry = 110,
    ZipCrcMismatch = 111,
    ZipDeflateInvalid = 112,
    XmlInvalid = 200,
    XmlEncodingUnsupported = 201,
    XmlDtdForbidden = 202,
    XmlEntityForbidden = 203,
    XmlDepthLimit = 204,
    XmlNodeLimit = 205,
    XmlAttributeLimit = 206,
    XmlSizeLimit = 207,
    ActiveContentBlocked = 300,
    ExternalResourceBlocked = 301,
    UnsupportedFeature = 302,
    FormatInvalid = 303,
    RelationshipLimit = 304,
    ObjectLimit = 305,
    ImageDimensionLimit = 306,
    LayoutBudgetExceeded = 307,
    EmbeddedFontInvalid = 308,
    FontBytesLimit = 309,
    ApproximateLayout = 310,
    PdfPasswordRequired = 311,
    PdfPasswordIncorrect = 312,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum Severity {
    Info,
    Warning,
    Error,
    Fatal,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum Phase {
    Identify,
    Input,
    Container,
    Xml,
    Parse,
    Layout,
    Render,
    Security,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum Fidelity {
    Exact,
    Approximate,
    Unsupported,
    Omitted,
    Blocked,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct DiagnosticLocation {
    pub byte_offset: Option<usize>,
    pub part: Option<String>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct Diagnostic {
    pub code: DiagnosticCode,
    pub severity: Severity,
    pub phase: Phase,
    pub fidelity: Fidelity,
    pub location: DiagnosticLocation,
    pub message: String,
    pub details: Vec<(String, String)>,
}

impl Diagnostic {
    pub fn warning(
        code: DiagnosticCode,
        phase: Phase,
        fidelity: Fidelity,
        message: impl Into<String>,
    ) -> Self {
        Self {
            code,
            severity: Severity::Warning,
            phase,
            fidelity,
            location: DiagnosticLocation {
                byte_offset: None,
                part: None,
            },
            message: message.into(),
            details: Vec::new(),
        }
    }

    pub fn fatal(
        code: DiagnosticCode,
        phase: Phase,
        byte_offset: Option<usize>,
        message: impl Into<String>,
    ) -> Self {
        Self {
            code,
            severity: Severity::Fatal,
            phase,
            fidelity: Fidelity::Blocked,
            location: DiagnosticLocation {
                byte_offset,
                part: None,
            },
            message: message.into(),
            details: Vec::new(),
        }
    }

    pub fn in_part(mut self, part: impl Into<String>) -> Self {
        self.location.part = Some(part.into());
        self
    }

    pub fn with_detail(mut self, key: impl Into<String>, value: impl Into<String>) -> Self {
        self.details.push((key.into(), value.into()));
        self
    }
}

impl core::fmt::Display for Diagnostic {
    fn fmt(&self, formatter: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
        write!(formatter, "{}", self.message)
    }
}

impl std::error::Error for Diagnostic {}
