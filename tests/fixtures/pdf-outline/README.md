# Preview outline presentation probes

Observed in macOS 27.0.1, Preview 11.0 on 2026-10-16. These small PDFs were
created with pypdf: three blank 300 x 200 pt pages and explicit outline trees.
They contain no customer document content. Node destinations use zero-based
page indices; all nodes have ordinary GoTo destinations.

Preview always provides a filename row. When the authored outline has exactly
one top-level item, that item is represented by the filename row and its children
are displayed directly underneath. This happens once, not recursively. The
label, number of children, and destination page do not determine the behavior.
PDFKit's `PDFDocument.outlineRoot` still returns the authored tree.

| Probe | Authored outline | Preview rows below filename |
| --- | --- | --- |
| 01 | Arbitrary root(0) -> Child A(0), Child B(1) | Child A, Child B |
| 02 | Root A(0) -> Child A(0); Root B(1) -> Child B(1) | Root A, Root B (expandable) |
| 03 | Outer(0) -> Middle(0) -> Inner(0) -> Child A(0), Child B(1) | Middle (expandable; only Outer absorbed) |
| 04 | Root targets page 3(2) -> Child A(0), Child B(1) | Child A, Child B |
| 05 | Only bookmark(1) | None; filename row remains |
| 06 | Parent(0) -> Only child(1) | Only child |
| 07 | Outer(0) -> Section A(0) -> Detail A(0); Section B(1) -> Detail B(1) | Section A, Section B; expanding A still shows Detail A |
| 08 | Default Theme(0) -> Child A(0); Other root(1) | Default Theme, Other root |

Viewer applies this structural simplification in its shared presentation layer,
without changing `DocumentInfo.outline` or format parsers. It omits one sole
parent and adjusts the displayed levels and navigation indices. A lone leaf is
retained: Viewer has no clickable filename tree row to preserve its destination.
Multiple roots and nested sections retain their titles and hierarchy, including
ordinary sections named Default Theme.

The real reported PDF has 33 pages and `Default Theme -> Slide 1 ... Slide 33`.
Its SHA-256 is `da9a2f04b20035cb0d6c9235c5ba8da98f7e6a934a3074885735d1ea390be509`.
It is not redistributed. Pass its local path to run the original-file regression:

```sh
npm run build
node scripts/run-viewer-outline-regression.mjs /path/to/WNCDKP2_Digital_Key_Proposal_Phase2_v1.3.pdf
```

Without the argument the eight probes and two existing real DOCX fixtures run.
Both modes use Chromium and WebKit, real Workers and parsers, and check source
outline preservation, displayed rows, indentation, ARIA hierarchy, navigation,
collapse/expand, keyboard focus, mode switching, reopening, and browser errors.

Cross-format boundary: PDF `/Outlines`, DOCX headings, and OFD outline elements
have different extraction rules and remain in their adapters. The single-root
presentation decision is applied once in the shared Viewer, after parsing. No
format parser needs a copy of this rule.
