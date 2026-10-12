# XPS directory reference regression

These files are byte-for-byte copies of the user-supplied files in
`.cache/xps-tests/directory-reference`, including the existing `control.xps`.
All four contain the same real page and embedded font, with text `This is a test`.

- `traversal.xps`: the sole document reference escapes the package root. Reject that URI and recover only from a unique declared document inside the validated package.
- `none.xps`: the sole reference names an empty/missing directory. Recover from the unique declared document.
- `ambiguous.xps`: the referenced directory contains byte-identical document indexes with the same relative URI base. Recover one index, without duplicating the page.
- `control.xps`: a valid explicit document reference; content and rendering must stay unchanged.

Recovery emits warnings. Distinct document contents, different URI bases, unsafe ZIP members, unsafe page references, and forbidden XML remain rejected. This index recovery belongs to the XPS adapter: OFD and OOXML have different document indexes and ownership/order rules. All retain the shared package URI resolver and bounded ZIP/XML validation.
