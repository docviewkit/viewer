# Pages external template images and dependency thumbnails

The unchanged supplied invitation is retained as
`tests/fixtures/pages-external-template-images.pages`, SHA-256
`92f21fb951d2fc0bcee9e59f92ad4bd93c0fd43b48d226ae67ce7a578a75ffbb`.
It contains three native ImageArchives (1360935, 1360937, 1360936), whose original
resources 15, 17 and 18 are external Pages template locators. Their metadata file
names are empty. The only packaged JPEG is resource 16, the first image's
thumbnail. Native Pages obtains the three original photos from its installed
resources; a browser cannot reconstruct absent original bytes from this package.

The shared drawable renderer correctly parsed only original/adjusted image data
fields, but then appended every archive-header dependency as a fallback. That
bypassed the thumbnail exclusion and rendered resource 16 as final content. The
renderer now uses only the image's declared original/adjusted resource IDs. The
same implementation serves Pages, Numbers and Keynote drawables; no parallel
adapter implementation was added. Existing image type/signature, crop, mask,
bounds and resource-limit behavior is unchanged. Unavailable originals retain
their frames, warning diagnostics and the remaining useful document content.

The unchanged-file regression fails before this fix because one image frame
contains a thumbnail. After the fix it retains three empty frames, three
IWORK_IMAGE_UNAVAILABLE diagnostics, both pages and the invitation text. The
existing shared parser test also excludes thumbnail reference fields 12 and 16.
This case verifies honest local degradation, not automatic external-asset recovery.

A separate, local-only recovered copy was created at the user's request context:
`/Users/wjfree/Downloads/Untitled-with-images.pages`. Native Pages exported the
original photos into a temporary DOCX. Each recovered JPEG's SHA-1 exactly matched
the original Pages metadata before insertion into a new copy:

- 15: 8bc249a448bca162690da62f4d58992324d5096f (161166 bytes)
- 17: a8517901bc612e418d3182cedfed51a3b1088b15 (51020 bytes)
- 18: 5aaa0973a97f54617c64d2b0f2d525a5464f8f19 (79174 bytes)

Only resource file names/materialized lengths in Index/Metadata.iwa and the three
new Data entries changed. Every other original archive entry stayed byte-identical;
Core page, text, shape and image geometry fields stayed identical. The recovered
copy renders three real images in Chromium and WebKit. It is not a converted Word
rendering, a page-preview fallback, or a bundled Apple template asset catalog.
Recovered external assets are not shipped with the SDK. Recovery evidence is in
`output/playwright/pages-floating-image/`.
