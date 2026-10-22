# Pages brochure: structural text marker and emphasis

The supplied `Untitled.pages` is retained unchanged as
`tests/fixtures/pages-brochure.pages`, SHA-256
`c3d971477eaea76a687716d3418c4da9a94938bdf7ee084bcaab8a897654feb9`.
The two-page layout document retains a body story (storage 33845) containing only
U+0004. Native Pages does not display this marker, but the old parser emitted a
one-character text object and exposed it through search/text fragments.

The common iWork story normalizer now removes U+0004 after resolving saved UTF-16
ranges and expanding attachments. Raw storage and style offsets are unchanged.
The existing U+2028 soft-break and U+FFFC attachment handling remains intact; tabs,
carriage returns, newlines and other characters retain their prior behavior.
Pages body ranges and shared TSD text boxes use this normalizer. Table text has a
separate, broader control-filtering contract and retains its existing behavior.

Storage 33703 explicitly applies character style 32924 (`Emphasis`, bold=true) to
both `Subheading` paragraphs. Current baseline parsing and Chromium/WebKit drawing
already preserve this emphasis. This change adds real-file coverage for both bold
headings and regular body text; it does not claim a newly reproduced weight bug.
No font substitution, font file or weight behavior was changed.

```sh
node scripts/test-pages-brochure.mjs
PAGES_BROCHURE_BASELINE_WASM=output/playwright/pages-brochure/before.wasm PAGES_BROCHURE_OUTPUT=output/playwright/pages-brochure/baseline node scripts/test-pages-brochure.mjs
```

The old pack fails only the story-marker check; emphasis checks already pass.
The rebuilt pack passes marker absence, heading fonts/ink and regular-body checks
in Chromium and WebKit without console errors. Core tests cover the unchanged
file and shared Pages/Keynote UTF-16 style boundaries around markers and a
supplementary-plane character. The iWork suite passes 122 cases; selected
Keynote/Numbers scene hashes remain unchanged. The reviewed Pages baseline changes
only by removing U+0004 from two stories and one now-empty run; every other scene
field, including bounds and remaining font styles, is identical. The PDF/iWork pack grows from
3,219,656 to 3,219,964 bytes (+308; 0.01%), within the existing budget. Captures and
logs are under `output/playwright/pages-brochure/`.
