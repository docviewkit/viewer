# DocViewKit commercial design system

## Direction

DocViewKit uses a restrained developer-infrastructure aesthetic: trustworthy, precise, local-first, and content-led. The interface must look established without pretending that unverified customer proof exists.

- Style: trust and authority, editorial minimalism, flat surfaces.
- Avoid: AI purple gradients, glass decoration, fake testimonials, oversized rounded cards, decorative motion, emoji icons.
- Primary action: one dominant CTA per screen.

## Color tokens

| Token | Light | Dark | Use |
| --- | --- | --- | --- |
| `--surface-page` | `#f7f7f5` | `#0d1117` | Paper-like page background |
| `--surface` | `#ffffff` | `#151b26` | Cards and panels |
| `--surface-subtle` | `#eceef2` | `#202838` | Secondary regions |
| `--text` | `#111827` | `#f3f4f6` | Ink-like primary text |
| `--text-muted` | `#5d6678` | `#a8b0bf` | Secondary text |
| `--line` | `#d7dbe2` | `#30394a` | Borders and dividers |
| `--brand` | `#2557d6` | `#7aa2ff` | Cobalt primary action and active state |
| `--brand-strong` | `#1746b8` | `#a7c0ff` | Hover and high emphasis |
| `--warm` | `#d95d39` | `#ff8a66` | Source location and limited conversion accent |
| `--danger` | `#b42318` | `#ff8a7d` | Destructive state |

All functional colors include text or icon reinforcement. Body text contrast must meet WCAG AA.

## Typography

- UI and prose: `Inter`, `IBM Plex Sans`, `-apple-system`, `BlinkMacSystemFont`, `Segoe UI`, sans-serif.
- Code and technical labels: `JetBrains Mono`, `SFMono-Regular`, `Consolas`, monospace.
- No external font request: use the local stack to preserve privacy and first render performance.
- Body: 16px minimum, 1.6 line height. Long-form text: 65–75 characters.
- Scale: 12 / 14 / 16 / 18 / 24 / 32 / 48 / 64.

## Layout and interaction

- Spacing uses a 4/8px scale.
- Content widths: 1180px marketing, 1440px docs shell.
- Breakpoints: 480 / 768 / 1024 / 1280px.
- All controls have a minimum 44px target.
- Focus ring: 3px brand tint plus 2px offset.
- Motion: 160–240ms, transform/opacity only, disabled under `prefers-reduced-motion`.
- Radius: 6px controls, 10px panels, 14px feature surfaces. Avoid pill-shaped containers except compact badges.
- Shadow: one restrained elevation scale; prefer borders over floating cards.

## Page patterns

### Landing

Hero → evidence strip → integration proof → source-location value → issue loop → plans → FAQ → final CTA. No invented logos or claims.

### Documentation

Persistent product header, searchable left navigation, readable article column, optional on-page outline. Every guide has a stable URL and a plain-text equivalent through `/llms.txt` and `/llms-full.txt`.
