import { DEFAULT_LIMITS } from "./core.js";
import { createOfficeEngine } from "./engine.js";
import { sheetAutoFitColumnWidth, sheetAutoFitRowHeight, sheetPrintPages } from "./render.js";
import { OPEN_SOURCE_LICENSE, type LicenseOptions, type LicenseState } from "./license.js";
import { OfficeEngineError } from "./types.js";
import type {
  Diagnostic,
  DocumentAction,
  DocumentInfo,
  DocumentObject,
  EngineOptions,
  HitResult,
  OfficeDocument,
  OfficeEngine,
  OpenOptions,
  Rect,
  RenderRequest,
  RenderResult,
  RenderedTextFragment,
  SheetAxis,
  TextSearchResult,
  UnitDescriptor,
} from "./types.js";

export type ViewerTheme = "auto" | "light" | "dark";
export type ViewerNavigation = "auto" | "visible" | "hidden";
export type ViewerPageMode = "single" | "continuous";
type ViewerNavigationMode = "outline" | "thumbnails";
export type ViewerInteractionMode = "object" | "display" | "text";
export type ViewerStatus = "idle" | "loading" | "ready" | "error" | "destroyed";

export interface ViewerFeatures {
  readonly search: boolean;
  readonly navigation: boolean;
  readonly zoom: boolean;
  readonly print: boolean;
  readonly fullscreen: boolean;
  readonly diagnostics: boolean;
  readonly pageMode: boolean;
  /** Shows the built-in object, display, and text interaction mode switcher. */
  readonly interactionModeSwitcher: boolean;
  /** Controls mutually exclusive object location, display-only, and text selection behavior. */
  readonly interactionMode: ViewerInteractionMode;
  /** @deprecated Use interactionMode: "text". */
  readonly textSelection?: boolean;
  /** @deprecated Use interactionMode: "object" or interactionMode: "display". */
  readonly objectSelection?: boolean;
  /** Activates safe document hyperlinks and internal navigation actions. */
  readonly hyperlinks: boolean;
}

export interface ViewerMessages {
  readonly navigation: string;
  readonly previous: string;
  readonly next: string;
  readonly search: string;
  readonly searchPlaceholder: string;
  readonly previousResult: string;
  readonly nextResult: string;
  readonly closeSearch: string;
  readonly zoomOut: string;
  readonly zoomIn: string;
  readonly zoom: string;
  readonly fit: string;
  readonly print: string;
  readonly fullscreen: string;
  readonly exitFullscreen: string;
  readonly themeAuto: string;
  readonly themeLight: string;
  readonly themeDark: string;
  readonly diagnostics: string;
  readonly continuousPages: string;
  readonly interactionModes: string;
  readonly objectMode: string;
  readonly displayMode: string;
  readonly textMode: string;
  readonly close: string;
  readonly loading: string;
  readonly empty: string;
  readonly emptyHint: string;
  readonly unsupported: string;
  readonly renderFailed: string;
  readonly noResults: string;
  readonly resultPosition: string;
  readonly pagePosition: string;
  readonly approximateLocation: string;
  readonly exactLocation: string;
  readonly noDiagnostics: string;
  readonly documentCanvas: string;
  readonly documentNavigation: string;
  readonly resizeNavigation: string;
  readonly outline: string;
  readonly thumbnails: string;
  readonly sheetTabs: string;
  readonly formulaBar: string;
  readonly cellAddress: string;
  readonly columnHeaders: string;
  readonly rowHeaders: string;
  readonly resizeColumn: string;
  readonly resizeRow: string;
  readonly passwordRequired: string;
  readonly passwordIncorrect: string;
  readonly passwordLabel: string;
  readonly unlock: string;
  readonly cancel: string;
}

export interface ViewerConfig {
  readonly theme?: ViewerTheme;
  readonly locale?: string;
  readonly messages?: Readonly<Record<string, Partial<ViewerMessages>>>;
  readonly features?: Partial<ViewerFeatures>;
  readonly navigation?: ViewerNavigation;
  readonly initialPageMode?: ViewerPageMode;
  readonly initialZoom?: "fit" | number;
  /** Hides all Viewer chrome and leaves only the rendering workspace. */
  readonly minimal?: boolean;
  /** Repeated text included in surfaces, thumbnails, and printing. */
  readonly watermark?: string;
  readonly engine?: EngineOptions;
  /** @deprecated Ignored; all capabilities are available under Apache-2.0. */
  readonly license?: LicenseOptions;
}

export type ViewerOpenOptions = Pick<OpenOptions, "password">;

export type ViewerTarget =
  | { readonly kind: "unit"; readonly unitIndex: number }
  | { readonly kind: "object"; readonly objectId: string }
  | { readonly kind: "region"; readonly unitIndex: number; readonly region: Rect }
  | {
      readonly kind: "source";
      readonly unitIndex?: number;
      readonly objectId?: string;
      readonly region?: Rect;
    };

export interface RevealResult {
  readonly unitIndex: number;
  readonly quality: "exact" | "approximate";
  readonly objectId?: string;
  readonly region?: Rect;
}

export interface ViewerState {
  readonly status: ViewerStatus;
  readonly info?: DocumentInfo;
  readonly currentUnitIndex: number;
  readonly unitCount: number;
  readonly zoom: number;
  readonly locale: string;
  readonly theme: "light" | "dark";
  readonly pageMode: ViewerPageMode;
  readonly license: LicenseState;
  readonly search: Readonly<{
    query: string;
    current: number;
    total: number;
  }>;
}

const DEFAULT_FEATURES: ViewerFeatures = Object.freeze({
  search: true,
  navigation: true,
  zoom: true,
  print: true,
  fullscreen: true,
  diagnostics: false,
  pageMode: true,
  interactionModeSwitcher: false,
  interactionMode: "object",
  hyperlinks: true,
});

const DEFAULT_PAGE_MODE: ViewerPageMode = "continuous";
const MIN_ZOOM = .25;
const MAX_ZOOM = 4;
const PRINT_RENDER_LONG_EDGE = 1600;

const EN: ViewerMessages = Object.freeze({
  navigation: "Toggle document navigation",
  previous: "Previous page",
  next: "Next page",
  search: "Search document",
  searchPlaceholder: "Search",
  previousResult: "Previous result",
  nextResult: "Next result",
  closeSearch: "Close search",
  zoomOut: "Zoom out",
  zoomIn: "Zoom in",
  zoom: "Zoom percentage",
  fit: "Fit to view",
  print: "Print document",
  fullscreen: "Enter fullscreen",
  exitFullscreen: "Exit fullscreen",
  themeAuto: "Theme: System. Switch to Light",
  themeLight: "Theme: Light. Switch to Dark",
  themeDark: "Theme: Dark. Follow system",
  diagnostics: "Compatibility diagnostics",
  continuousPages: "Continuous pages",
  interactionModes: "Viewer interaction mode",
  objectMode: "Locate",
  displayMode: "View",
  textMode: "Text",
  close: "Close",
  loading: "Opening document…",
  empty: "No document open",
  emptyHint: "The host application opens a local document here.",
  unsupported: "This document cannot be displayed.",
  renderFailed: "The page could not be rendered.",
  noResults: "No results",
  resultPosition: "{current} of {total}",
  pagePosition: "{current} / {total}",
  approximateLocation: "The closest available location is shown.",
  exactLocation: "Source location revealed.",
  noDiagnostics: "No compatibility diagnostics",
  documentCanvas: "Document canvas",
  documentNavigation: "Document navigation",
  resizeNavigation: "Resize document navigation",
  outline: "Document outline",
  thumbnails: "Page thumbnails",
  sheetTabs: "Workbook sheets",
  formulaBar: "Formula bar",
  cellAddress: "Cell address",
  columnHeaders: "Column headers",
  rowHeaders: "Row headers",
  resizeColumn: "Resize column {label}",
  resizeRow: "Resize row {label}",
  passwordRequired: "Enter the password to open this PDF.",
  passwordIncorrect: "Incorrect password. Try again.",
  passwordLabel: "PDF password",
  unlock: "Open",
  cancel: "Cancel",
});

const ZH_CN: ViewerMessages = Object.freeze({
  navigation: "切换文档导航",
  previous: "上一页",
  next: "下一页",
  search: "搜索文档",
  searchPlaceholder: "搜索",
  previousResult: "上一个结果",
  nextResult: "下一个结果",
  closeSearch: "关闭搜索",
  zoomOut: "缩小",
  zoomIn: "放大",
  zoom: "缩放百分比",
  fit: "适应窗口",
  print: "打印文档",
  fullscreen: "进入全屏",
  exitFullscreen: "退出全屏",
  themeAuto: "主题：跟随系统。切换为浅色",
  themeLight: "主题：浅色。切换为深色",
  themeDark: "主题：深色。切换为跟随系统",
  diagnostics: "兼容性诊断",
  continuousPages: "连续页面",
  interactionModes: "Viewer 交互模式",
  objectMode: "对象定位",
  displayMode: "纯显示",
  textMode: "文字选择",
  close: "关闭",
  loading: "正在打开文档…",
  empty: "尚未打开文档",
  emptyHint: "宿主应用会在这里打开本地文档。",
  unsupported: "无法显示此文档。",
  renderFailed: "无法呈现当前页面。",
  noResults: "无搜索结果",
  resultPosition: "{current} / {total}",
  pagePosition: "{current} / {total}",
  approximateLocation: "已显示最接近的可用位置。",
  exactLocation: "已定位来源。",
  noDiagnostics: "没有兼容性诊断",
  documentCanvas: "文档画布",
  documentNavigation: "文档导航",
  resizeNavigation: "调整文档导航宽度",
  outline: "目录",
  thumbnails: "缩略图",
  sheetTabs: "工作表",
  formulaBar: "公式栏",
  cellAddress: "单元格地址",
  columnHeaders: "列标题",
  rowHeaders: "行标题",
  resizeColumn: "调整 {label} 列宽",
  resizeRow: "调整第 {label} 行高",
  passwordRequired: "请输入密码以打开此 PDF。",
  passwordIncorrect: "密码错误，请重试。",
  passwordLabel: "PDF 密码",
  unlock: "打开",
  cancel: "取消",
});

const BUILTIN_MESSAGES: Readonly<Record<string, ViewerMessages>> = Object.freeze({
  en: EN,
  "zh-CN": ZH_CN,
  zh: ZH_CN,
});

const ICONS = Object.freeze({
  menu: '<path d="M4 7h16M4 12h16M4 17h16"/>',
  previous: '<path d="m15 18-6-6 6-6"/>',
  next: '<path d="m9 18 6-6-6-6"/>',
  search: '<circle cx="11" cy="11" r="6.5"/><path d="m16 16 4 4"/>',
  close: '<path d="m6 6 12 12M18 6 6 18"/>',
  up: '<path d="m7 14 5-5 5 5"/>',
  down: '<path d="m7 10 5 5 5-5"/>',
  minus: '<path d="M5 12h14"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  fit: '<rect x="4.5" y="5.5" width="15" height="13" rx="2"/><rect x="8" y="9" width="8" height="6" rx="1"/>',
  print: '<path d="M7 8V3h10v5M7 17H5a2 2 0 0 1-2-2v-4a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2v4a2 2 0 0 1-2 2h-2M7 14h10v7H7z"/>',
  themeAuto: '<rect x="3" y="4" width="18" height="13" rx="2"/><path d="M8 21h8m-4-4v4"/>',
  themeLight: '<circle cx="12" cy="12" r="4"/><path d="M12 2v2m0 16v2M2 12h2m16 0h2M5 5l1.5 1.5m11 11L19 19M5 19l1.5-1.5m11-11L19 5"/>',
  themeDark: '<path d="M20.5 14A9 9 0 0 1 10 3.5 9 9 0 1 0 20.5 14Z"/>',
  fullscreen: '<path d="M8 3H3v5M16 3h5v5M8 21H3v-5M16 21h5v-5"/>',
  continuous: '<rect x="5" y="3" width="14" height="7" rx="2"/><rect x="5" y="14" width="14" height="7" rx="2"/>',
  objectMode: '<path d="m5 4 13 8-6 2-2 6Z"/>',
  displayMode: '<path d="M2.5 12s3.5-5 9.5-5 9.5 5 9.5 5-3.5 5-9.5 5-9.5-5-9.5-5Z"/><circle cx="12" cy="12" r="2.5"/>',
  textMode: '<path d="M5 5h14M12 5v14M8 19h8"/>',
  outline: '<path d="M8 6h12M8 12h12M8 18h12M4 6h.01M4 12h.01M4 18h.01"/>',
  thumbnails: '<rect x="4" y="4" width="7" height="7" rx="1"/><rect x="13" y="4" width="7" height="7" rx="1"/><rect x="4" y="13" width="7" height="7" rx="1"/><rect x="13" y="13" width="7" height="7" rx="1"/>',
  warning: '<path d="M12 3 2.8 20h18.4L12 3Z"/><path d="M12 9v5M12 17.5v.1"/>',
});

const TEMPLATE = `
  <style>
    :host {
      --dv-accent: #1769d2;
      --dv-accent-contrast: #ffffff;
      --dv-background: #eef0f3;
      --dv-surface: #ffffff;
      --dv-surface-muted: #f7f8fa;
      --dv-text: #17191d;
      --dv-text-muted: #68707c;
      --dv-border: #d8dce2;
      --dv-hover: #edf3fb;
      --dv-focus: #1769d2;
      --dv-warning: #a15c00;
      --dv-selection: rgb(23 105 210 / 16%);
      --dv-object-hover: rgb(23 105 210 / 8%);
      --dv-object-hover-border: rgb(23 105 210 / 58%);
      --dv-page-shadow: 0 4px 18px rgb(21 28 38 / 18%);
      --dv-toolbar-height: 44px;
      --dv-status-height: 28px;
      --dv-sheet-tabs-height: 34px;
      --dv-formula-bar-height: 38px;
      --dv-navigation-width: 176px;
      --dv-icon-size: 20px;
      --dv-icon-stroke: 1.5;
      --dv-radius: 8px;
      --dv-motion-fast: 140ms;
      --dv-motion-state: 240ms;
      --dv-motion-layout: 180ms;
      --dv-ease-out: cubic-bezier(.22, 1, .36, 1);
      --dv-ease-toggle: cubic-bezier(.65, 0, .35, 1);
      display: block;
      min-width: 240px;
      min-height: 320px;
      color: var(--dv-text);
      background: var(--dv-background);
      color-scheme: light;
      contain: layout style;
      container: docviewkit / inline-size;
      font: 13px/1.4 Inter, ui-sans-serif, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }

    :host([data-resolved-theme="dark"]) {
      --dv-accent: #6aa7ff;
      --dv-accent-contrast: #0d1b2f;
      --dv-background: #17191d;
      --dv-surface: #22252b;
      --dv-surface-muted: #1d2025;
      --dv-text: #f4f5f7;
      --dv-text-muted: #a9afb9;
      --dv-border: #343842;
      --dv-hover: #2c3440;
      --dv-focus: #83b6ff;
      --dv-warning: #f2b55f;
      --dv-selection: rgb(106 167 255 / 18%);
      --dv-object-hover: rgb(106 167 255 / 10%);
      --dv-object-hover-border: rgb(131 182 255 / 72%);
      --dv-page-shadow: 0 5px 24px rgb(0 0 0 / 45%);
      color-scheme: dark;
    }

    *, *::before, *::after { box-sizing: border-box; }
    button, input { font: inherit; }
    button { color: inherit; }
    [hidden] { display: none !important; }

    .shell {
      position: relative;
      display: grid;
      grid-template: var(--dv-toolbar-height) minmax(0, 1fr) var(--dv-status-height)
        / var(--dv-navigation-width) 6px minmax(0, 1fr);
      width: 100%;
      height: 100%;
      min-height: inherit;
      overflow: hidden;
      background: var(--dv-background);
      border: 1px solid var(--dv-border);
      transition: color var(--dv-motion-state) var(--dv-ease-toggle),
        background-color var(--dv-motion-state) var(--dv-ease-toggle),
        border-color var(--dv-motion-state) var(--dv-ease-toggle),
        grid-template-columns var(--dv-motion-layout) var(--dv-ease-out);
    }

    .toolbar {
      position: relative;
      grid-column: 1 / -1;
      display: flex;
      align-items: center;
      gap: clamp(1px, .35cqi, 4px);
      min-width: 0;
      padding: 3px clamp(6px, .8cqi, 10px);
      background: var(--dv-surface);
      border-block-end: 1px solid var(--dv-border);
      z-index: 20;
      transition: background-color var(--dv-motion-state) var(--dv-ease-toggle),
        border-color var(--dv-motion-state) var(--dv-ease-toggle);
    }

    .toolbar-spacer { flex: 1; min-width: 4px; }
    .divider { width: 1px; height: 20px; margin-inline: clamp(2px, .4cqi, 5px); background: var(--dv-border); }
    .divider.secondary:not(:has(~ [data-action="print"]:not([hidden]), ~ [data-action="fullscreen"]:not([hidden]))) { display: none; }

    .interaction-switcher {
      display: inline-flex;
      flex: none;
      padding: 2px;
      border: 1px solid var(--dv-border);
      border-radius: 8px;
      background: var(--dv-surface-muted);
    }
    .interaction-option {
      display: inline-flex;
      width: 30px;
      height: 30px;
      align-items: center;
      justify-content: center;
      padding: 0;
      color: var(--dv-text-muted);
      background: transparent;
      border: 0;
      border-radius: 5px;
      cursor: pointer;
      transition: color var(--dv-motion-fast) var(--dv-ease-toggle),
        background-color var(--dv-motion-fast) var(--dv-ease-toggle),
        box-shadow var(--dv-motion-fast) var(--dv-ease-toggle);
    }
    .interaction-option:hover { color: var(--dv-text); background: var(--dv-hover); }
    .interaction-option[aria-pressed="true"] {
      color: var(--dv-accent);
      background: var(--dv-surface);
      box-shadow: 0 1px 2px rgb(20 27 38 / 14%);
    }
    .interaction-option:focus-visible { outline: 2px solid var(--dv-focus); outline-offset: 1px; z-index: 1; }
    .interaction-option .icon { display: block; width: 18px; height: 18px; }
    .interaction-option span { display: none; }

    @container docviewkit (min-width: 601px) {
      .interaction-switcher {
        position: absolute;
        left: 50%;
        transform: translateX(-50%);
      }
    }
    .toolbar:has(.search:not([hidden])) .interaction-switcher {
      position: static;
      transform: none;
    }

    .icon-button {
      display: inline-grid;
      place-items: center;
      flex: 0 0 32px;
      width: 32px;
      height: 32px;
      padding: 0;
      border: 0;
      border-radius: var(--dv-radius);
      background: transparent;
      cursor: pointer;
      transition: color var(--dv-motion-fast) var(--dv-ease-toggle),
        background-color var(--dv-motion-fast) var(--dv-ease-toggle),
        opacity var(--dv-motion-fast) var(--dv-ease-toggle),
        transform var(--dv-motion-fast) var(--dv-ease-out);
    }

    .icon-button:hover:not(:disabled), .icon-button[aria-pressed="true"] { background: var(--dv-hover); }
    .icon-button:active:not(:disabled) { transform: scale(.96); }
    .icon-button:focus-visible, .unit-button:focus-visible {
      outline: 2px solid var(--dv-focus);
      outline-offset: 1px;
    }
    .icon-button:disabled { opacity: .38; cursor: default; }
    .icon {
      width: var(--dv-icon-size);
      height: var(--dv-icon-size);
      fill: none;
      stroke: currentColor;
      stroke-width: var(--dv-icon-stroke);
      stroke-linecap: round;
      stroke-linejoin: round;
    }

    .search {
      display: flex;
      align-items: center;
      width: min(300px, 32cqi);
      max-width: 300px;
      height: 36px;
      border: 1px solid var(--dv-border);
      border-radius: var(--dv-radius);
      background: var(--dv-surface-muted);
      overflow: hidden;
      transition: background-color var(--dv-motion-state) var(--dv-ease-toggle),
        border-color var(--dv-motion-fast) var(--dv-ease-toggle),
        box-shadow var(--dv-motion-fast) var(--dv-ease-toggle);
    }
    .search:not([hidden]) { animation: control-in var(--dv-motion-state) var(--dv-ease-out); }
    .search:focus-within { border-color: var(--dv-focus); box-shadow: 0 0 0 1px var(--dv-focus); }
    .search .icon-button { flex-basis: 32px; width: 32px; height: 32px; border-radius: 6px; }
    .search > .icon { flex: 0 0 22px; margin-inline-start: 8px; }
    .search-input { min-width: 0; flex: 1; height: 100%; padding-inline: 8px 4px; border: 0; outline: 0; color: inherit; background: transparent; }
    .search-position { min-width: 48px; color: var(--dv-text-muted); text-align: center; font-variant-numeric: tabular-nums; }

    .page-position, .zoom-value {
      min-width: 66px;
      color: var(--dv-text);
      text-align: center;
      font-size: 14px;
      font-variant-numeric: tabular-nums;
      white-space: nowrap;
    }
    .zoom-value {
      display: inline-flex;
      min-width: 52px;
      align-items: center;
      justify-content: center;
      gap: 1px;
    }
    .zoom-value:focus-within { border-radius: 4px; outline: 2px solid var(--dv-focus); outline-offset: 1px; }
    .zoom-input {
      width: 32px;
      padding: 0;
      color: inherit;
      background: transparent;
      border: 0;
      outline: 0;
      appearance: textfield;
      font: inherit;
      text-align: end;
    }
    .zoom-input::-webkit-inner-spin-button,
    .zoom-input::-webkit-outer-spin-button { margin: 0; appearance: none; }
    .zoom-input:disabled { opacity: .38; }
    .object-clipboard {
      position: fixed;
      inset: 0 auto auto -10000px;
      width: 1px;
      height: 1px;
      opacity: 0;
      pointer-events: none;
    }

    .navigation {
      grid-row: 2;
      grid-column: 1;
      width: 100%;
      min-width: 0;
      overflow: auto;
      padding: 12px 10px;
      background: var(--dv-surface-muted);
      border-inline-end: 1px solid var(--dv-border);
      scrollbar-width: thin;
      z-index: 12;
      transition: background-color var(--dv-motion-state) var(--dv-ease-toggle),
        border-color var(--dv-motion-state) var(--dv-ease-toggle),
        opacity var(--dv-motion-layout) var(--dv-ease-out),
        transform var(--dv-motion-layout) var(--dv-ease-out),
        visibility 0s;
    }
    .shell[data-navigation="false"] { grid-template-columns: 0 0 minmax(0, 1fr); }
    .shell[data-navigation="false"] .navigation,
    .shell[data-navigation="false"] .navigation-resizer {
      opacity: 0;
      visibility: hidden;
      pointer-events: none;
    }
    .shell[data-navigation="false"] .navigation {
      transform: translateX(-12px);
      transition-delay: 0s, 0s, 0s, 0s, var(--dv-motion-layout);
    }
    .shell[dir="rtl"][data-navigation="false"] .navigation { transform: translateX(12px); }
    .shell[data-navigation="false"] .navigation-resizer {
      transition-delay: 0s, var(--dv-motion-layout);
    }
    .shell[data-spreadsheet="true"] .navigation,
    .shell[data-spreadsheet="true"] .navigation-resizer,
    .shell[data-spreadsheet="true"] [data-action="navigation"] { display: none; }
    .shell[data-spreadsheet="true"] .workspace { grid-row: 3; grid-column: 1 / -1; }
    .navigation-resizer {
      position: relative;
      z-index: 13;
      grid-row: 2;
      grid-column: 2;
      cursor: col-resize;
      touch-action: none;
      transition: opacity var(--dv-motion-layout) var(--dv-ease-out), visibility 0s;
    }
    .navigation-resizer::before {
      content: "";
      position: absolute;
      inset-block: 0;
      inset-inline-start: 2px;
      width: 2px;
      background: transparent;
      transition: background-color var(--dv-motion-fast) var(--dv-ease-toggle);
    }
    .navigation-resizer:hover::before,
    .navigation-resizer:focus-visible::before,
    .navigation-resizer[data-resizing="true"]::before { background: var(--dv-accent); }
    .navigation-switcher {
      position: sticky;
      z-index: 2;
      inset-block-start: -12px;
      display: flex;
      align-items: center;
      justify-content: flex-start;
      gap: 0;
      margin: -12px -10px 10px;
      padding: 7px 10px;
      background: var(--dv-surface-muted);
      border-block-end: 1px solid var(--dv-border);
    }
    .navigation-mode {
      position: relative;
      display: inline-flex;
      width: 32px;
      height: 28px;
      flex: 0 0 32px;
      align-items: center;
      justify-content: center;
      padding: 0;
      color: var(--dv-text-muted);
      background: var(--dv-surface-muted);
      border: 1px solid var(--dv-border);
      border-inline-start-width: 0;
      border-radius: 0;
      cursor: pointer;
      transition:
        color var(--dv-motion-fast) var(--dv-ease-toggle),
        background-color var(--dv-motion-fast) var(--dv-ease-toggle),
        box-shadow var(--dv-motion-fast) var(--dv-ease-toggle);
    }
    .navigation-mode:first-child {
      border-inline-start-width: 1px;
      border-start-start-radius: 6px;
      border-end-start-radius: 6px;
    }
    .navigation-mode:last-child {
      border-start-end-radius: 6px;
      border-end-end-radius: 6px;
    }
    .navigation-mode:hover {
      color: var(--dv-text);
      background: var(--dv-hover);
    }
    .navigation-mode[aria-selected="true"] {
      z-index: 1;
      color: var(--dv-text);
      background: var(--dv-surface);
      box-shadow: 0 1px 2px rgb(20 27 38 / 12%);
    }
    .navigation-mode:focus-visible {
      z-index: 2;
      outline: 2px solid var(--dv-focus);
      outline-offset: 1px;
    }
    .navigation-mode .icon { width: 15px; height: 15px; }
    @media (pointer: coarse) {
      .navigation-mode {
        width: 40px;
        height: 40px;
        flex-basis: 40px;
      }
    }
    .unit-list { display: grid; gap: 10px; }
    .unit-button {
      display: grid;
      grid-template-columns: minmax(0, 1fr);
      grid-template-rows: auto auto;
      align-items: center;
      gap: 4px;
      width: 100%;
      padding: 4px;
      color: var(--dv-text-muted);
      background: transparent;
      border: 1px solid transparent;
      border-radius: 6px;
      cursor: pointer;
      transition: color var(--dv-motion-fast) var(--dv-ease-toggle),
        background-color var(--dv-motion-fast) var(--dv-ease-toggle),
        border-color var(--dv-motion-fast) var(--dv-ease-toggle),
        transform var(--dv-motion-fast) var(--dv-ease-out);
    }
    .unit-button[aria-current="page"] { color: var(--dv-accent); border-color: var(--dv-accent); background: var(--dv-surface); }
    .unit-button:active { transform: scale(.985); }
    .unit-number {
      grid-row: 2;
      text-align: center;
      font-variant-numeric: tabular-nums;
    }
    .thumbnail {
      --dv-thumbnail-aspect: 1;
      --dv-thumbnail-max-width: 96px;
      display: grid;
      place-items: center;
      justify-self: center;
      width: min(100%, var(--dv-thumbnail-max-width));
      aspect-ratio: var(--dv-thumbnail-aspect);
      overflow: hidden;
      background: #ffffff;
      box-shadow: inset 0 0 0 1px #cfd4dc, 0 1px 3px rgb(20 27 38 / 12%);
      transition: transform var(--dv-motion-state) var(--dv-ease-out),
        box-shadow var(--dv-motion-state) var(--dv-ease-toggle);
    }
    .unit-button[aria-current="page"] .thumbnail {
      transform: scale(1.015);
      box-shadow: inset 0 0 0 1px #cfd4dc, 0 3px 10px rgb(20 27 38 / 16%);
    }
    .thumbnail canvas { display: block; width: 100%; height: 100%; min-width: 0; min-height: 0; }
    .outline-button {
      width: 100%;
      min-height: 32px;
      padding: 6px 8px;
      padding-inline-start: calc(8px + min(var(--dv-outline-level), 8) * 14px);
      overflow: hidden;
      color: var(--dv-text-muted);
      background: transparent;
      border: 0;
      border-radius: 6px;
      text-align: start;
      text-overflow: ellipsis;
      white-space: nowrap;
      cursor: pointer;
    }
    .outline-button:hover { background: var(--dv-hover); }
    .outline-button[aria-current="page"] { color: var(--dv-accent); background: var(--dv-surface); }

    .workspace {
      position: relative;
      grid-row: 2;
      grid-column: 3;
      min-width: 0;
      min-height: 0;
      overflow: auto;
      overscroll-behavior: contain;
      scrollbar-width: thin;
      transition: background-color var(--dv-motion-state) var(--dv-ease-toggle);
    }
    .stage-wrap {
      display: flex;
      align-items: flex-start;
      justify-content: center;
      width: max-content;
      min-width: 100%;
      min-height: 100%;
      padding: clamp(18px, 2.2cqi, 28px);
    }
    .shell[data-spreadsheet="true"] {
      grid-template-rows: var(--dv-toolbar-height) var(--dv-formula-bar-height) minmax(0, 1fr) var(--dv-sheet-tabs-height) var(--dv-status-height);
    }
    .formula-bar {
      display: none;
      grid-column: 1 / -1;
      grid-row: 2;
      align-items: center;
      min-width: 0;
      color: var(--dv-text);
      background: var(--dv-surface);
      border-block-end: 1px solid var(--dv-border);
      z-index: 20;
    }
    .shell[data-spreadsheet="true"] .formula-bar { display: flex; }
    .cell-address {
      flex: 0 0 88px;
      padding-inline: 12px;
      font-variant-numeric: tabular-nums;
    }
    .formula-symbol {
      flex: 0 0 44px;
      color: var(--dv-text-muted);
      border-inline: 1px solid var(--dv-border);
      font: italic 20px/1 Georgia, serif;
      text-align: center;
    }
    .formula-value {
      min-width: 0;
      padding-inline: 14px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .shell[data-spreadsheet="true"] .stage-wrap {
      position: relative;
      display: grid;
      place-items: initial;
      justify-content: start;
      align-content: start;
      min-width: 100%;
      min-height: 100%;
      padding: 0;
      background: #ffffff;
    }
    .shell[data-spreadsheet="true"] .stage {
      position: absolute;
      margin: 0;
      box-shadow: none;
    }
    .sheet-tile { position: absolute; pointer-events: none; }
    .continuous-view {
      display: flex;
      width: max-content;
      min-width: 100%;
      min-height: 100%;
      flex-direction: column;
      align-items: center;
      gap: clamp(16px, 2cqi, 24px);
      padding: clamp(18px, 2.2cqi, 28px);
    }
    .continuous-page {
      position: relative;
      flex: none;
      overflow: hidden;
      transform-origin: top left;
      content-visibility: auto;
      contain: layout paint style;
      contain-intrinsic-size: auto 1100px;
      background: #ffffff;
      box-shadow: var(--dv-page-shadow);
      transition: box-shadow var(--dv-motion-state) var(--dv-ease-toggle);
    }
    .continuous-page canvas { display: block; width: 100%; height: 100%; background: #ffffff; }
    .continuous-selection, .object-hover {
      position: absolute;
      z-index: 3;
      border-radius: 3px;
      pointer-events: none;
    }
    .continuous-selection {
      border: 2px solid var(--dv-accent);
      background: var(--dv-selection);
      animation: locate 220ms var(--dv-ease-out);
    }
    .object-hover {
      border: 1px solid var(--dv-object-hover-border);
      background: var(--dv-object-hover);
      animation: object-hover-in var(--dv-motion-fast) var(--dv-ease-out);
    }
    .shell[data-object-hover="true"] .stage,
    .shell[data-object-hover="true"] .continuous-page { cursor: pointer; }
    .stage {
      position: relative;
      flex: none;
      background: #ffffff;
      box-shadow: var(--dv-page-shadow);
      transform-origin: top left;
      margin: auto;
      transition: box-shadow var(--dv-motion-state) var(--dv-ease-toggle);
    }
    .surface { position: relative; z-index: 1; display: block; background: #ffffff; }
    .text-layer {
      position: absolute;
      z-index: 2;
      inset: 0;
      overflow: hidden;
      pointer-events: none;
      user-select: text;
      -webkit-user-select: text;
    }
    .text-layer-item {
      position: absolute;
      display: flex;
      flex-direction: column;
      overflow: hidden;
      margin: 0;
      padding: 0;
      color: transparent;
      background: transparent;
      font: 12px/1.2 Aptos, "Segoe UI Variable", "Microsoft YaHei UI", sans-serif;
      overflow-wrap: anywhere;
      white-space: pre-wrap;
      cursor: text;
      pointer-events: auto;
      user-select: text;
      -webkit-user-select: text;
    }
    .text-layer-item[data-layout-source="render"] { display: contents; }
    .text-layer-fragment {
      position: absolute;
      display: block;
      overflow: visible;
      margin: 0;
      padding: 0;
      color: transparent;
      background: transparent;
      white-space: pre;
      cursor: text;
      pointer-events: auto;
      transform-origin: 0 0;
      user-select: text;
      -webkit-user-select: text;
    }
    .text-layer-item::selection { color: transparent; background: var(--dv-selection); }
    .text-layer-fragment::selection { color: transparent; background: transparent; }
    .search-highlight-fragment { pointer-events: none; user-select: none; -webkit-user-select: none; }
    .search-highlight-match {
      color: transparent;
      background: rgb(255 196 0 / 48%);
      border-radius: 2px;
    }
    .text-selection-highlight {
      position: absolute;
      background: var(--dv-selection);
      pointer-events: none;
      user-select: none;
      -webkit-user-select: none;
    }
    .page-transition-outgoing {
      position: absolute;
      inset: 0;
      z-index: 2;
      display: block;
      width: 100%;
      height: 100%;
      background: #ffffff;
      pointer-events: none;
    }
    .selection {
      position: absolute;
      z-index: 3;
      border: 2px solid var(--dv-accent);
      background: var(--dv-selection);
      border-radius: 3px;
      pointer-events: none;
      animation: locate 220ms var(--dv-ease-out);
    }
    .sheet-header-axis, .sheet-header-corner, .sheet-resize-guide { display: none; }
    .shell[data-spreadsheet="true"] .sheet-header-axis,
    .shell[data-spreadsheet="true"] .sheet-header-corner { display: block; }
    .sheet-header-axis, .sheet-header-corner {
      position: sticky;
      z-index: 6;
      grid-area: 1 / 1;
      align-self: start;
      justify-self: start;
      overflow: hidden;
      color: #3f4650;
      background: #f3f5f6;
      border-color: #c8cdd3;
      user-select: none;
      -webkit-user-select: none;
    }
    .shell[data-sheet-resizing="true"] .workspace,
    .shell[data-sheet-resizing="true"] .workspace * {
      user-select: none;
      -webkit-user-select: none;
    }
    .sheet-column-headers {
      top: 0;
      height: 24px;
      margin-inline-start: 48px;
      border-block: 1px solid #c8cdd3;
    }
    .sheet-row-headers {
      left: 0;
      width: 48px;
      margin-block-start: 24px;
      border-inline: 1px solid #c8cdd3;
    }
    .sheet-header-corner {
      top: 0;
      left: 0;
      width: 48px;
      height: 24px;
      border: 1px solid #c8cdd3;
      z-index: 7;
    }
    .sheet-header-item {
      position: absolute;
      display: grid;
      place-items: center;
      min-width: 0;
      min-height: 0;
      border-inline-end: 1px solid #c8cdd3;
      border-block-end: 1px solid #c8cdd3;
      font-size: 11px;
      line-height: 1;
    }
    .sheet-header-resizer {
      position: absolute;
      z-index: 1;
      touch-action: none;
    }
    .sheet-column-headers .sheet-header-resizer {
      inset-block: 0;
      inset-inline-end: -4px;
      width: 8px;
      cursor: col-resize;
    }
    .sheet-row-headers .sheet-header-resizer {
      inset-inline: 0;
      inset-block-end: -4px;
      height: 8px;
      cursor: row-resize;
    }
    .sheet-resize-guide {
      position: absolute;
      z-index: 8;
      background: var(--dv-accent);
      pointer-events: none;
    }
    .sheet-resize-guide[data-visible="true"] { display: block; }
    .sheet-page-breaks {
      position: absolute;
      z-index: 5;
      pointer-events: none;
    }
    .sheet-page-break {
      position: absolute;
      border-color: #5b9bd5;
      border-style: dashed;
    }
    .sheet-page-break[data-axis="column"] { inset-block: 0; border-inline-start-width: 2px; }
    .sheet-page-break[data-axis="row"] { inset-inline: 0; border-block-start-width: 2px; }
    .sheet-resize-guide[data-axis="column"] { width: 1px; }
    .sheet-resize-guide[data-axis="row"] { height: 1px; }

    .sheet-tabs {
      display: none;
      grid-column: 1 / -1;
      grid-row: 4;
      align-items: end;
      min-width: 0;
      overflow-x: auto;
      overflow-y: hidden;
      padding-inline: 10px;
      background: var(--dv-surface-muted);
      border-block-start: 1px solid var(--dv-border);
      scrollbar-width: thin;
      z-index: 20;
    }
    .shell[data-spreadsheet="true"] .sheet-tabs { display: flex; }
    .sheet-tab {
      position: relative;
      flex: 0 0 auto;
      height: 33px;
      padding: 0 16px;
      color: var(--dv-text-muted);
      background: transparent;
      border: 0;
      border-inline-end: 1px solid var(--dv-border);
      cursor: pointer;
      white-space: nowrap;
    }
    .sheet-tab[aria-selected="true"] { color: var(--dv-text); background: var(--dv-surface); font-weight: 600; }
    .sheet-tab[data-tab-color]::after,
    .sheet-tab[aria-selected="true"]::after {
      content: "";
      position: absolute;
      inset-inline: 10px;
      inset-block-end: 0;
      height: 2px;
      background: var(--dv-sheet-tab-color, var(--dv-accent));
    }
    .sheet-tab:focus-visible, .sheet-header-resizer:focus-visible, .navigation-resizer:focus-visible {
      outline: 2px solid var(--dv-focus);
      outline-offset: -2px;
    }
    @keyframes locate { from { opacity: 0; transform: scale(.98); } }
    @keyframes object-hover-in { from { opacity: 0; transform: scale(.995); } }
    @keyframes control-in { from { opacity: 0; transform: translateY(-4px) scale(.99); } }
    @keyframes state-in { from { opacity: 0; transform: translateY(5px); } }

    .empty, .loading, .error {
      position: absolute;
      inset: 0;
      display: grid;
      place-content: center;
      justify-items: center;
      gap: 8px;
      padding: 28px;
      color: var(--dv-text-muted);
      text-align: center;
      animation: state-in var(--dv-motion-state) var(--dv-ease-out);
    }
    .state-title { color: var(--dv-text); font-size: 15px; font-weight: 600; }
    .state-hint { max-width: 36ch; }
    .spinner { width: 28px; height: 28px; border: 2px solid var(--dv-border); border-block-start-color: var(--dv-accent); border-radius: 50%; animation: spin .8s linear infinite; }
    @keyframes spin { to { transform: rotate(1turn); } }

    .statusbar {
      grid-column: 1 / -1;
      grid-row: 3;
      display: flex;
      align-items: center;
      gap: 8px;
      min-width: 0;
      padding: 0 12px;
      color: var(--dv-text-muted);
      background: var(--dv-surface);
      border-block-start: 1px solid var(--dv-border);
      z-index: 20;
      transition: color var(--dv-motion-state) var(--dv-ease-toggle),
        background-color var(--dv-motion-state) var(--dv-ease-toggle),
        border-color var(--dv-motion-state) var(--dv-ease-toggle);
    }
    .shell[data-spreadsheet="true"] .statusbar { grid-row: 5; }
    .status-message { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .statusbar .icon-button, .statusbar .page-position { display: none; }
    .statusbar .page-position { margin-inline: auto; }

    .diagnostic-panel {
      position: absolute;
      inset-block: var(--dv-toolbar-height) var(--dv-status-height);
      inset-inline-end: 0;
      z-index: 30;
      width: min(360px, 88cqi);
      overflow: auto;
      padding: 16px;
      color: var(--dv-text);
      background: var(--dv-surface);
      border-inline-start: 1px solid var(--dv-border);
      box-shadow: -8px 0 24px rgb(16 24 40 / 14%);
      animation: diagnostic-in var(--dv-motion-layout) var(--dv-ease-out);
    }
    .shell[data-spreadsheet="true"] .diagnostic-panel {
      inset-block-end: calc(var(--dv-status-height) + var(--dv-sheet-tabs-height));
    }
    @keyframes diagnostic-in { from { opacity: 0; transform: translateX(16px); } }
    .diagnostic-heading { display: flex; align-items: center; justify-content: space-between; margin-block-end: 12px; }
    .diagnostic-heading strong { font-size: 14px; }
    .diagnostic-list { display: grid; gap: 10px; margin: 0; padding: 0; list-style: none; }
    .diagnostic-item { padding-block-end: 10px; border-block-end: 1px solid var(--dv-border); }
    .diagnostic-code { display: block; margin-block-end: 3px; color: var(--dv-warning); font: 600 11px/1.3 ui-monospace, monospace; }

    .password-dialog {
      width: min(360px, calc(100% - 32px));
      padding: 0;
      color: var(--dv-text);
      background: var(--dv-surface);
      border: 1px solid var(--dv-border);
      border-radius: 12px;
      box-shadow: 0 18px 48px rgb(16 24 40 / 24%);
    }
    .password-dialog::backdrop { background: rgb(16 24 40 / 42%); backdrop-filter: blur(2px); }
    .password-form { display: grid; gap: 16px; padding: 20px; }
    .password-heading { font-size: 15px; }
    .password-field { display: grid; gap: 6px; color: var(--dv-text-muted); }
    .password-input {
      width: 100%;
      min-height: 40px;
      padding: 8px 10px;
      color: var(--dv-text);
      background: var(--dv-surface);
      border: 1px solid var(--dv-border);
      border-radius: 7px;
      outline: none;
    }
    .password-input:focus { border-color: var(--dv-focus); box-shadow: 0 0 0 2px var(--dv-selection); }
    .password-actions { display: flex; justify-content: flex-end; gap: 8px; }
    .password-action {
      min-height: 36px;
      padding: 7px 14px;
      border: 1px solid var(--dv-border);
      border-radius: 7px;
      color: var(--dv-text);
      background: var(--dv-surface);
      cursor: pointer;
    }
    .password-action-primary {
      color: var(--dv-accent-contrast);
      background: var(--dv-accent);
      border-color: var(--dv-accent);
    }

    @container docviewkit (max-width: 720px) {
      .shell {
        grid-template-columns: 0 0 minmax(0, 1fr);
      }
      .toolbar { gap: 1px; padding: 1px 6px; }
      .toolbar .icon-button { flex-basis: 40px; width: 40px; height: 40px; }
      .toolbar .desktop-secondary, .toolbar .divider.secondary,
      .toolbar [data-action="fullscreen"] { display: none; }
      .search { position: absolute; inset: 2px 46px; z-index: 2; width: auto; height: 40px; background: var(--dv-surface); }
      .search[hidden] { display: none !important; }
      .toolbar:has(.search:not([hidden])) .interaction-switcher { display: none; }
      .navigation {
        position: absolute;
        inset-block: var(--dv-toolbar-height) var(--dv-status-height);
        inset-inline-start: 0;
        width: min(220px, 80%);
        box-shadow: 8px 0 24px rgb(0 0 0 / 22%);
      }
      .navigation-resizer { display: none; }
      .workspace { grid-column: 1 / -1; }
      .stage-wrap { padding: 18px 12px; }
      .shell[data-spreadsheet="true"] .stage-wrap { padding: 0; }
      .continuous-view { gap: 16px; padding: 18px 12px; }
      .toolbar .page-position { display: none; }
      .interaction-option { width: 40px; height: 40px; padding: 0; }
      .statusbar .page-position { display: block; }
      .status-message { display: none; }
      .statusbar { justify-content: space-between; padding-inline: 8px; }
    }

    @container docviewkit (max-width: 540px) {
      .toolbar .interaction-switcher { display: none; }
    }

    @container docviewkit (max-width: 480px) {
      .toolbar [data-action="fit"],
      .toolbar [data-action="fullscreen"],
      .toolbar .desktop-page-mode { display: none; }
      .statusbar .page-position { display: none; }
        }

    @container docviewkit (max-width: 360px) {
      .toolbar [data-action="previous"], .toolbar [data-action="next"] { display: none; }
    }

    @media (prefers-reduced-motion: reduce) {
      *, *::before, *::after {
        animation-duration: .01ms !important;
        animation-iteration-count: 1 !important;
        transition-duration: .01ms !important;
        scroll-behavior: auto !important;
      }
    }

    @media (hover: hover) and (pointer: fine) {
      .icon-button:hover:not(:disabled) { transform: translateY(-1px); }
      .unit-button:hover:not([aria-current="page"]) { transform: translateY(-1px); }
    }

    .shell[data-minimal="true"] { display: block; border: 0; }
    .shell[data-minimal="true"] :is(.toolbar, .formula-bar, .navigation, .navigation-resizer, .sheet-tabs, .statusbar, .diagnostic-panel) { display: none; }
    .shell[data-minimal="true"] .workspace { width: 100%; height: 100%; }

    @media print {
      :host { height: auto !important; overflow: visible !important; }
      .toolbar, .formula-bar, .navigation, .navigation-resizer, .sheet-tabs, .statusbar, .diagnostic-panel, .selection,
      .sheet-header-axis, .sheet-header-corner, .sheet-resize-guide, .sheet-tile { display: none !important; }
      .shell, .workspace { display: block; border: 0; overflow: visible; background: #ffffff; }
      .stage-wrap { display: block; min-width: 0; padding: 0; }
      .stage { position: static !important; box-shadow: none; }
      .continuous-view { gap: 0; padding: 0; }
      .continuous-page { box-shadow: none; break-after: page; }
    }
  </style>
  <section class="shell" part="shell" data-navigation="false">
    <header class="toolbar" part="toolbar">
      <button class="icon-button" data-action="navigation"><svg class="icon" viewBox="0 0 24 24">${ICONS.menu}</svg></button>
      <slot name="toolbar-start"></slot>
      <span class="divider"></span>
      <button class="icon-button" data-action="previous"><svg class="icon" viewBox="0 0 24 24">${ICONS.previous}</svg></button>
      <button class="icon-button" data-action="next"><svg class="icon" viewBox="0 0 24 24">${ICONS.next}</svg></button>
      <span class="divider"></span>
      <button class="icon-button" data-action="search"><svg class="icon" viewBox="0 0 24 24">${ICONS.search}</svg></button>
      <form class="search" role="search" hidden>
        <svg class="icon" viewBox="0 0 24 24" aria-hidden="true">${ICONS.search}</svg>
        <input class="search-input" type="search" autocomplete="off" maxlength="1024">
        <output class="search-position" aria-live="polite">0 / 0</output>
        <button class="icon-button" type="button" data-action="previous-result"><svg class="icon" viewBox="0 0 24 24">${ICONS.up}</svg></button>
        <button class="icon-button" type="button" data-action="next-result"><svg class="icon" viewBox="0 0 24 24">${ICONS.down}</svg></button>
        <button class="icon-button" type="button" data-action="close-search"><svg class="icon" viewBox="0 0 24 24">${ICONS.close}</svg></button>
      </form>
      <span class="divider desktop-secondary"></span>
      <output class="page-position">0 / 0</output>
      <span class="toolbar-spacer"></span>
      <div class="interaction-switcher" role="group" hidden>
        <button class="interaction-option" type="button" data-action="interaction-mode" data-interaction-mode="object" aria-pressed="true"><svg class="icon" viewBox="0 0 24 24">${ICONS.objectMode}</svg><span></span></button>
        <button class="interaction-option" type="button" data-action="interaction-mode" data-interaction-mode="display" aria-pressed="false"><svg class="icon" viewBox="0 0 24 24">${ICONS.displayMode}</svg><span></span></button>
        <button class="interaction-option" type="button" data-action="interaction-mode" data-interaction-mode="text" aria-pressed="false"><svg class="icon" viewBox="0 0 24 24">${ICONS.textMode}</svg><span></span></button>
      </div>
      <button class="icon-button" data-action="zoom-out"><svg class="icon" viewBox="0 0 24 24">${ICONS.minus}</svg></button>
      <label class="zoom-value"><input class="zoom-input" type="number" min="${MIN_ZOOM * 100}"
        max="${MAX_ZOOM * 100}" step="1" value="100"><span aria-hidden="true">%</span></label>
      <button class="icon-button" data-action="zoom-in"><svg class="icon" viewBox="0 0 24 24">${ICONS.plus}</svg></button>
      <button class="icon-button" data-action="fit"><svg class="icon" viewBox="0 0 24 24">${ICONS.fit}</svg></button>
      <button class="icon-button desktop-page-mode" data-action="page-mode"><svg class="icon" viewBox="0 0 24 24">${ICONS.continuous}</svg></button>
      <span class="divider secondary"></span>
      <button class="icon-button desktop-secondary" data-action="print"><svg class="icon" viewBox="0 0 24 24">${ICONS.print}</svg></button>
      <button class="icon-button" type="button" data-action="theme"><svg class="icon" aria-hidden="true" viewBox="0 0 24 24"></svg></button>
      <button class="icon-button" data-action="fullscreen"><svg class="icon" viewBox="0 0 24 24">${ICONS.fullscreen}</svg></button>
      <slot name="toolbar-end"></slot>
      <button class="icon-button" data-action="diagnostics"><svg class="icon" viewBox="0 0 24 24">${ICONS.warning}</svg></button>
    </header>
    <nav class="navigation" id="navigation-panel" part="navigation">
      <div class="navigation-switcher" role="tablist">
        <button class="navigation-mode" type="button" role="tab" aria-controls="navigation-view" data-action="navigation-outline">
          <svg class="icon" viewBox="0 0 24 24">${ICONS.outline}</svg>
        </button>
        <button class="navigation-mode" type="button" role="tab" aria-controls="navigation-view" data-action="navigation-thumbnails">
          <svg class="icon" viewBox="0 0 24 24">${ICONS.thumbnails}</svg>
        </button>
      </div>
      <div id="navigation-view" role="tabpanel"><div class="unit-list" role="listbox"></div></div>
    </nav>
    <div class="navigation-resizer" part="navigation-resizer" role="separator" aria-orientation="vertical"
      aria-controls="navigation-panel" aria-valuemin="120" aria-valuemax="400" tabindex="0"></div>
    <div class="formula-bar" part="formula-bar" role="group">
      <output class="cell-address">A1</output><span class="formula-symbol" aria-hidden="true">fx</span><output class="formula-value"></output>
  </div>
  <textarea class="object-clipboard" tabindex="-1" aria-hidden="true"></textarea>
  <main class="workspace" part="workspace" tabindex="0">
      <div class="stage-wrap">
        <div class="stage" part="document"><canvas class="surface"></canvas><div class="text-layer" hidden></div><div class="object-hover" hidden></div><div class="selection" hidden></div></div>
        <div class="sheet-page-breaks" aria-hidden="true" hidden></div>
        <div class="sheet-column-headers sheet-header-axis" role="row" part="column-headers"></div>
        <div class="sheet-row-headers sheet-header-axis" role="rowgroup" part="row-headers"></div>
        <div class="sheet-header-corner" aria-hidden="true"></div>
        <div class="sheet-resize-guide" data-visible="false" aria-hidden="true"></div>
      </div>
      <div class="continuous-view" part="continuous" hidden></div>
      <div class="empty"><strong class="state-title"></strong><span class="state-hint"></span></div>
      <div class="loading" hidden><span class="spinner" aria-hidden="true"></span><strong class="state-title"></strong></div>
      <div class="error" hidden><svg class="icon" viewBox="0 0 24 24" aria-hidden="true">${ICONS.warning}</svg><strong class="state-title"></strong><span class="state-hint"></span></div>
    </main>
    <nav class="sheet-tabs" part="sheet-tabs" role="tablist"></nav>
    <footer class="statusbar" part="statusbar">
      <span class="status-message" aria-live="polite"></span>
      <button class="icon-button" data-action="mobile-previous"><svg class="icon" viewBox="0 0 24 24">${ICONS.previous}</svg></button>
      <output class="page-position">0 / 0</output>
      <button class="icon-button" data-action="mobile-next"><svg class="icon" viewBox="0 0 24 24">${ICONS.next}</svg></button>
      <button class="icon-button mobile-page-mode" data-action="page-mode"><svg class="icon" viewBox="0 0 24 24">${ICONS.continuous}</svg></button>
    </footer>
    <aside class="diagnostic-panel" part="diagnostics" hidden>
      <div class="diagnostic-heading"><strong></strong><button class="icon-button" data-action="close-diagnostics"><svg class="icon" viewBox="0 0 24 24">${ICONS.close}</svg></button></div>
      <ul class="diagnostic-list"></ul>
    </aside>
  </section>
  <dialog class="password-dialog" part="password-dialog">
    <form class="password-form" method="dialog">
      <strong class="password-heading"></strong>
      <label class="password-field">
        <span class="password-label"></span>
        <input class="password-input" type="password" autocomplete="current-password" maxlength="1024">
      </label>
      <div class="password-actions">
        <button class="password-action" type="button" data-action="password-cancel"></button>
        <button class="password-action password-action-primary" value="unlock"></button>
      </div>
    </form>
  </dialog>
`;

function normalizedLocale(value: string): string {
  const locale = value.trim().replace(/_/gu, "-");
  if (locale.length === 0 || locale === "auto") return globalThis.navigator?.language ?? "en";
  try {
    return new Intl.Locale(locale).toString();
  } catch {
    return "en";
  }
}

function messageCatalog(locale: string, custom: ViewerConfig["messages"]): ViewerMessages {
  const language = locale.split("-")[0] ?? "en";
  const builtin = BUILTIN_MESSAGES[locale] ?? BUILTIN_MESSAGES[language] ?? EN;
  const customValues = custom?.[locale] ?? custom?.[language];
  return Object.freeze({ ...builtin, ...customValues });
}

function formatMessage(template: string, values: Readonly<Record<string, string | number>>): string {
  return template.replace(/\{(\w+)\}/gu, (_, key: string) => String(values[key] ?? ""));
}

function isRtl(locale: string): boolean {
  return /^(?:ar|fa|he|ur)(?:-|$)/iu.test(locale);
}

function errorMessage(cause: unknown): string {
  return cause instanceof Error ? cause.message : String(cause);
}

function copyRect(rect: Rect): Rect {
  return Object.freeze({ x: rect.x, y: rect.y, width: rect.width, height: rect.height });
}

function safeHyperlink(value: string): string | undefined {
  const candidate = value.trim();
  if (candidate.length === 0 || candidate.length > 2_048 || /\s/u.test(candidate)) return undefined;
  try {
    const parsed = new URL(candidate);
    return ["http:", "https:", "mailto:", "tel:"].includes(parsed.protocol) ? parsed.href : undefined;
  } catch {
    return undefined;
  }
}

const CONTINUOUS_CACHE_PIXELS = 40_000_000;
const CONTINUOUS_SMOOTH_SCROLL_VIEWPORTS = 2;
const PINCH_RENDER_DELAY = 120;
const NAVIGATION_WINDOW_SIZE = 40;
const NAVIGATION_ITEM_HEIGHT = 132;
const MIN_NAVIGATION_WIDTH = 120;
const MAX_NAVIGATION_WIDTH = 400;
const SHEET_ROW_HEADER_WIDTH = 48;
const SHEET_COLUMN_HEADER_HEIGHT = 24;
const SHEET_VIEWPORT_OVERSCAN_RATIO = 0.5;
const MIN_COLUMN_WIDTH = 12;
const MIN_ROW_HEIGHT = 8;
const MAX_TEXT_LAYER_OBJECTS = 10_000;

class SheetAxisLayout {
  readonly #axis: SheetAxis;
  readonly #count: number;
  readonly #overrides: ReadonlyMap<number, number>;

  constructor(axis: SheetAxis, count: number, overrides: ReadonlyMap<number, number>) {
    this.#axis = axis;
    this.#count = count;
    this.#overrides = overrides;
  }

  #baseSize(index: number): number {
    return this.#axis.spans.find((span) => span.start <= index && index <= span.end)?.size
      ?? this.#axis.defaultSize;
  }

  #baseOffset(index: number): number {
    let offset = index * this.#axis.defaultSize;
    for (const span of this.#axis.spans) {
      if (span.start >= index) break;
      const covered = Math.min(index - 1, span.end) - span.start + 1;
      if (covered > 0) offset += covered * (span.size - this.#axis.defaultSize);
    }
    return offset;
  }

  offset(index: number): number {
    let offset = this.#baseOffset(index);
    for (const [overrideIndex, size] of this.#overrides) {
      if (overrideIndex < index) offset += size - this.#baseSize(overrideIndex);
    }
    return offset;
  }

  size(index: number): number {
    return this.#overrides.get(index) ?? this.#baseSize(index);
  }

  indexAt(value: number): number {
    if (this.#count === 0) return -1;
    let low = 0;
    let high = this.#count;
    while (low < high) {
      const middle = Math.floor((low + high + 1) / 2);
      if (this.offset(middle) <= value) low = middle;
      else high = middle - 1;
    }
    return Math.min(this.#count - 1, Math.max(0, low));
  }

  map(value: number): number {
    if (value <= 0) return value;
    const baseTotal = this.#baseOffset(this.#count);
    if (value >= baseTotal) return this.total + value - baseTotal;
    let low = 0;
    let high = this.#count;
    while (low < high) {
      const middle = Math.floor((low + high + 1) / 2);
      if (this.#baseOffset(middle) <= value) low = middle;
      else high = middle - 1;
    }
    const originalSize = this.#baseSize(low);
    return this.offset(low) + (originalSize <= 0 ? 0 : (value - this.#baseOffset(low)) * this.size(low) / originalSize);
  }

  unmap(value: number): number {
    if (value <= 0) return value;
    const total = this.total;
    const baseTotal = this.#baseOffset(this.#count);
    if (value >= total) return baseTotal + value - total;
    const index = this.indexAt(value);
    const targetSize = this.size(index);
    return this.#baseOffset(index)
      + (targetSize <= 0 ? 0 : (value - this.offset(index)) * this.#baseSize(index) / targetSize);
  }

  get total(): number {
    return this.offset(this.#count);
  }
}

function columnLabel(index: number): string {
  let label = "";
  for (let value = index + 1; value > 0; value = Math.floor((value - 1) / 26)) {
    label = String.fromCharCode(65 + (value - 1) % 26) + label;
  }
  return label;
}

interface SheetOverrides {
  readonly rows: Map<number, number>;
  readonly columns: Map<number, number>;
}

interface SheetResizeState {
  readonly axis: "row" | "column";
  readonly index: number;
  readonly pointerId: number;
  readonly start: number;
  readonly startSize: number;
  readonly scale: number;
  size: number;
}

interface NavigationResizeState {
  readonly pointerId: number;
  readonly start: number;
  readonly startWidth: number;
}

interface ZoomAnchor {
  readonly unitIndex: number;
  readonly documentX: number;
  readonly documentY?: number;
  readonly clientX: number;
  readonly clientY?: number;
  readonly targetScale: number;
}

type PinchVerticalAnchor = Readonly<{
  unitIndex: number;
  documentY: number;
  clientY: number;
}>;

interface ContinuousPageRecord {
  readonly element: HTMLElement;
  readonly canvas: HTMLCanvasElement;
  readonly textLayer: HTMLElement;
  readonly hover: HTMLElement;
  readonly selection: HTMLElement;
  scale: number;
  renderedScale: number;
  rendered: boolean;
  rendering: boolean;
  visible: boolean;
  pixels: number;
  lastUsed: number;
  diagnostics: readonly Diagnostic[];
  textFragments: readonly RenderedTextFragment[] | undefined;
}

export class DocViewKitViewerElement extends HTMLElement {
  static readonly observedAttributes = ["theme", "lang"];

  readonly #root: ShadowRoot;
  #config: ViewerConfig = Object.freeze({});
  #features: ViewerFeatures = DEFAULT_FEATURES;
  #messages: ViewerMessages = EN;
  #locale = "en";
  #resolvedTheme: "light" | "dark" = "light";
  #status: ViewerStatus = "idle";
  #engine: OfficeEngine | undefined;
  #enginePromise: Promise<OfficeEngine> | undefined;
  #renderPixelLimit = DEFAULT_LIMITS.renderPixels;
  #document: OfficeDocument | undefined;
  #info: DocumentInfo | undefined;
  #unitIndex = 0;
  #zoom = 1;
  #fit = true;
  #pageMode: ViewerPageMode = DEFAULT_PAGE_MODE;
  #rendered: { frame: RenderResult; unit: UnitDescriptor; scale: number } | undefined;
  #searchQuery = "";
  #searchResults: readonly TextSearchResult[] = [];
  #searchIndex = -1;
  #renderRevision = 0;
  #sessionRevision = 0;
  #openController: AbortController | undefined;
  #resizeObserver: ResizeObserver | undefined;
  #compactLayout: boolean | undefined;
  #themeMedia: MediaQueryList | undefined;
  #thumbnailObserver: IntersectionObserver | undefined;
  #thumbnailRequests = new Map<number, AbortController>();
  #prefetchedPages = new Map<number, OfficeDocument>();
  #pagePrefetchTimer: ReturnType<typeof setTimeout> | undefined;
  #printController: AbortController | undefined;
  #printFrame: HTMLIFrameElement | undefined;
  #printUrls: string[] = [];
  #printing = false;
  #navigationStart = -1;
  #navigationEnd = -1;
  #navigationMode: ViewerNavigationMode = "thumbnails";
  #outlineSelection = -1;
  #navigationTransitionRevision = 0;
  #navigationTransitioning = false;
  #navigationResize: NavigationResizeState | undefined;
  #sheetSizes = new Map<number, SheetOverrides>();
  #sheetResize: SheetResizeState | undefined;
  #sheetScrollFrame: number | undefined;
  #sheetScrollRendering = false;
  #sheetTileKey = "";
  #sheetScrollPosition = { x: 0, y: 0, dx: 0, dy: 0 };
  readonly #sheetTiles = new Map<string, {
    viewport: Rect;
    used: boolean;
    controller: AbortController;
    promise: Promise<void>;
    canvas?: HTMLCanvasElement;
    frame?: RenderResult;
  }>();
  #continuousObserver: IntersectionObserver | undefined;
  #continuousNavigationObserver: IntersectionObserver | undefined;
  #continuousPages = new Map<number, ContinuousPageRecord>();
  #continuousCenterPages = new Set<number>();
  #continuousRevision = 0;
  #continuousPixels = 0;
  #zoomAnchor: ZoomAnchor | undefined;
  #singleLayoutFrame: number | undefined;
  #singleLayoutPending = false;
  #singleLayoutRendering = false;
  #continuousLayoutFrame: number | undefined;
  #wheelDelta = 0;
  #wheelDirection = 0;
  #wheelPageLocked = false;
  #wheelResetTimer: ReturnType<typeof setTimeout> | undefined;
  #pinchZoomTimer: ReturnType<typeof setTimeout> | undefined;
  #pinchPreviewBaseScale = 0;
  #pinchPreviewScale = 0;
  #pinchSettleRect: Readonly<{ left: number; width: number }> | undefined;
  #pinchSettleAnimation: Animation | undefined;
  #pinchVerticalAnchor: PinchVerticalAnchor | undefined;
  #pageAnimations: Animation[] = [];
  #outgoingPage: HTMLCanvasElement | undefined;
  #selectionBounds: Rect | undefined;
  #selectionUnitIndex = -1;
  #selectionObjectId: string | undefined;
  #selectionText: string | undefined;
  #hoverBounds: Rect | undefined;
  #hoverUnitIndex = -1;
  #hoverPoint: Readonly<{ unitIndex: number; x: number; y: number }> | undefined;
  #hoverTimer: ReturnType<typeof setTimeout> | undefined;
  #hoverRevision = 0;
  #destroyed = false;
  #searchTimer: ReturnType<typeof setTimeout> | undefined;
  #interactionRevision = 0;
  #hitRevision = 0;
  #textSelectionFrame: number | undefined;
  #textSelectionAnchor: Readonly<{ fragment: HTMLElement; offset: number }> | undefined;
  #pointerTextSelection: Readonly<{
    anchor: Readonly<{ fragment: HTMLElement; offset: number }>;
    focus: Readonly<{ fragment: HTMLElement; offset: number }>;
  }> | undefined;
  #pointerSelectionRectangles: DOMRect[] = [];
  readonly #selectionChange = (): void => this.#queueTextSelectionPaint();

  constructor() {
    super();
    this.#root = this.attachShadow({ mode: "open", delegatesFocus: true });
    this.#root.innerHTML = TEMPLATE;
    this.#bindEvents();
    this.#setStatus("idle");
  }

  get config(): ViewerConfig {
    return this.#config;
  }

  set config(value: ViewerConfig) {
    if (value === null || typeof value !== "object") throw new TypeError("Viewer config must be an object");
    this.#config = Object.freeze({ ...value });
    this.#applyConfig();
  }

  get state(): Readonly<ViewerState> {
    const info = this.#info;
    return Object.freeze({
      status: this.#status,
      ...(info === undefined ? {} : { info }),
      currentUnitIndex: this.#unitIndex,
      unitCount: info?.units.length ?? 0,
      zoom: this.#zoom,
      locale: this.#locale,
      theme: this.#resolvedTheme,
      pageMode: this.#pageMode,
      license: OPEN_SOURCE_LICENSE,
      search: Object.freeze({
        query: this.#searchQuery,
        current: this.#searchIndex + 1,
        total: this.#searchResults.length,
      }),
    });
  }

  connectedCallback(): void {
    if (this.#destroyed) return;
    globalThis.document.addEventListener("selectionchange", this.#selectionChange);
    this.#applyConfig();
    this.#resizeObserver ??= new ResizeObserver(() => {
      const compactLayout = this.clientWidth <= 720;
      if (compactLayout !== this.#compactLayout) {
        this.#compactLayout = compactLayout;
        this.#updateNavigationVisibility();
      }
      if (this.#document === undefined || (!this.#fit && this.#rendered?.unit.type !== "sheet")) return;
      if (this.#navigationTransitioning) {
        if (this.#pageMode === "continuous") this.#previewContinuousNavigationLayout();
        else this.#previewSingleLayout();
        return;
      }
      if (this.#pageMode === "continuous") this.#queueContinuousLayout();
      else this.#queueSingleLayout();
    });
    this.#resizeObserver.observe(this.#workspace());
  }

  disconnectedCallback(): void {
    globalThis.document.removeEventListener("selectionchange", this.#selectionChange);
    if (this.#textSelectionFrame !== undefined) cancelAnimationFrame(this.#textSelectionFrame);
    this.#textSelectionFrame = undefined;
    this.#resizeObserver?.disconnect();
    this.#printController?.abort();
    this.#printController = undefined;
    this.#clearPrintView();
    this.#cancelLayoutFrames();
    this.#clearObjectHover();
    this.#resetWheelGesture();
    this.#finishNavigationResize();
    this.#finishSheetResize();
  }

  attributeChangedCallback(name: string): void {
    if (name === "theme") this.#applyTheme();
    else this.#applyConfig();
  }

  async open(
    source: File | Blob | ArrayBuffer | Uint8Array,
    options: ViewerOpenOptions = {},
  ): Promise<DocumentInfo> {
    this.#assertUsable();
    await this.close();
    const session = ++this.#sessionRevision;
    const controller = new AbortController();
    this.#openController = controller;
    this.#setStatus("loading");
    try {
      let bytes = source instanceof Blob ? await source.arrayBuffer() : source;
      if (session !== this.#sessionRevision) throw new DOMException("Superseded", "AbortError");
      const engine = await this.#getEngine();
      const fileName = source instanceof File && source.name.length > 0 ? source.name : undefined;
      let password = options.password;
      let rejectedPasswords = 0;
      let document: OfficeDocument;
      for (;;) {
        try {
          document = await engine.open(bytes, {
            signal: controller.signal,
            ...(source instanceof Blob ? { transferInput: true } : {}),
            ...(password === undefined ? {} : { password }),
            ...(fileName === undefined ? {} : { fileName }),
          });
          password = undefined;
          break;
        } catch (cause) {
          const passwordFailure = cause instanceof OfficeEngineError
            && (cause.code === "PDF_PASSWORD_REQUIRED" || cause.code === "PDF_PASSWORD_INCORRECT");
          if (!passwordFailure) throw cause;
          if (cause.code === "PDF_PASSWORD_INCORRECT") rejectedPasswords += 1;
          if (source instanceof Blob && bytes instanceof ArrayBuffer && bytes.byteLength === 0) {
            bytes = await source.arrayBuffer();
          }
          password = undefined;
          if (rejectedPasswords >= 3) throw cause;
          const supplied = await this.#requestPassword(cause.code === "PDF_PASSWORD_INCORRECT");
          if (supplied === undefined) throw cause;
          if (session !== this.#sessionRevision) throw new DOMException("Superseded", "AbortError");
          password = supplied;
        }
      }
      if (session !== this.#sessionRevision) {
        document.close();
        throw new DOMException("Superseded", "AbortError");
      }
      this.#document = document;
      this.#info = document.info;
      this.#unitIndex = 0;
      const spreadsheetDefaultZoom = this.#config.initialZoom === undefined && document.info.kind === "spreadsheet";
      this.#fit = !spreadsheetDefaultZoom
        && (this.#config.initialZoom === undefined || this.#config.initialZoom === "fit");
      this.#zoom = typeof this.#config.initialZoom === "number"
        ? Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, this.#config.initialZoom))
        : 1;
      this.#buildNavigation();
      this.#updateFormulaBar();
      const requestedMode = this.#config.initialPageMode ?? DEFAULT_PAGE_MODE;
      this.#pageMode = requestedMode === "continuous" && this.#supportsContinuous() ? "continuous" : "single";
      this.#updatePageModeControls();
      if (this.#pageMode === "continuous") await this.#buildContinuousPages();
      else await this.#renderCurrent(0, true);
      this.#setStatus("ready");
      this.#setMessage(document.info.format.toUpperCase());
      this.#emit("docviewkit-ready", { info: document.info });
      this.#emitState();
      return document.info;
    } catch (cause) {
      if (session === this.#sessionRevision && !(cause instanceof DOMException && cause.name === "AbortError")) {
        this.#setStatus("error", errorMessage(cause));
        this.#emit("docviewkit-error", { cause });
      }
      throw cause;
    } finally {
      if (this.#openController === controller) this.#openController = undefined;
    }
  }

  async reveal(target: ViewerTarget): Promise<RevealResult> {
    this.#assertUsable();
    const document = this.#document;
    if (document === undefined) throw new Error("No document is open");
    let unitIndex: number | undefined;
    let object: DocumentObject | undefined;
    let region: Rect | undefined;
    if (target.kind === "unit") unitIndex = target.unitIndex;
    else if (target.kind === "region") {
      unitIndex = target.unitIndex;
      region = target.region;
    } else {
      const objectId = target.kind === "object" ? target.objectId : target.objectId;
      if (objectId !== undefined) object = await document.getObject(objectId);
      unitIndex = object?.unitIndex ?? (target.kind === "source" ? target.unitIndex : undefined);
      region = object?.bounds ?? (target.kind === "source" ? target.region : undefined);
    }
    const resolvedUnit = unitIndex ?? 0;
    this.#assertUnit(resolvedUnit);
    await this.#selectUnit(resolvedUnit);
    if (region !== undefined) this.#showSelection(region, true, object?.id, object?.text);
    const exact = object !== undefined || target.kind === "unit" || target.kind === "region";
    const result: RevealResult = Object.freeze({
      unitIndex: resolvedUnit,
      quality: exact ? "exact" : "approximate",
      ...(object === undefined ? {} : { objectId: object.id }),
      ...(region === undefined ? {} : { region: copyRect(region) }),
    });
    this.#setMessage(exact ? this.#messages.exactLocation : this.#messages.approximateLocation);
    this.#emit("docviewkit-statechange", { state: this.state, reveal: result });
    return result;
  }

  async close(): Promise<void> {
    this.#sessionRevision++;
    const passwordDialog = this.#root.querySelector<HTMLDialogElement>(".password-dialog");
    if (passwordDialog?.open === true) passwordDialog.close("cancel");
    this.#openController?.abort();
    this.#openController = undefined;
    this.#printController?.abort();
    this.#printController = undefined;
    this.#clearPrintView();
    this.#cancelLayoutFrames();
    this.#clearSearch();
    this.#clearObjectHover();
    this.#resetWheelGesture();
    this.#closeRenderedFrame();
    this.#clearSheetTiles();
    this.#clearContinuousPages();
    this.#clearThumbnails();
    this.#sheetSizes.clear();
    this.#finishNavigationResize();
    this.#finishSheetResize();
    this.#document?.close();
    this.#document = undefined;
    this.#info = undefined;
    this.#unitIndex = 0;
    this.#root.querySelector(".unit-list")?.replaceChildren();
    this.#root.querySelector(".sheet-tabs")?.replaceChildren();
    this.#element<HTMLElement>(".shell").dataset.spreadsheet = "false";
    this.#updateFormulaBar();
    this.#resetSheetSurface();
    this.#setStatus(this.#destroyed ? "destroyed" : "idle");
    this.#emitState();
  }

  destroy(): void {
    if (this.#destroyed) return;
    this.#destroyed = true;
    void this.close();
    this.#engine?.close();
    this.#engine = undefined;
    this.#resizeObserver?.disconnect();
    this.#themeMedia?.removeEventListener("change", this.#handleThemeChange);
    this.#thumbnailObserver?.disconnect();
    this.#cancelLayoutFrames();
    document.removeEventListener("fullscreenchange", this.#handleFullscreenChange);
    this.#setStatus("destroyed");
  }

  async #getEngine(): Promise<OfficeEngine> {
    if (this.#engine !== undefined) return this.#engine;
    const engineOptions = this.#config.engine ?? {};
    const renderPixelLimit = engineOptions.limits?.renderPixels ?? DEFAULT_LIMITS.renderPixels;
    const pending = this.#enginePromise ??= createOfficeEngine(engineOptions);
    try {
      const engine = await pending;
      if (this.#destroyed) {
        engine.close();
        throw new Error("Viewer has been destroyed");
      }
      this.#engine = engine;
      this.#renderPixelLimit = renderPixelLimit;
      return engine;
    } finally {
      if (this.#enginePromise === pending) this.#enginePromise = undefined;
    }
  }

  #requestPassword(incorrect: boolean): Promise<string | undefined> {
    const dialog = this.#element<HTMLDialogElement>(".password-dialog");
    const input = this.#element<HTMLInputElement>(".password-input");
    dialog.returnValue = "";
    input.value = "";
    this.#element<HTMLElement>(".password-heading").textContent = incorrect
      ? this.#messages.passwordIncorrect
      : this.#messages.passwordRequired;
    dialog.showModal();
    queueMicrotask(() => input.focus());
    return new Promise((resolve) => {
      dialog.addEventListener("close", () => {
        const password = dialog.returnValue === "unlock" ? input.value : undefined;
        input.value = "";
        resolve(password);
      }, { once: true });
    });
  }

  #assertUsable(): void {
    if (this.#destroyed) throw new Error("Viewer has been destroyed");
  }

  #applyConfig(): void {
    const localeSource = this.#config.locale ?? this.getAttribute("lang") ?? document.documentElement.lang ?? "auto";
    this.#locale = normalizedLocale(localeSource);
    this.#messages = messageCatalog(this.#locale, this.#config.messages);
    const {
      textSelection: legacyTextSelection,
      objectSelection: legacyObjectSelection,
      ...configuredFeatures
    } = this.#config.features ?? {};
    const configuredInteractionMode = configuredFeatures.interactionMode;
    const interactionMode = configuredInteractionMode === "object"
      || configuredInteractionMode === "display"
      || configuredInteractionMode === "text"
      ? configuredInteractionMode
      : legacyTextSelection === true
        ? "text"
        : legacyObjectSelection === false
          ? "display"
          : "object";
    this.#features = Object.freeze({ ...DEFAULT_FEATURES, ...configuredFeatures, interactionMode });
    this.#interactionRevision += 1;
    this.#hitRevision += 1;
    this.#hoverRevision += 1;
    const shell = this.#element<HTMLElement>(".shell");
    shell.lang = this.#locale;
    shell.dir = isRtl(this.#locale) ? "rtl" : "ltr";
    shell.dataset.interactionMode = this.#features.interactionMode;
    this.#applyTheme();
    this.#applyFeatureVisibility();
    this.#applyInteractionFeatures();
    this.#applyMinimalMode();
    this.#updateNavigationVisibility();
    this.#updateControls();
  }

  #applyTheme(): void {
    const previous = this.#resolvedTheme;
    const configuredTheme = this.getAttribute("theme") ?? this.#config.theme ?? "auto";
    const theme: ViewerTheme = configuredTheme === "light" || configuredTheme === "dark" ? configuredTheme : "auto";
    this.#themeMedia ??= matchMedia("(prefers-color-scheme: dark)");
    this.#themeMedia.removeEventListener("change", this.#handleThemeChange);
    this.#themeMedia.addEventListener("change", this.#handleThemeChange);
    this.#resolvedTheme = theme === "auto" ? (this.#themeMedia.matches ? "dark" : "light") : theme;
    this.dataset.resolvedTheme = this.#resolvedTheme;
    this.#localize();
    if (previous !== this.#resolvedTheme) this.#emitState();
  }

  readonly #handleThemeChange = (): void => this.#applyTheme();

  #applyMinimalMode(): void {
    const enabled = this.#config.minimal === true;
    this.#element<HTMLElement>(".shell").dataset.minimal = String(enabled);
  }

  #localize(): void {
    const labels: Readonly<Record<string, keyof ViewerMessages>> = {
      navigation: "navigation",
      previous: "previous",
      next: "next",
      search: "search",
      "previous-result": "previousResult",
      "next-result": "nextResult",
      "close-search": "closeSearch",
      "zoom-out": "zoomOut",
      "zoom-in": "zoomIn",
      fit: "fit",
      print: "print",
      fullscreen: document.fullscreenElement === this ? "exitFullscreen" : "fullscreen",
      diagnostics: "diagnostics",
      "page-mode": "continuousPages",
      "close-diagnostics": "close",
      "mobile-previous": "previous",
      "mobile-next": "next",
    };
    for (const [action, key] of Object.entries(labels)) {
      const label = this.#messages[key];
      this.#root.querySelectorAll<HTMLButtonElement>(`[data-action="${action}"]`).forEach((button) => {
        button.ariaLabel = label;
        button.title = label;
      });
    }
    const theme = this.getAttribute("theme") ?? this.#config.theme ?? "auto";
    const themeKey = theme === "light" ? "themeLight" : theme === "dark" ? "themeDark" : "themeAuto";
    const themeButton = this.#element<HTMLButtonElement>('[data-action="theme"]');
    themeButton.ariaLabel = this.#messages[themeKey];
    themeButton.title = this.#messages[themeKey];
    themeButton.querySelector("svg")!.innerHTML = ICONS[themeKey];
    const input = this.#element<HTMLInputElement>(".search-input");
    input.ariaLabel = this.#messages.search;
    input.placeholder = this.#messages.searchPlaceholder;
    const passwordInput = this.#element<HTMLInputElement>(".password-input");
    passwordInput.ariaLabel = this.#messages.passwordLabel;
    this.#element<HTMLElement>(".password-label").textContent = this.#messages.passwordLabel;
    this.#element<HTMLButtonElement>('[data-action="password-cancel"]').textContent = this.#messages.cancel;
    this.#element<HTMLButtonElement>('.password-action[value="unlock"]').textContent = this.#messages.unlock;
    this.#element<HTMLElement>(".workspace").ariaLabel = this.#messages.documentCanvas;
    this.#element<HTMLElement>(".navigation").ariaLabel = this.#messages.documentNavigation;
    const outlineMode = this.#element<HTMLButtonElement>('[data-action="navigation-outline"]');
    outlineMode.ariaLabel = this.#messages.outline;
    outlineMode.title = this.#messages.outline;
    const thumbnailMode = this.#element<HTMLButtonElement>('[data-action="navigation-thumbnails"]');
    thumbnailMode.ariaLabel = this.#messages.thumbnails;
    thumbnailMode.title = this.#messages.thumbnails;
    const interactionSwitcher = this.#element<HTMLElement>(".interaction-switcher");
    interactionSwitcher.ariaLabel = this.#messages.interactionModes;
    const interactionLabels: Readonly<Record<ViewerInteractionMode, keyof ViewerMessages>> = {
      object: "objectMode",
      display: "displayMode",
      text: "textMode",
    };
    this.#root.querySelectorAll<HTMLButtonElement>('[data-action="interaction-mode"]').forEach((button) => {
      const mode = button.dataset.interactionMode as ViewerInteractionMode;
      const label = this.#messages[interactionLabels[mode]];
      button.ariaLabel = label;
      button.title = label;
      button.querySelector<HTMLElement>("span")!.textContent = label;
    });
    this.#element<HTMLInputElement>(".zoom-input").ariaLabel = this.#messages.zoom;
    const navigationResizer = this.#element<HTMLElement>(".navigation-resizer");
    navigationResizer.ariaLabel = this.#messages.resizeNavigation;
    const navigationWidth = this.#element<HTMLElement>(".navigation").getBoundingClientRect().width;
    navigationResizer.setAttribute("aria-valuenow", String(Math.round(navigationWidth || 176)));
    this.#element<HTMLElement>(".sheet-tabs").ariaLabel = this.#messages.sheetTabs;
    this.#element<HTMLElement>(".formula-bar").ariaLabel = this.#messages.formulaBar;
    this.#element<HTMLOutputElement>(".cell-address").ariaLabel = this.#messages.cellAddress;
    this.#element<HTMLElement>(".sheet-column-headers").ariaLabel = this.#messages.columnHeaders;
    this.#element<HTMLElement>(".sheet-row-headers").ariaLabel = this.#messages.rowHeaders;
    this.#element<HTMLElement>(".empty .state-title").textContent = this.#messages.empty;
    this.#element<HTMLElement>(".empty .state-hint").textContent = this.#messages.emptyHint;
    this.#element<HTMLElement>(".loading .state-title").textContent = this.#messages.loading;
    this.#element<HTMLElement>(".diagnostic-heading strong").textContent = this.#messages.diagnostics;
    this.#renderDiagnostics();
  }

  #applyFeatureVisibility(): void {
    const searchExpanded = !this.#element<HTMLElement>(".search").hidden;
    const featureActions: Readonly<Partial<Record<keyof ViewerFeatures, readonly string[]>>> = {
      search: ["search"],
      navigation: ["navigation"],
      zoom: ["zoom-out", "zoom-in", "fit"],
      print: ["print"],
      fullscreen: ["fullscreen"],
      diagnostics: ["diagnostics"],
      pageMode: ["page-mode"],
    };
    for (const [feature, actions] of Object.entries(featureActions) as [keyof ViewerFeatures, readonly string[]][]) {
      for (const action of actions) {
        this.#root.querySelectorAll<HTMLElement>(`[data-action="${action}"]`).forEach((element) => {
          element.hidden = !this.#features[feature] || (action === "search" && searchExpanded);
        });
      }
    }
    this.#element<HTMLElement>(".zoom-value").hidden = !this.#features.zoom;
    this.#element<HTMLElement>(".interaction-switcher").hidden = !this.#features.interactionModeSwitcher;
    if (!this.#features.search) {
      this.#element<HTMLElement>(".search").hidden = true;
      this.#clearSearch();
    }
    if (!this.#features.diagnostics) this.#element<HTMLElement>(".diagnostic-panel").hidden = true;
  }

  #applyInteractionFeatures(): void {
    this.#clearObjectHover();
    this.#root.querySelectorAll<HTMLButtonElement>('[data-action="interaction-mode"]').forEach((button) => {
      button.setAttribute("aria-pressed", String(button.dataset.interactionMode === this.#features.interactionMode));
    });
    if (this.#features.interactionMode !== "text") {
      this.#textSelectionAnchor = undefined;
      this.#pointerSelectionRectangles = [];
      this.#root.querySelectorAll<HTMLElement>(".text-layer").forEach((layer) => {
        layer.hidden = true;
        layer.replaceChildren();
      });
      if (this.#features.interactionMode !== "object") this.#hideSelection();
      this.#refreshSearchHighlight();
      return;
    }
    this.#hideSelection();
    void this.#refreshTextLayers();
  }

  #bindEvents(): void {
    this.#root.addEventListener("selectionchange", this.#selectionChange);
    this.#root.addEventListener("click", (event) => {
      const button = event.target instanceof Element ? event.target.closest<HTMLButtonElement>("[data-action]") : null;
      if (button === null || button.disabled) return;
      const action = button.dataset.action;
      if (action === "navigation") this.#toggleNavigation();
      else if (action === "navigation-outline") this.#setNavigationMode("outline");
      else if (action === "navigation-thumbnails") this.#setNavigationMode("thumbnails");
      else if (action === "previous" || action === "mobile-previous") void this.#selectUnit(this.#unitIndex - 1);
      else if (action === "next" || action === "mobile-next") void this.#selectUnit(this.#unitIndex + 1);
      else if (action === "search") this.#openSearch();
      else if (action === "close-search") this.#closeSearch();
      else if (action === "previous-result") void this.#stepSearch(-1);
      else if (action === "next-result") void this.#stepSearch(1);
      else if (action === "zoom-out") this.#setZoom(this.#zoom - .25);
      else if (action === "zoom-in") this.#setZoom(this.#zoom + .25);
      else if (action === "fit") this.#fitView();
      else if (action === "interaction-mode") this.#setInteractionMode(button.dataset.interactionMode);
      else if (action === "page-mode") void this.#setPageMode(this.#pageMode === "continuous" ? "single" : "continuous");
      else if (action === "print") void this.#printDocument();
      else if (action === "theme") {
        const theme = this.getAttribute("theme") ?? this.#config.theme ?? "auto";
        this.setAttribute("theme", theme === "light" ? "dark" : theme === "dark" ? "auto" : "light");
      } else if (action === "fullscreen") void this.#toggleFullscreen();
      else if (action === "diagnostics") this.#toggleDiagnostics();
      else if (action === "close-diagnostics") this.#toggleDiagnostics(false);
      else if (action === "password-cancel") this.#element<HTMLDialogElement>(".password-dialog").close("cancel");
    });
    this.#element<HTMLFormElement>(".search").addEventListener("submit", (event) => {
      event.preventDefault();
      void this.#stepSearch(1);
    });
    this.#element<HTMLInputElement>(".search-input").addEventListener("input", (event) => {
      if (this.#searchTimer !== undefined) clearTimeout(this.#searchTimer);
      this.#searchQuery = (event.currentTarget as HTMLInputElement).value.trim();
      this.#searchResults = [];
      this.#searchIndex = -1;
      this.#clearSearchHighlight();
      this.#hideSelection();
      this.#updateSearchPosition();
      this.#searchTimer = setTimeout(() => void this.#performSearch(), 180);
    });
    const zoomInput = this.#element<HTMLInputElement>(".zoom-input");
    zoomInput.addEventListener("focus", () => zoomInput.select());
    zoomInput.addEventListener("change", () => this.#commitZoomInput());
    zoomInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        this.#commitZoomInput();
        zoomInput.select();
      } else if (event.key === "Escape") {
        event.preventDefault();
        this.#syncZoomInput();
        zoomInput.select();
      }
    });
    this.#root.addEventListener("keydown", (event) => this.#handleKeydown(event as KeyboardEvent));
    for (const type of ["copy", "cut"] as const) {
      this.#root.addEventListener(type, (event) => this.#copySelectedObject(event as ClipboardEvent));
    }
    const workspace = this.#workspace();
    workspace.addEventListener("wheel", (event) => this.#handleWheel(event), { passive: false });
    workspace.addEventListener("scroll", () => {
      this.#queueSheetScroll();
      this.#clearObjectHover();
    }, { passive: true });
    workspace.addEventListener("click", (event) => void this.#selectObjectAt(event));
    workspace.addEventListener("pointermove", (event) => this.#queueObjectHover(event), { passive: true });
    workspace.addEventListener("pointerleave", () => this.#clearObjectHover());
    this.#element<HTMLElement>(".navigation").addEventListener("scroll", () => {
      if (this.#navigationMode !== "thumbnails") return;
      const navigation = this.#element<HTMLElement>(".navigation");
      const center = Math.floor(navigation.scrollTop / NAVIGATION_ITEM_HEIGHT);
      if (center < this.#navigationStart + 8 || center >= this.#navigationEnd - 8) {
        this.#renderNavigationWindow(center);
      }
    }, { passive: true });
    this.#root.addEventListener("pointerdown", (event) => {
      this.#beginNavigationResize(event as PointerEvent);
      this.#beginSheetResize(event as PointerEvent);
      this.#beginTextSelection(event as PointerEvent);
    });
    this.#root.addEventListener("pointermove", (event) => {
      this.#updateNavigationResize(event as PointerEvent);
      this.#updateSheetResize(event as PointerEvent);
      this.#updateTextSelection(event as PointerEvent);
    });
    this.#root.addEventListener("pointerup", (event) => {
      this.#finishNavigationResize(event as PointerEvent);
      this.#finishSheetResize(event as PointerEvent);
      this.#finishTextSelection(event as PointerEvent);
    });
    this.#root.addEventListener("pointercancel", (event) => {
      this.#finishNavigationResize(event as PointerEvent);
      this.#finishSheetResize(event as PointerEvent);
    });
    this.#root.addEventListener("selectstart", (event) => {
      const target = event.target instanceof Element ? event.target : null;
      if (this.#navigationResize !== undefined || this.#sheetResize !== undefined
        || (target !== null && target.closest(".sheet-header-axis, .sheet-header-corner") !== null)) {
        event.preventDefault();
        this.#clearNativeSelection();
      }
    });
    this.#root.addEventListener("dblclick", (event) => void this.#autoFitSheetSize(event));
    document.addEventListener("fullscreenchange", this.#handleFullscreenChange);
  }

  readonly #handleFullscreenChange = (): void => this.#localize();

  #setInteractionMode(value: string | undefined): void {
    if (value !== "object" && value !== "display" && value !== "text") return;
    this.config = {
      ...this.#config,
      features: { ...this.#config.features, interactionMode: value },
    };
  }

  #handleKeydown(event: KeyboardEvent): void {
    const target = event.composedPath()[0];
    const editing = (target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement)
      && !target.classList.contains("object-clipboard");
    if (target instanceof HTMLElement && target.classList.contains("navigation-resizer")
      && (event.key === "ArrowLeft" || event.key === "ArrowRight")) {
      event.preventDefault();
      const rtl = this.#element<HTMLElement>(".shell").dir === "rtl";
      const direction = (event.key === "ArrowRight" ? 1 : -1) * (rtl ? -1 : 1);
      const step = event.shiftKey ? 24 : 8;
      this.#setNavigationWidth(this.#element<HTMLElement>(".navigation").getBoundingClientRect().width
        + direction * step);
    } else if (target instanceof HTMLElement && target.classList.contains("sheet-header-resizer")
      && ["ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown"].includes(event.key)) {
      const unit = this.#info?.units[this.#unitIndex];
      const axis = target.dataset.axis;
      const index = Number(target.dataset.index);
      if (unit?.type !== "sheet" || (axis !== "row" && axis !== "column") || !Number.isInteger(index)) return;
      event.preventDefault();
      const layout = this.#sheetLayout(unit);
      const current = (axis === "column" ? layout.columns : layout.rows).size(index);
      const direction = event.key === "ArrowRight" || event.key === "ArrowDown" ? 1 : -1;
      const minimum = axis === "column" ? MIN_COLUMN_WIDTH : MIN_ROW_HEIGHT;
      const overrides = this.#sheetOverrides(unit.index);
      (axis === "column" ? overrides.columns : overrides.rows).set(index,
        Math.max(minimum, current + direction * (event.shiftKey ? 10 : 1)));
      void this.#renderCurrent();
    } else if (target instanceof HTMLElement && target.classList.contains("sheet-tab")
      && (event.key === "ArrowLeft" || event.key === "ArrowRight")) {
      event.preventDefault();
      void this.#selectUnit(this.#unitIndex + (event.key === "ArrowRight" ? 1 : -1));
    } else if (target instanceof HTMLElement && target.closest(".unit-button, .outline-button") !== null
      && (event.key === "ArrowUp" || event.key === "ArrowDown")) {
      event.preventDefault();
      this.#stepNavigationUnit(event.key === "ArrowDown" ? 1 : -1);
    } else if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "f" && this.#features.search) {
      event.preventDefault();
      this.#openSearch();
    } else if (!editing && (event.key === "PageUp" || event.key === "ArrowLeft")) {
      event.preventDefault();
      void this.#selectUnit(this.#unitIndex - 1);
    } else if (!editing && (event.key === "PageDown" || event.key === "ArrowRight")) {
      event.preventDefault();
      void this.#selectUnit(this.#unitIndex + 1);
    } else if (!editing && (event.key === "+" || event.key === "=")) {
      event.preventDefault();
      this.#setZoom(this.#zoom + .25);
    } else if (!editing && event.key === "-") {
      event.preventDefault();
      this.#setZoom(this.#zoom - .25);
    } else if (!editing && event.key === "0") {
      event.preventDefault();
      this.#fitView();
    } else if (event.key === "Escape" && !this.#element<HTMLElement>(".search").hidden) {
      event.preventDefault();
      this.#closeSearch();
    } else if (!editing && event.key === "Escape" && this.#selectionBounds !== undefined) {
      event.preventDefault();
      this.#hideSelection();
    }
  }

  #copySelectedObject(event: ClipboardEvent): void {
    const target = event.composedPath()[0];
    if ((target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement)
      && !target.classList.contains("object-clipboard")) return;
    if (event.clipboardData === null) return;
    if (this.#features.interactionMode === "text") {
      const selection = (this.#root as ShadowRoot & { getSelection?: () => Selection | null }).getSelection?.()
        ?? globalThis.document.getSelection();
      const text = selection === null ? undefined : this.#selectedTextContent(selection);
      if (text === undefined) return;
      event.clipboardData.setData("text/plain", text);
      event.preventDefault();
      return;
    }
    if (this.#features.interactionMode !== "object" || this.#selectionText === undefined) return;
    event.clipboardData.setData("text/plain", this.#selectionText);
    event.preventDefault();
  }

  #handleWheel(event: WheelEvent): void {
    if (event.ctrlKey) {
      event.preventDefault();
      this.#resetWheelGesture();
      if (this.#document !== undefined && this.#features.zoom) this.#handlePinchZoom(event);
      return;
    }
    if (this.#pageMode !== "single" || this.#document === undefined
      || this.#info?.units[this.#unitIndex]?.type === "sheet"
      || Math.abs(event.deltaY) <= Math.abs(event.deltaX)) {
      this.#resetWheelGesture();
      return;
    }
    const workspace = this.#workspace();
    const unit = event.deltaMode === 1 ? 16 : event.deltaMode === 2 ? workspace.clientHeight : 1;
    const delta = event.deltaY * unit;
    const direction = Math.sign(delta);
    if (direction !== this.#wheelDirection) {
      this.#wheelDelta = 0;
      this.#wheelPageLocked = false;
      this.#wheelDirection = direction;
    }
    if (this.#wheelPageLocked) {
      event.preventDefault();
      this.#scheduleWheelReset();
      return;
    }
    const maxScrollTop = Math.max(0, workspace.scrollHeight - workspace.clientHeight);
    const atBoundary = direction > 0
      ? workspace.scrollTop >= maxScrollTop - 2
      : workspace.scrollTop <= 2;
    const unitCount = this.#info?.units.length ?? 0;
    const canNavigate = direction > 0 ? this.#unitIndex < unitCount - 1 : this.#unitIndex > 0;
    if (!atBoundary || !canNavigate) {
      this.#resetWheelGesture();
      return;
    }
    event.preventDefault();
    this.#scheduleWheelReset();
    this.#wheelDelta += delta;
    const threshold = Math.min(120, Math.max(48, workspace.clientHeight * .08));
    if (Math.abs(this.#wheelDelta) < threshold) return;
    this.#wheelDelta = 0;
    this.#wheelPageLocked = true;
    void this.#selectUnit(this.#unitIndex + direction);
  }

  #scheduleWheelReset(): void {
    if (this.#wheelResetTimer !== undefined) clearTimeout(this.#wheelResetTimer);
    this.#wheelResetTimer = setTimeout(() => this.#resetWheelGesture(), 260);
  }

  #handlePinchZoom(event: WheelEvent): void {
    const workspace = this.#workspace();
    const unit = event.deltaMode === 1 ? 16 : event.deltaMode === 2 ? workspace.clientHeight : 1;
    const delta = Math.min(100, Math.max(-100, event.deltaY * unit));
    const factor = Math.min(1.15, Math.max(.85, Math.exp(-delta * .01)));
    const target = Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, this.#zoom * factor));
    const continuing = this.#pinchZoomTimer !== undefined;
    if (!continuing) this.#pinchVerticalAnchor = this.#capturePinchVerticalAnchor(event);
    const zoom = this.#setZoom(target, false, event, true);
    if (zoom !== undefined) this.#previewPinchZoom(zoom, event, continuing);
    if (zoom !== undefined || continuing) this.#schedulePinchRender();
  }

  #previewPinchZoom(targetScale: number, event: WheelEvent, continuing: boolean): void {
    this.#pinchSettleAnimation?.cancel();
    this.#pinchSettleAnimation = undefined;
    if (this.#pageMode === "continuous") {
      const units = this.#info?.units ?? [];
      for (const unit of units) {
        const record = this.#continuousPages.get(unit.index);
        if (record === undefined) continue;
        record.scale = targetScale;
        record.element.style.width = `${unit.width * targetScale}px`;
        record.element.style.height = `${unit.height * targetScale}px`;
        const textFactor = record.renderedScale > 0 ? targetScale / record.renderedScale : 1;
        record.textLayer.style.transformOrigin = "top left";
        record.textLayer.style.transform = `scale(${textFactor})`;
      }
      this.#refreshObjectHighlights();
      this.#applyContinuousZoomAnchor();
      return;
    }
    const rendered = this.#rendered;
    if (rendered === undefined) return;
    const stage = this.#element<HTMLElement>(".stage");
    this.#cancelPageAnimation();
    if (rendered.unit.type !== "sheet") {
      this.#previewSinglePageScale(targetScale);
      this.#pinchPreviewScale = targetScale;
      this.#applySingleZoomAnchor(rendered.frame, targetScale);
      return;
    }
    if (!continuing) {
      this.#pinchPreviewBaseScale = rendered.scale;
      stage.style.removeProperty("transform");
      const rect = stage.getBoundingClientRect();
      const originX = Math.min(rect.width, Math.max(0, event.clientX - rect.left));
      const originY = Math.min(rect.height, Math.max(0, event.clientY - rect.top));
      stage.style.transformOrigin = `${originX}px ${originY}px`;
      stage.style.willChange = "transform";
    }
    const factor = targetScale / this.#pinchPreviewBaseScale;
    if (!Number.isFinite(factor) || Math.abs(factor - 1) < .005) {
      stage.style.removeProperty("transform");
      return;
    }
    stage.style.transform = `scale(${factor})`;
  }

  #previewSinglePageScale(targetScale: number): void {
    const rendered = this.#rendered;
    if (rendered === undefined) return;
    const cssWidth = rendered.frame.viewport.width * targetScale;
    const cssHeight = rendered.frame.viewport.height * targetScale;
    const canvas = this.#element<HTMLCanvasElement>(".surface");
    canvas.style.width = `${cssWidth}px`;
    canvas.style.height = `${cssHeight}px`;
    const stage = this.#element<HTMLElement>(".stage");
    stage.style.width = `${cssWidth}px`;
    stage.style.height = `${cssHeight}px`;
    const textLayer = this.#element<HTMLElement>(".text-layer");
    const textFactor = targetScale / rendered.scale;
    textLayer.style.transformOrigin = "top left";
    textLayer.style.transform = `scale(${textFactor})`;
    this.#refreshObjectHighlights(targetScale);
  }

  #schedulePinchRender(): void {
    if (this.#pinchZoomTimer !== undefined) clearTimeout(this.#pinchZoomTimer);
    this.#pinchZoomTimer = setTimeout(() => this.#commitPinchZoom(), PINCH_RENDER_DELAY);
  }

  #commitPinchZoom(): void {
    this.#pinchZoomTimer = undefined;
    this.#pinchVerticalAnchor = undefined;
    if (this.#document === undefined) {
      this.#clearPinchPreview();
      return;
    }
    if (this.#pageMode === "continuous") {
      this.#queueContinuousLayout();
    } else {
      const rect = this.#element<HTMLElement>(".stage").getBoundingClientRect();
      this.#pinchSettleRect = {
        left: rect.left,
        width: rect.width,
      };
      this.#queueSingleLayout();
    }
  }

  #cancelPinchZoom(): void {
    if (this.#pinchZoomTimer !== undefined) clearTimeout(this.#pinchZoomTimer);
    this.#pinchZoomTimer = undefined;
    this.#zoomAnchor = undefined;
    this.#pinchSettleRect = undefined;
    this.#pinchSettleAnimation?.cancel();
    this.#pinchSettleAnimation = undefined;
    this.#pinchVerticalAnchor = undefined;
    this.#clearPinchPreview();
  }

  #clearPinchPreview(): void {
    this.#pinchPreviewBaseScale = 0;
    this.#pinchPreviewScale = 0;
    const stage = this.#root.querySelector<HTMLElement>(".stage");
    stage?.style.removeProperty("transform");
    stage?.style.removeProperty("transform-origin");
    stage?.style.removeProperty("will-change");
    const textLayer = this.#root.querySelector<HTMLElement>(".stage .text-layer");
    textLayer?.style.removeProperty("transform");
    textLayer?.style.removeProperty("transform-origin");
    for (const record of this.#continuousPages.values()) {
      record.textLayer.style.removeProperty("transform");
      record.textLayer.style.removeProperty("transform-origin");
    }
  }

  #resetWheelGesture(): void {
    if (this.#wheelResetTimer !== undefined) clearTimeout(this.#wheelResetTimer);
    this.#wheelResetTimer = undefined;
    this.#wheelDelta = 0;
    this.#wheelDirection = 0;
    this.#wheelPageLocked = false;
  }

  #sheetOverrides(unitIndex: number): SheetOverrides {
    let overrides = this.#sheetSizes.get(unitIndex);
    if (overrides === undefined) {
      overrides = { rows: new Map(), columns: new Map() };
      this.#sheetSizes.set(unitIndex, overrides);
    }
    return overrides;
  }

  #sheetLayout(unit: Extract<UnitDescriptor, { type: "sheet" }>): {
    readonly rows: SheetAxisLayout;
    readonly columns: SheetAxisLayout;
  } {
    const overrides = this.#sheetOverrides(unit.index);
    return {
      rows: new SheetAxisLayout(unit.rowAxis, unit.rows, overrides.rows),
      columns: new SheetAxisLayout(unit.columnAxis, unit.columns, overrides.columns),
    };
  }

  #sheetSizeRequest(unit: Extract<UnitDescriptor, { type: "sheet" }>): {
    readonly rows: readonly { readonly index: number; readonly size: number }[];
    readonly columns: readonly { readonly index: number; readonly size: number }[];
  } {
    const overrides = this.#sheetOverrides(unit.index);
    return {
      rows: [...overrides.rows].map(([index, size]) => ({ index, size })),
      columns: [...overrides.columns].map(([index, size]) => ({ index, size })),
    };
  }

  #sheetVisibleViewport(unit: Extract<UnitDescriptor, { type: "sheet" }>, scale: number): Rect {
    const workspace = this.#workspace();
    const layout = this.#sheetLayout(unit);
    const x = Math.min(layout.columns.total, Math.max(0, workspace.scrollLeft / scale));
    const y = Math.min(layout.rows.total, Math.max(0, workspace.scrollTop / scale));
    return {
      x,
      y,
      width: Math.min(
        Math.max(1, layout.columns.total - x),
        Math.max(1, (workspace.clientWidth - SHEET_ROW_HEADER_WIDTH) / scale),
      ),
      height: Math.min(
        Math.max(1, layout.rows.total - y),
        Math.max(1, (workspace.clientHeight - SHEET_COLUMN_HEADER_HEIGHT) / scale),
      ),
    };
  }

  #sheetRenderViewport(unit: Extract<UnitDescriptor, { type: "sheet" }>, scale: number): Rect {
    const visible = this.#sheetVisibleViewport(unit, scale);
    const layout = this.#sheetLayout(unit);
    const width = Math.min(layout.columns.total, visible.width * (1 + SHEET_VIEWPORT_OVERSCAN_RATIO));
    const height = Math.min(layout.rows.total, visible.height * (1 + SHEET_VIEWPORT_OVERSCAN_RATIO));
    const pixels = scale * Math.min(2, Math.max(1, devicePixelRatio || 1));
    const snap = (value: number): number => Math.floor(value * pixels) / pixels;
    const left = Math.max(0, Math.min(layout.columns.total - width, visible.x - (width - visible.width) / 2));
    const top = Math.max(0, Math.min(layout.rows.total - height, visible.y - (height - visible.height) / 2));
    const x = snap(left), y = snap(top);
    return {
      x, y,
      width: Math.max(1, Math.ceil((left + width - x) * pixels) / pixels),
      height: Math.max(1, Math.ceil((top + height - y) * pixels) / pixels),
    };
  }

  #placeSheet(unit: Extract<UnitDescriptor, { type: "sheet" }>, viewport: Rect, scale: number): void {
    const layout = this.#sheetLayout(unit);
    const wrap = this.#element<HTMLElement>(".stage-wrap");
    const stage = this.#element<HTMLElement>(".stage");
    wrap.style.width = `${Math.max(this.#workspace().clientWidth, layout.columns.total * scale + SHEET_ROW_HEADER_WIDTH)}px`;
    wrap.style.height = `${Math.max(this.#workspace().clientHeight, layout.rows.total * scale + SHEET_COLUMN_HEADER_HEIGHT)}px`;
    stage.style.left = `${SHEET_ROW_HEADER_WIDTH + viewport.x * scale}px`;
    stage.style.top = `${SHEET_COLUMN_HEADER_HEIGHT + viewport.y * scale}px`;
    this.#renderSheetHeaders(unit, viewport, scale);
    this.#renderSheetPageBreaks(unit, layout, scale);
  }

  #resetSheetSurface(): void {
    this.#clearSheetTiles();
    const wrap = this.#element<HTMLElement>(".stage-wrap");
    const stage = this.#element<HTMLElement>(".stage");
    wrap.style.removeProperty("width");
    wrap.style.removeProperty("height");
    stage.style.removeProperty("left");
    stage.style.removeProperty("top");
    this.#element<HTMLElement>(".sheet-column-headers").replaceChildren();
    this.#element<HTMLElement>(".sheet-row-headers").replaceChildren();
    this.#element<HTMLElement>(".sheet-page-breaks").replaceChildren();
  }

  #renderSheetHeaders(unit: Extract<UnitDescriptor, { type: "sheet" }>, viewport: Rect, scale: number): void {
    const layout = this.#sheetLayout(unit);
    const columns = document.createDocumentFragment();
    const startColumn = layout.columns.indexAt(viewport.x);
    const endColumn = layout.columns.indexAt(viewport.x + viewport.width);
    for (let index = startColumn; index >= 0 && index <= endColumn; index += 1) {
      const size = layout.columns.size(index);
      if (size <= 0) continue;
      columns.append(this.#sheetHeaderItem("column", index, layout.columns.offset(index) * scale,
        size * scale, columnLabel(index)));
    }
    const rows = document.createDocumentFragment();
    const startRow = layout.rows.indexAt(viewport.y);
    const endRow = layout.rows.indexAt(viewport.y + viewport.height);
    for (let index = startRow; index >= 0 && index <= endRow; index += 1) {
      const size = layout.rows.size(index);
      if (size <= 0) continue;
      rows.append(this.#sheetHeaderItem("row", index, layout.rows.offset(index) * scale,
        size * scale, String(index + 1)));
    }
    const columnHeaders = this.#element<HTMLElement>(".sheet-column-headers");
    const rowHeaders = this.#element<HTMLElement>(".sheet-row-headers");
    columnHeaders.replaceChildren(columns);
    rowHeaders.replaceChildren(rows);
    columnHeaders.style.width = `${layout.columns.total * scale}px`;
    rowHeaders.style.height = `${layout.rows.total * scale}px`;
  }

  #renderSheetPageBreaks(
    unit: Extract<UnitDescriptor, { type: "sheet" }>,
    layout: { readonly rows: SheetAxisLayout; readonly columns: SheetAxisLayout },
    scale: number,
  ): void {
    const target = this.#element<HTMLElement>(".sheet-page-breaks");
    target.replaceChildren();
    target.hidden = unit.printSettings?.viewMode === undefined;
    if (target.hidden) return;
    target.style.left = `${SHEET_ROW_HEADER_WIDTH}px`;
    target.style.top = `${SHEET_COLUMN_HEADER_HEIGHT}px`;
    target.style.width = `${layout.columns.total * scale}px`;
    target.style.height = `${layout.rows.total * scale}px`;
    const pages = sheetPrintPages(unit, this.#sheetSizeRequest(unit));
    const columns = new Set(pages.flatMap(({ viewport }) => [viewport.x, viewport.x + viewport.width]));
    const rows = new Set(pages.flatMap(({ viewport }) => [viewport.y, viewport.y + viewport.height]));
    for (const [axis, boundaries, total] of [
      ["column", columns, layout.columns.total],
      ["row", rows, layout.rows.total],
    ] as const) {
      for (const boundary of boundaries) {
        if (boundary <= 0 || boundary >= total) continue;
        const line = globalThis.document.createElement("span");
        line.className = "sheet-page-break";
        line.dataset.axis = axis;
        if (axis === "column") line.style.left = `${boundary * scale}px`;
        else line.style.top = `${boundary * scale}px`;
        target.append(line);
      }
    }
  }

  #sheetHeaderItem(axis: "row" | "column", index: number, start: number, size: number, label: string): HTMLElement {
    const item = document.createElement("div");
    item.className = "sheet-header-item";
    if (axis === "column") {
      item.style.left = `${start}px`;
      item.style.width = `${size}px`;
      item.style.height = "100%";
    } else {
      item.style.top = `${start}px`;
      item.style.width = "100%";
      item.style.height = `${size}px`;
    }
    item.textContent = label;
    const resizer = document.createElement("span");
    resizer.className = "sheet-header-resizer";
    resizer.dataset.axis = axis;
    resizer.dataset.index = String(index);
    resizer.tabIndex = 0;
    resizer.setAttribute("role", "separator");
    resizer.setAttribute("aria-orientation", axis === "column" ? "vertical" : "horizontal");
    resizer.ariaLabel = formatMessage(axis === "column" ? this.#messages.resizeColumn : this.#messages.resizeRow, { label });
    item.append(resizer);
    return item;
  }

  #sheetFrameCoversViewport(unit: Extract<UnitDescriptor, { type: "sheet" }>, frame: Rect, visible: Rect): boolean {
    const layout = this.#sheetLayout(unit);
    const marginX = Math.max(0, (frame.width - visible.width) / 4);
    const marginY = Math.max(0, (frame.height - visible.height) / 4);
    return visible.x >= frame.x + (frame.x <= 0 ? 0 : marginX)
      && visible.y >= frame.y + (frame.y <= 0 ? 0 : marginY)
      && visible.x + visible.width <= frame.x + frame.width
        - (frame.x + frame.width >= layout.columns.total ? 0 : marginX)
      && visible.y + visible.height <= frame.y + frame.height
        - (frame.y + frame.height >= layout.rows.total ? 0 : marginY);
  }

  #queueSheetScroll(): void {
    if (this.#info?.units[this.#unitIndex]?.type !== "sheet" || this.#status !== "ready") return;
    if (this.#sheetScrollFrame !== undefined) return;
    this.#sheetScrollFrame = requestAnimationFrame(() => {
      this.#sheetScrollFrame = undefined;
      const unit = this.#info?.units[this.#unitIndex];
      const rendered = this.#rendered;
      if (unit?.type !== "sheet" || rendered?.unit.index !== unit.index || this.#status !== "ready") return;
      if (!this.#fit && Math.abs(rendered.scale - this.#zoom) >= .005) return;
      const visible = this.#sheetVisibleViewport(unit, rendered.scale);
      const previous = this.#sheetScrollPosition;
      this.#sheetScrollPosition = { x: visible.x, y: visible.y,
        dx: Math.sign(visible.x - previous.x) || previous.dx, dy: Math.sign(visible.y - previous.y) || previous.dy };
      this.#renderSheetHeaders(unit, visible, rendered.scale);
      this.#prepareSheetTiles(this.#document!, unit, {
        unitIndex: unit.index, viewport: this.#sheetRenderViewport(unit, rendered.scale),
        scale: rendered.scale, pixelRatio: Math.min(2, Math.max(1, devicePixelRatio || 1)),
        background: "#ffffff", includeTextFragments: true,
        ...this.#watermarkRequest(), sheetSizes: this.#sheetSizeRequest(unit),
      });
      if (this.#sheetScrollRendering) return;
      const frame = rendered.frame.viewport;
      if (this.#sheetFrameCoversViewport(unit, frame, visible)) return;
      this.#sheetScrollRendering = true;
      void this.#renderCurrent().finally(() => {
        this.#sheetScrollRendering = false;
        this.#queueSheetScroll();
      });
    });
  }

  #clearSheetTiles(): void {
    for (const tile of this.#sheetTiles.values()) {
      tile.controller.abort();
      if (tile.canvas !== undefined) {
        tile.canvas.remove();
        tile.canvas.width = tile.canvas.height = 0;
      }
    }
    this.#sheetTiles.clear();
    this.#sheetTileKey = "";
    this.#sheetScrollPosition = { x: 0, y: 0, dx: 0, dy: 0 };
  }

  #prepareSheetTiles(document: OfficeDocument, unit: Extract<UnitDescriptor, { type: "sheet" }>, request: RenderRequest) {
    const signature = JSON.stringify({ ...request, viewport: undefined });
    if (signature !== this.#sheetTileKey) {
      this.#clearSheetTiles();
      this.#sheetTileKey = signature;
    }
    const scale = request.scale ?? 1, ratio = request.pixelRatio ?? 1;
    const tilePixels = Math.ceil(512 * ratio);
    const step = tilePixels / (scale * ratio);
    const visible = request.viewport!;
    const layout = this.#sheetLayout(unit);
    const firstX = Math.floor(visible.x / step), firstY = Math.floor(visible.y / step);
    const lastX = Math.ceil((visible.x + visible.width) / step) - 1;
    const lastY = Math.ceil((visible.y + visible.height) / step) - 1;
    const candidates = [];
    for (let y = Math.max(0, firstY - 2); y <= Math.min(lastY + 2, Math.ceil(layout.rows.total / step) - 1); y++) {
      for (let x = Math.max(0, firstX - 2); x <= Math.min(lastX + 2, Math.ceil(layout.columns.total / step) - 1); x++) {
        candidates.push({ key: `${x}:${y}`, x, y,
          required: x >= firstX && x <= lastX && y >= firstY && y <= lastY,
          distance: Math.abs(x - (firstX + lastX) / 2) + Math.abs(y - (firstY + lastY) / 2)
            - 0.6 * ((x - (firstX + lastX) / 2) * this.#sheetScrollPosition.dx
              + (y - (firstY + lastY) / 2) * this.#sheetScrollPosition.dy),
        });
      }
    }
    candidates.sort((a, b) => Number(b.required) - Number(a.required) || a.distance - b.distance);
    // Each cached canvas is at most 512 CSS pixels square; cap backing stores at 128 MiB (64 MiB at 1x).
    const capacity = Math.min(64, Math.floor(32 * 1024 * 1024 / tilePixels ** 2));
    // Preserve viewport-anchored watermarks and the caller's render limit. Oversized
    // viewports use the existing bounded single-frame renderer, without a tile cache.
    if (request.watermark !== undefined || candidates.filter(tile => tile.required).length > capacity
      || (544 * ratio) ** 2 > this.#renderPixelLimit) {
      this.#clearSheetTiles();
      return [];
    }
    const requiredCount = candidates.filter(tile => tile.required).length;
    // Keep room for recently viewed tiles; speculative work must not evict the return path.
    const wanted = candidates.slice(0, requiredCount + Math.min(8, Math.floor((capacity - requiredCount) / 2)));
    const wantedKeys = new Set(wanted.map(tile => tile.key));
    for (const [key, tile] of this.#sheetTiles) {
      if (tile.frame === undefined && !wantedKeys.has(key)) {
        tile.controller.abort();
        this.#sheetTiles.delete(key);
      }
    }
    for (const { key, x, y, required } of wanted) {
      const cached = this.#sheetTiles.get(key);
      if (cached !== undefined) {
        cached.used ||= required;
        this.#sheetTiles.delete(key);
        this.#sheetTiles.set(key, cached);
        continue;
      }
      const viewport = { x: x * step, y: y * step,
        width: Math.min(step, layout.columns.total - x * step),
        height: Math.min(step, layout.rows.total - y * step) };
      // Paint past tile edges before cropping, so glyphs, borders and effects crossing a seam survive.
      const gutter = Math.ceil(16 * ratio) / (scale * ratio);
      const left = Math.max(0, viewport.x - gutter), top = Math.max(0, viewport.y - gutter);
      const padded = { x: left, y: top,
        width: Math.min(layout.columns.total, viewport.x + viewport.width + gutter) - left,
        height: Math.min(layout.rows.total, viewport.y + viewport.height + gutter) - top };
      const controller = new AbortController();
      const promise = document.render({ ...request, viewport: padded }, {
        priority: required ? "interactive" : "prefetch", signal: controller.signal,
        supersedeKey: `viewer-sheet:${key}`,
      }).then(frame => {
        const tile = this.#sheetTiles.get(key);
        try {
          if (controller.signal.aborted || document !== this.#document || tile?.controller !== controller) return;
          this.#paintSheetTile(tile, frame, scale, ratio);
        } finally { frame.bitmap.close(); }
      }).catch(cause => {
        if (this.#sheetTiles.get(key)?.controller === controller) this.#sheetTiles.delete(key);
        throw cause;
      });
      // Prefetch failures must not become unhandled rejections; visible consumers still receive the error.
      void promise.catch(() => undefined);
      this.#sheetTiles.set(key, { viewport, used: required, controller, promise });
    }
    const evictionOrder = [...this.#sheetTiles].sort((a, b) => Number(a[1].used) - Number(b[1].used));
    for (const [key, tile] of evictionOrder) {
      if (this.#sheetTiles.size <= capacity) break;
      if (wantedKeys.has(key)) continue;
      tile.controller.abort();
      tile.canvas?.remove();
      if (tile.canvas !== undefined) tile.canvas.width = tile.canvas.height = 0;
      this.#sheetTiles.delete(key);
    }
    return wanted.filter(tile => tile.required).map(tile => this.#sheetTiles.get(tile.key)!);
  }

  #paintSheetTile(tile: { viewport: Rect; canvas?: HTMLCanvasElement; frame?: RenderResult }, frame: RenderResult, scale: number, ratio: number): void {
    const viewport = tile.viewport;
    const canvas = globalThis.document.createElement("canvas");
    canvas.className = "sheet-tile";
    canvas.ariaHidden = "true";
    canvas.width = Math.max(1, Math.ceil(viewport.width * scale * ratio - 1e-6));
    canvas.height = Math.max(1, Math.ceil(viewport.height * scale * ratio - 1e-6));
    canvas.style.cssText = `left:${SHEET_ROW_HEADER_WIDTH + viewport.x * scale}px;top:${SHEET_COLUMN_HEADER_HEIGHT + viewport.y * scale}px;width:${viewport.width * scale}px;height:${viewport.height * scale}px`;
    const context = canvas.getContext("2d");
    if (context === null) throw new Error("Canvas 2D is unavailable");
    context.drawImage(frame.bitmap, -Math.round((viewport.x - frame.viewport.x) * scale * ratio), -Math.round((viewport.y - frame.viewport.y) * scale * ratio));
    tile.canvas = canvas;
    tile.frame = frame;
    this.#element<HTMLElement>(".stage").before(canvas);
  }

  async #renderSheetFrame(document: OfficeDocument, unit: Extract<UnitDescriptor, { type: "sheet" }>, request: RenderRequest): Promise<RenderResult> {
    const viewport = request.viewport!, pixels = (request.scale ?? 1) * (request.pixelRatio ?? 1);
    const pixelWidth = Math.max(1, Math.ceil(viewport.width * pixels));
    const pixelHeight = Math.max(1, Math.ceil(viewport.height * pixels));
    if (!Number.isSafeInteger(pixelWidth * pixelHeight) || pixelWidth * pixelHeight > this.#renderPixelLimit) {
      throw new OfficeEngineError("RENDER_PIXEL_LIMIT", "Sheet viewport exceeds the configured render pixel limit");
    }
    const signature = JSON.stringify({ ...request, viewport: undefined });
    const scale = request.scale ?? 1, ratio = request.pixelRatio ?? 1;
    const step = Math.ceil(512 * ratio) / pixels;
    const layout = this.#sheetLayout(unit);
    const x1 = Math.max(0, Math.floor(viewport.x / step) - 1), y1 = Math.max(0, Math.floor(viewport.y / step) - 1);
    const x2 = Math.min(Math.ceil(layout.columns.total / step), Math.ceil((viewport.x + viewport.width) / step) + 1);
    const y2 = Math.min(Math.ceil(layout.rows.total / step), Math.ceil((viewport.y + viewport.height) / step) + 1);
    const aligned = { x: x1 * step, y: y1 * step,
      width: Math.min(layout.columns.total, x2 * step) - x1 * step,
      height: Math.min(layout.rows.total, y2 * step) - y1 * step };
    if (signature !== this.#sheetTileKey && request.watermark === undefined
      && Math.ceil(aligned.width * pixels) * Math.ceil(aligned.height * pixels) <= Math.min(this.#renderPixelLimit, 16 * 1024 * 1024)) {
      this.#clearSheetTiles();
      this.#sheetTileKey = signature;
      const first = await document.render({ ...request, viewport: aligned }, { priority: "interactive", supersedeKey: "viewer-surface" });
      try {
        if (document !== this.#document || signature !== this.#sheetTileKey) {
          throw new OfficeEngineError("OPERATION_ABORTED", "Sheet view changed during rendering");
        }
        for (let y = y1; y < y2; y++) for (let x = x1; x < x2; x++) {
          const tile = { viewport: { x: x * step, y: y * step,
            width: Math.min(step, layout.columns.total - x * step), height: Math.min(step, layout.rows.total - y * step) },
            used: x * step < viewport.x + viewport.width && (x + 1) * step > viewport.x
              && y * step < viewport.y + viewport.height && (y + 1) * step > viewport.y, controller: new AbortController(), promise: Promise.resolve() };
          this.#paintSheetTile(tile, first, scale, ratio);
          this.#sheetTiles.set(`${x}:${y}`, tile);
        }
      } finally { first.bitmap.close(); }
    }
    const tiles = this.#prepareSheetTiles(document, unit, request);
    if (tiles.length === 0) return document.render(request, { priority: "interactive", supersedeKey: "viewer-surface" });
    await Promise.all(tiles.map(tile => tile.promise));
    if (tiles.some(tile => tile.controller.signal.aborted || tile.frame === undefined || tile.canvas === undefined)) {
      throw new OfficeEngineError("OPERATION_ABORTED", "Sheet viewport was superseded");
    }
    const canvas = globalThis.document.createElement("canvas");
    canvas.width = pixelWidth;
    canvas.height = pixelHeight;
    const context = canvas.getContext("2d");
    if (context === null) throw new Error("Canvas 2D is unavailable");
    const fragments = new Map<string, RenderedTextFragment>();
    const diagnostics = new Map<string, Diagnostic>();
    const media = new Map<string, RenderResult["media"][number]>();
    const frames = new Set(tiles.map(tile => tile.frame!));
    for (const tile of tiles) {
      context.drawImage(tile.canvas!, Math.round((tile.viewport.x - viewport.x) * pixels), Math.round((tile.viewport.y - viewport.y) * pixels));
    }
    for (const frame of frames) {
      for (const fragment of frame.textFragments ?? []) fragments.set(JSON.stringify(fragment), fragment);
      for (const diagnostic of frame.diagnostics) diagnostics.set(JSON.stringify(diagnostic), diagnostic);
      for (const item of frame.media) media.set(item.objectId, item);
    }
    const bitmap = await createImageBitmap(canvas);
    canvas.width = canvas.height = 0;
    return { bitmap, viewport, pixelWidth: bitmap.width, pixelHeight: bitmap.height,
      renderedObjectCount: [...frames].reduce((sum, frame) => sum + frame.renderedObjectCount, 0),
      textFragments: [...fragments.values()], diagnostics: [...diagnostics.values()], media: [...media.values()] };
  }

  #setNavigationWidth(width: number): void {
    const value = Math.min(MAX_NAVIGATION_WIDTH, Math.max(MIN_NAVIGATION_WIDTH, Math.round(width)));
    this.#element<HTMLElement>(".shell").style.setProperty("--dv-navigation-width", `${value}px`);
    this.#element<HTMLElement>(".navigation-resizer").setAttribute("aria-valuenow", String(value));
  }

  #beginNavigationResize(event: PointerEvent): void {
    const target = event.target instanceof Element ? event.target.closest<HTMLElement>(".navigation-resizer") : null;
    const shell = this.#element<HTMLElement>(".shell");
    if (target === null || this.clientWidth <= 720 || shell.dataset.navigation !== "true") return;
    event.preventDefault();
    target.setPointerCapture(event.pointerId);
    target.dataset.resizing = "true";
    this.#navigationResize = {
      pointerId: event.pointerId,
      start: event.clientX,
      startWidth: this.#element<HTMLElement>(".navigation").getBoundingClientRect().width,
    };
  }

  #updateNavigationResize(event: PointerEvent): void {
    const resize = this.#navigationResize;
    if (resize === undefined || resize.pointerId !== event.pointerId) return;
    event.preventDefault();
    const rtl = this.#element<HTMLElement>(".shell").dir === "rtl";
    const delta = (event.clientX - resize.start) * (rtl ? -1 : 1);
    this.#setNavigationWidth(resize.startWidth + delta);
  }

  #finishNavigationResize(event?: PointerEvent): void {
    const resize = this.#navigationResize;
    if (resize === undefined || (event !== undefined && resize.pointerId !== event.pointerId)) return;
    this.#navigationResize = undefined;
    this.#element<HTMLElement>(".navigation-resizer").dataset.resizing = "false";
  }

  #beginSheetResize(event: PointerEvent): void {
    const target = event.target instanceof Element ? event.target.closest<HTMLElement>(".sheet-header-resizer") : null;
    const unit = this.#info?.units[this.#unitIndex];
    const rendered = this.#rendered;
    const axis = target?.dataset.axis;
    const index = Number(target?.dataset.index);
    if (target === null || unit?.type !== "sheet" || rendered?.unit.index !== unit.index
      || (axis !== "row" && axis !== "column") || !Number.isInteger(index)) return;
    event.preventDefault();
    target.setPointerCapture(event.pointerId);
    this.#clearNativeSelection();
    this.#element<HTMLElement>(".shell").dataset.sheetResizing = "true";
    const layout = this.#sheetLayout(unit);
    const startSize = (axis === "column" ? layout.columns : layout.rows).size(index);
    this.#sheetResize = {
      axis,
      index,
      pointerId: event.pointerId,
      start: axis === "column" ? event.clientX : event.clientY,
      startSize,
      size: startSize,
      scale: rendered.scale,
    };
    const guide = this.#element<HTMLElement>(".sheet-resize-guide");
    guide.dataset.axis = axis;
    guide.dataset.visible = "true";
    this.#updateSheetResize(event);
  }

  #updateSheetResize(event: PointerEvent): void {
    const resize = this.#sheetResize;
    if (resize === undefined || resize.pointerId !== event.pointerId) return;
    event.preventDefault();
    this.#clearNativeSelection();
    const coordinate = resize.axis === "column" ? event.clientX : event.clientY;
    resize.size = Math.max(resize.axis === "column" ? MIN_COLUMN_WIDTH : MIN_ROW_HEIGHT,
      resize.startSize + (coordinate - resize.start) / resize.scale);
    const workspace = this.#workspace();
    const bounds = workspace.getBoundingClientRect();
    const guide = this.#element<HTMLElement>(".sheet-resize-guide");
    if (resize.axis === "column") {
      guide.style.width = "1px";
      guide.style.left = `${workspace.scrollLeft + event.clientX - bounds.left}px`;
      guide.style.top = `${workspace.scrollTop}px`;
      guide.style.height = `${workspace.clientHeight}px`;
    } else {
      guide.style.height = "1px";
      guide.style.left = `${workspace.scrollLeft}px`;
      guide.style.top = `${workspace.scrollTop + event.clientY - bounds.top}px`;
      guide.style.width = `${workspace.clientWidth}px`;
    }
  }

  #finishSheetResize(event?: PointerEvent): void {
    const resize = this.#sheetResize;
    if (resize === undefined || (event !== undefined && resize.pointerId !== event.pointerId)) return;
    this.#sheetResize = undefined;
    this.#clearNativeSelection();
    delete this.#element<HTMLElement>(".shell").dataset.sheetResizing;
    this.#element<HTMLElement>(".sheet-resize-guide").dataset.visible = "false";
    const unit = this.#info?.units[this.#unitIndex];
    if (unit?.type !== "sheet") return;
    const overrides = this.#sheetOverrides(unit.index);
    (resize.axis === "column" ? overrides.columns : overrides.rows).set(resize.index, resize.size);
    void this.#renderCurrent();
  }

  #clearNativeSelection(): void {
    this.#textSelectionAnchor = undefined;
    this.#pointerTextSelection = undefined;
    this.#pointerSelectionRectangles = [];
    globalThis.document.getSelection()?.removeAllRanges();
    (this.#root as ShadowRoot & { getSelection?: () => Selection | null }).getSelection?.()?.removeAllRanges();
    this.#queueTextSelectionPaint();
  }

  #textPositionAt(event: PointerEvent): Readonly<{ fragment: HTMLElement; offset: number }> | undefined {
    const pathTarget = event.composedPath().find((candidate): candidate is HTMLElement =>
      candidate instanceof HTMLElement && candidate.classList.contains("text-layer-fragment"));
    const target = pathTarget
      ?? this.#root.elementFromPoint(event.clientX, event.clientY)?.closest<HTMLElement>(".text-layer-fragment");
    const text = target?.textContent;
    if (target === undefined || target === null || text === undefined || text.length === 0) return undefined;
    const bounds = target.getBoundingClientRect();
    const transform = new DOMMatrixReadOnly(getComputedStyle(target).transform);
    const horizontal = Math.abs(transform.a) >= Math.abs(transform.b);
    const rawRatio = horizontal
      ? (event.clientX - bounds.left) / Math.max(1, bounds.width)
      : (event.clientY - bounds.top) / Math.max(1, bounds.height);
    const ratio = Math.min(1, Math.max(0, getComputedStyle(target).direction === "rtl" ? 1 - rawRatio : rawRatio));
    const offset = [...text].length === 1
      ? (ratio < 0.5 ? 0 : text.length)
      : Math.round(text.length * ratio);
    return { fragment: target, offset };
  }

  #beginTextSelection(event: PointerEvent): void {
    if (event.button !== 0 || this.#features.interactionMode !== "text") return;
    this.#textSelectionAnchor = this.#textPositionAt(event);
    this.#pointerTextSelection = undefined;
    this.#pointerSelectionRectangles = [];
    this.#queueTextSelectionPaint();
  }

  #updateTextSelection(event: PointerEvent): void {
    const anchor = this.#textSelectionAnchor;
    if (anchor === undefined || (event.buttons & 1) === 0) return;
    const focus = this.#textPositionAt(event);
    if (focus === undefined) return;
    this.#pointerTextSelection = { anchor, focus };
    this.#pointerSelectionRectangles = this.#textRectanglesBetween(anchor, focus);
    this.#queueTextSelectionPaint();
  }

  #finishTextSelection(event: PointerEvent): void {
    const anchor = this.#textSelectionAnchor;
    if (anchor === undefined) return;
    const focus = this.#textPositionAt(event);
    if (focus !== undefined && this.#pointerTextSelection !== undefined) {
      this.#pointerTextSelection = { anchor, focus };
      this.#pointerSelectionRectangles = this.#textRectanglesBetween(anchor, focus);
      const selection = (this.#root as ShadowRoot & { getSelection?: () => Selection | null }).getSelection?.()
        ?? globalThis.document.getSelection();
      const anchorText = anchor.fragment.firstChild;
      const focusText = focus.fragment.firstChild;
      // Native selection can stop before the final absolutely positioned glyph in WebKit.
      // Keep native caret measurement for longer runs with unequal character advances.
      if (anchorText instanceof Text && focusText instanceof Text
        && [...anchorText.data].length === 1 && [...focusText.data].length === 1) {
        selection?.setBaseAndExtent(anchorText, anchor.offset, focusText, focus.offset);
      }
    }
    this.#textSelectionAnchor = undefined;
    this.#queueTextSelectionPaint();
  }

  #textPartsBetween(
    anchor: Readonly<{ fragment: HTMLElement; offset: number }>,
    focus: Readonly<{ fragment: HTMLElement; offset: number }>,
  ): Array<Readonly<{ fragment: HTMLElement; text: Text; start: number; end: number }>> {
    const fragments = Array.from(this.#root.querySelectorAll<HTMLElement>(".text-layer-fragment"));
    let startIndex = fragments.indexOf(anchor.fragment);
    let endIndex = fragments.indexOf(focus.fragment);
    let startOffset = anchor.offset;
    let endOffset = focus.offset;
    if (startIndex > endIndex || (startIndex === endIndex && startOffset > endOffset)) {
      [startIndex, endIndex] = [endIndex, startIndex];
      [startOffset, endOffset] = [endOffset, startOffset];
    }
    if (startIndex < 0 || endIndex < startIndex) return [];
    const parts = [];
    for (let index = startIndex; index <= endIndex; index += 1) {
      const fragment = fragments[index]!;
      const text = fragment.firstChild;
      if (!(text instanceof Text)) continue;
      const start = index === startIndex ? startOffset : 0;
      const end = index === endIndex ? endOffset : text.data.length;
      if (end > start) parts.push({ fragment, text, start, end });
    }
    return parts;
  }

  #textRectanglesBetween(
    anchor: Readonly<{ fragment: HTMLElement; offset: number }>,
    focus: Readonly<{ fragment: HTMLElement; offset: number }>,
  ): DOMRect[] {
    const rectangles: DOMRect[] = [];
    for (const { fragment, text, start, end } of this.#textPartsBetween(anchor, focus)) {
      const rectangle = this.#textFragmentRectangle(fragment, text, start, end);
      if (rectangle !== undefined) rectangles.push(rectangle);
    }
    return rectangles;
  }

  #queueTextSelectionPaint(): void {
    if (this.#textSelectionFrame !== undefined) return;
    this.#textSelectionFrame = requestAnimationFrame(() => {
      this.#textSelectionFrame = undefined;
      this.#paintTextSelection();
    });
  }

  #paintTextSelection(): void {
    this.#root.querySelectorAll(".text-selection-highlight").forEach((highlight) => highlight.remove());
    if (this.#features.interactionMode !== "text") return;
    const selection = (this.#root as ShadowRoot & { getSelection?: () => Selection | null }).getSelection?.()
      ?? globalThis.document.getSelection();
    const exactRectangles = selection === null || selection.isCollapsed || selection.rangeCount === 0
      ? []
      : this.#selectedTextRectangles(selection);
    const rectangles = exactRectangles.length > 0 ? exactRectangles : this.#pointerSelectionRectangles;
    if (rectangles.length === 0) return;
    for (const layer of Array.from(this.#root.querySelectorAll<HTMLElement>(".text-layer:not([hidden])"))) {
      const layerBounds = layer.getBoundingClientRect();
      const lines: Array<{ top: number; bottom: number; rectangles: DOMRect[] }> = [];
      for (const rectangle of rectangles) {
        if (rectangle.right <= layerBounds.left || rectangle.left >= layerBounds.right
          || rectangle.bottom <= layerBounds.top || rectangle.top >= layerBounds.bottom) continue;
        const middle = (rectangle.top + rectangle.bottom) / 2;
        const line = lines.find((candidate) => middle >= candidate.top && middle <= candidate.bottom);
        if (line === undefined) {
          lines.push({ top: rectangle.top, bottom: rectangle.bottom, rectangles: [rectangle] });
        } else {
          line.top = Math.min(line.top, rectangle.top);
          line.bottom = Math.max(line.bottom, rectangle.bottom);
          line.rectangles.push(rectangle);
        }
      }
      const scaleX = layerBounds.width / Math.max(1, layer.clientWidth);
      const scaleY = layerBounds.height / Math.max(1, layer.clientHeight);
      const highlights = globalThis.document.createDocumentFragment();
      for (const line of lines) {
        const ordered = [...line.rectangles].sort((left, right) => left.left - right.left);
        const bands: Array<{ left: number; right: number }> = [];
        for (const rectangle of ordered) {
          const band = bands.at(-1);
          if (band === undefined || rectangle.left - band.right > line.bottom - line.top) {
            bands.push({ left: rectangle.left, right: rectangle.right });
          } else {
            band.right = Math.max(band.right, rectangle.right);
          }
        }
        for (const band of bands) {
          const highlight = globalThis.document.createElement("span");
          highlight.className = "text-selection-highlight";
          highlight.ariaHidden = "true";
          highlight.style.left = `${(Math.max(layerBounds.left, band.left) - layerBounds.left) / scaleX}px`;
          highlight.style.top = `${(Math.max(layerBounds.top, line.top) - layerBounds.top) / scaleY}px`;
          highlight.style.width = `${Math.max(0, Math.min(layerBounds.right, band.right)
            - Math.max(layerBounds.left, band.left)) / scaleX}px`;
          highlight.style.height = `${Math.max(0, Math.min(layerBounds.bottom, line.bottom)
            - Math.max(layerBounds.top, line.top)) / scaleY}px`;
          highlights.append(highlight);
        }
      }
      layer.append(highlights);
    }
  }

  #selectedTextParts(selection: Selection): Array<Readonly<{
    fragment: HTMLElement;
    text: Text;
    start: number;
    end: number;
  }>> | undefined {
    if (selection.rangeCount === 0) return [];
    const range = selection.getRangeAt(0);
    const containingFragment = (node: Node): HTMLElement | undefined => {
      const element = node instanceof Element ? node : node.parentElement;
      return element?.closest<HTMLElement>(".text-layer-fragment") ?? undefined;
    };
    // Firefox may retain element boundaries for selectNodeContents. Resolve their
    // text endpoints so block boxes do not inflate the glyph selection highlight.
    const boundary = (node: Node, offset: number, end: boolean): { node: Node; offset: number } => {
      if (node instanceof Text) return { node, offset };
      let child: Node | undefined | null = node.childNodes[end ? offset - 1 : offset];
      while (child !== undefined && child !== null && !(child instanceof Text)) {
        child = end ? child.lastChild : child.firstChild;
      }
      return child instanceof Text ? { node: child, offset: end ? child.length : 0 } : { node, offset };
    };
    const start = boundary(range.startContainer, range.startOffset, false);
    const end = boundary(range.endContainer, range.endOffset, true);
    const startFragment = containingFragment(start.node);
    const endFragment = containingFragment(end.node);
    if (startFragment === undefined || endFragment === undefined) return undefined;
    const startText = startFragment.firstChild;
    const endText = endFragment.firstChild;
    if (!(startText instanceof Text) || !(endText instanceof Text)) return undefined;
    return this.#textPartsBetween(
      { fragment: startFragment, offset: start.node === startText ? start.offset : 0 },
      { fragment: endFragment, offset: end.node === endText ? end.offset : endText.data.length },
    );
  }

  #selectedTextContent(selection: Selection): string | undefined {
    const pointer = this.#pointerTextSelection;
    const nativeParts = this.#selectedTextParts(selection);
    const parts = nativeParts !== undefined && nativeParts.length > 0
      ? nativeParts
      : pointer === undefined ? undefined : this.#textPartsBetween(pointer.anchor, pointer.focus);
    if (parts === undefined || parts.length === 0) return undefined;
    let value = "";
    let previousItem: Element | null = null;
    let previousFragment: HTMLElement | undefined;
    for (const { fragment, text, start, end } of parts) {
      const item = fragment.closest(".text-layer-item");
      if (previousItem !== null && item !== previousItem && previousFragment !== undefined) {
        const previousBounds = previousFragment.getBoundingClientRect();
        const bounds = fragment.getBoundingClientRect();
        const overlap = Math.min(previousBounds.bottom, bounds.bottom)
          - Math.max(previousBounds.top, bounds.top);
        const sameLine = overlap >= Math.min(previousBounds.height, bounds.height) / 2;
        if (!value.endsWith("\n") && !value.endsWith(" ")) {
          if (!sameLine) value += "\n";
          else if (bounds.left - previousBounds.right
            > Math.min(previousBounds.height, bounds.height) * 0.08) value += " ";
        }
      }
      value += text.data.slice(start, end);
      previousItem = item;
      previousFragment = fragment;
    }
    return value;
  }

  #selectedTextRectangles(selection: Selection): DOMRect[] {
    const parts = this.#selectedTextParts(selection);
    if (parts === undefined) {
      const range = selection.getRangeAt(0);
      return Array.from(range.getClientRects()).filter(({ width, height }) => width > 0 && height > 0);
    }
    const rectangles: DOMRect[] = [];
    for (const { fragment, text, start, end } of parts) {
      const rectangle = this.#textFragmentRectangle(fragment, text, start, end);
      if (rectangle !== undefined) rectangles.push(rectangle);
    }
    return rectangles;
  }

  #textFragmentRectangle(fragment: HTMLElement, text: Text, start: number, end: number): DOMRect | undefined {
    const safeStart = Math.min(text.data.length, Math.max(0, start));
    const safeEnd = Math.min(text.data.length, Math.max(safeStart, end));
    if (safeEnd <= safeStart) return undefined;
    const fragmentRange = globalThis.document.createRange();
    fragmentRange.setStart(text, safeStart);
    fragmentRange.setEnd(text, safeEnd);
    const rectangle = fragmentRange.getBoundingClientRect();
    if (rectangle.width > 0 && rectangle.height > 0) return rectangle;
    const bounds = fragment.getBoundingClientRect();
    const direction = getComputedStyle(fragment).direction;
    const startRatio = safeStart / text.data.length;
    const endRatio = safeEnd / text.data.length;
    const left = direction === "rtl"
      ? bounds.right - bounds.width * endRatio
      : bounds.left + bounds.width * startRatio;
    return new DOMRect(left, bounds.top, bounds.width * (endRatio - startRatio), bounds.height);
  }

  async #autoFitSheetSize(event: Event): Promise<void> {
    const target = event.target instanceof Element ? event.target.closest<HTMLElement>(".sheet-header-resizer") : null;
    const unit = this.#info?.units[this.#unitIndex];
    const axis = target?.dataset.axis;
    const index = Number(target?.dataset.index);
    if (target === null || unit?.type !== "sheet" || (axis !== "row" && axis !== "column") || !Number.isInteger(index)) return;
    event.preventDefault();
    const overrides = this.#sheetOverrides(unit.index);
    const document = this.#document;
    if (document === undefined) return;
    if (axis === "column") {
      const sourceColumns = new SheetAxisLayout(unit.columnAxis, unit.columns, new Map());
      const x = sourceColumns.offset(index);
      const sourceWidth = sourceColumns.size(index);
      try {
        const objects = await document.listObjects({
          unitIndex: unit.index,
          types: ["cell"],
          textOnly: true,
          viewport: { x, y: 0, width: sourceWidth, height: unit.height },
          limit: 20_000,
        });
        if (document !== this.#document || unit !== this.#info?.units[this.#unitIndex]) return;
        const context = globalThis.document.createElement("canvas").getContext("2d");
        if (context === null) return;
        const cells = objects
          .filter((object) => object.text !== undefined && object.bounds.width <= sourceWidth + 0.01)
          .map((object) => ({
            text: object.text!,
            width: sourceWidth,
            wrap: object.wrapText ?? false,
            ...(object.fontRuns === undefined ? {} : { fontRuns: object.fontRuns }),
          }));
        if (cells.length === 0) overrides.columns.delete(index);
        else overrides.columns.set(index, sheetAutoFitColumnWidth(context, cells, MIN_COLUMN_WIDTH));
      } catch {
        if (document !== this.#document) return;
        overrides.columns.delete(index);
      }
      void this.#renderCurrent();
      return;
    }
    const sourceRows = new SheetAxisLayout(unit.rowAxis, unit.rows, new Map());
    const y = sourceRows.offset(index);
    const sourceHeight = sourceRows.size(index);
    try {
      const objects = await document.listObjects({
        unitIndex: unit.index,
        types: ["cell"],
        textOnly: true,
        viewport: { x: 0, y, width: unit.width, height: sourceHeight },
        limit: 20_000,
      });
      if (document !== this.#document || unit !== this.#info?.units[this.#unitIndex]) return;
      const context = globalThis.document.createElement("canvas").getContext("2d");
      if (context === null) return;
      const columns = this.#sheetLayout(unit).columns;
      const cells = objects
        .filter((object) => object.text !== undefined && object.bounds.height <= sourceHeight + 0.01)
        .map((object) => ({
          text: object.text!,
          width: columns.map(object.bounds.x + object.bounds.width) - columns.map(object.bounds.x),
          wrap: object.wrapText ?? false,
          ...(object.fontRuns === undefined ? {} : { fontRuns: object.fontRuns }),
        }));
      overrides.rows.set(index, sheetAutoFitRowHeight(context, cells, sourceHeight));
    } catch {
      if (document !== this.#document) return;
      overrides.rows.delete(index);
    }
    void this.#renderCurrent();
  }

  #watermarkRequest(): Pick<RenderRequest, "watermark"> {
    const watermark = this.#config.watermark;
    return watermark !== undefined ? { watermark } : {};
  }

  async #renderCurrent(transitionDirection = 0, propagateError = false): Promise<void> {
    if (this.#pagePrefetchTimer !== undefined) clearTimeout(this.#pagePrefetchTimer);
    this.#pagePrefetchTimer = undefined;
    const document = this.#document;
    const unit = this.#info?.units[this.#unitIndex];
    if (document === undefined || unit === undefined || this.#pageMode !== "single") return;
    const revision = ++this.#renderRevision;
    const scale = this.#fit ? this.#fitScale(unit) : this.#zoom;
    const pixelRatio = Math.min(2, Math.max(1, devicePixelRatio || 1));
    let viewport: Rect | undefined;
    if (unit.type === "sheet") viewport = this.#sheetRenderViewport(unit, scale);
    if (unit.type !== "sheet" && this.#rendered?.unit.index !== unit.index
      && this.#showThumbnailPreview(unit.index, this.#element<HTMLCanvasElement>(".surface"))) {
      this.#cancelPageAnimation();
      this.#closeRenderedFrame();
      this.#hideSelection();
      this.#clearObjectHover();
      const stage = this.#element<HTMLElement>(".stage");
      stage.style.removeProperty("transform");
      stage.style.removeProperty("transform-origin");
      stage.style.removeProperty("will-change");
      const canvas = this.#element<HTMLCanvasElement>(".surface");
      canvas.style.width = stage.style.width = `${unit.width * scale}px`;
      canvas.style.height = stage.style.height = `${unit.height * scale}px`;
      this.#element<HTMLElement>(".stage-wrap").hidden = false;
      this.#zoom = scale;
      this.#updateControls();
    }
    try {
      const request: RenderRequest = {
        unitIndex: unit.index,
        ...(viewport === undefined ? {} : { viewport }),
        scale,
        pixelRatio,
        background: "#ffffff",
        includeTextFragments: true,
        ...this.#watermarkRequest(),
        ...(unit.type === "sheet" ? { sheetSizes: this.#sheetSizeRequest(unit) } : {}),
      };
      const frame = unit.type === "sheet"
        ? await this.#renderSheetFrame(document, unit, request)
        : await document.render(request, { priority: "interactive", supersedeKey: "viewer-surface" });
      if (revision !== this.#renderRevision || document !== this.#document
        || (!this.#fit && Math.abs(scale - this.#zoom) >= .005)) {
        frame.bitmap.close();
        return;
      }
      const stage = this.#element<HTMLElement>(".stage");
      this.#cancelPageAnimation();
      stage.style.removeProperty("transform");
      stage.style.removeProperty("transform-origin");
      stage.style.removeProperty("will-change");
      this.#closeRenderedFrame();
      const currentCanvas = this.#element<HTMLCanvasElement>(".surface");
      const transition = transitionDirection !== 0
        && currentCanvas.dataset.preview !== "true"
        && this.#info?.format !== "pdf"
        && !matchMedia("(prefers-reduced-motion: reduce)").matches
        && currentCanvas.width > 1
        && currentCanvas.height > 1;
      const canvas = transition ? globalThis.document.createElement("canvas") : currentCanvas;
      if (transition) {
        canvas.className = "surface";
        canvas.ariaLabel = currentCanvas.ariaLabel;
      }
      const context = canvas.getContext("2d");
      if (context === null) throw new Error("Canvas 2D is unavailable");
      canvas.width = frame.pixelWidth;
      canvas.height = frame.pixelHeight;
      const cssWidth = frame.viewport.width * scale;
      const cssHeight = frame.viewport.height * scale;
      canvas.style.width = `${cssWidth}px`;
      canvas.style.height = `${cssHeight}px`;
      stage.style.width = `${cssWidth}px`;
      stage.style.height = `${cssHeight}px`;
      if (unit.type === "sheet") this.#placeSheet(unit, frame.viewport, scale);
      else this.#resetSheetSurface();
      context.drawImage(frame.bitmap, 0, 0);
      delete canvas.dataset.preview;
      this.#drawThumbnail(unit.index, canvas, frame.viewport);
      this.#applySingleZoomAnchor(frame, scale);
      this.#clearPinchPreview();
      this.#animatePinchSettle(stage);
      if (transition) {
        currentCanvas.className = "page-transition-outgoing";
        currentCanvas.ariaHidden = "true";
        currentCanvas.before(canvas);
        this.#animatePageTransition(canvas, currentCanvas, transitionDirection);
      }
      this.#rendered = { frame, unit, scale };
      await this.#renderTextLayer(
        document,
        unit,
        frame.viewport,
        scale,
        frame.textFragments,
        this.#element<HTMLElement>(".text-layer"),
        () => revision === this.#renderRevision && document === this.#document && this.#pageMode === "single",
      );
      this.#zoom = scale;
      this.#refreshObjectHighlights();
      this.#renderDiagnostics();
      this.#setStatus("ready");
      this.#setMessage(this.#info?.format.toUpperCase() ?? "");
      this.#updateControls();
      this.#emit("docviewkit-diagnostic", { diagnostics: [...document.diagnostics(), ...frame.diagnostics] });
      this.#emitState();
      this.#prefetchedPages.set(unit.index, document);
      this.#pagePrefetchTimer = setTimeout(() => {
        this.#pagePrefetchTimer = undefined;
        this.#prefetchAdjacentPages(document, unit.index);
      }, 80);
    } catch (cause) {
      if (revision !== this.#renderRevision || document !== this.#document) return;
      if (cause instanceof OfficeEngineError && cause.code === "OPERATION_ABORTED") return;
      if (propagateError) throw cause;
      if (cause instanceof OfficeEngineError && cause.code === "RENDER_PIXEL_LIMIT"
        && this.#rendered !== undefined) {
        this.#fit = false;
        this.#zoom = this.#rendered.scale;
        this.#zoomAnchor = undefined;
        const stage = this.#element<HTMLElement>(".stage");
        stage.style.removeProperty("transform");
        stage.style.removeProperty("transform-origin");
        stage.style.removeProperty("will-change");
        this.#updateControls();
        this.#emitState();
        return;
      }
      this.#setStatus("error", errorMessage(cause));
      this.#emit("docviewkit-error", { cause });
    }
  }

  #prefetchAdjacentPages(document: OfficeDocument, unitIndex: number): void {
    if (this.#info?.format !== "pdf" || this.#pageMode !== "single" || document !== this.#document) return;
    for (const adjacentIndex of [unitIndex + 1, unitIndex - 1]) {
      const unit = this.#info.units[adjacentIndex];
      if (unit?.type !== "page" || this.#prefetchedPages.get(adjacentIndex) === document) continue;
      this.#prefetchedPages.set(adjacentIndex, document);
      const scale = Math.min(.22, 120 / Math.max(1, unit.width), 96 / Math.max(1, unit.height));
      void document.render({
        unitIndex: adjacentIndex,
        scale,
        pixelRatio: 1,
        background: "#ffffff",
        ...this.#watermarkRequest(),
      }, {
        priority: "prefetch",
        supersedeKey: `viewer-page-prefetch:${adjacentIndex}`,
      }).then((frame) => frame.bitmap.close(), () => {
        if (this.#prefetchedPages.get(adjacentIndex) === document) {
          this.#prefetchedPages.delete(adjacentIndex);
        }
      });
    }
  }

  #queueSingleLayout(): void {
    if (this.#pinchZoomTimer !== undefined) return;
    this.#singleLayoutPending = true;
    if (this.#singleLayoutFrame !== undefined) return;
    this.#singleLayoutFrame = requestAnimationFrame(() => {
      this.#singleLayoutFrame = undefined;
      if (this.#pageMode !== "single") return;
      this.#previewSingleLayout();
      if (this.#singleLayoutRendering) return;
      this.#singleLayoutPending = false;
      this.#singleLayoutRendering = true;
      void this.#renderCurrent().finally(() => {
        this.#singleLayoutRendering = false;
        if (this.#singleLayoutPending) this.#queueSingleLayout();
      });
    });
  }

  #previewSingleLayout(): void {
    const rendered = this.#rendered;
    if (!this.#fit || rendered === undefined) return;
    const stage = this.#element<HTMLElement>(".stage");
    this.#cancelPageAnimation();
    stage.style.removeProperty("transform");
    stage.style.removeProperty("transform-origin");
    const targetScale = this.#fitScale(rendered.unit);
    const factor = targetScale / rendered.scale;
    if (!Number.isFinite(factor) || Math.abs(factor - 1) < .005) {
      stage.style.removeProperty("will-change");
      return;
    }
    this.#previewSinglePageScale(targetScale);
  }

  #cancelLayoutFrames(): void {
    this.#cancelPinchZoom();
    if (this.#singleLayoutFrame !== undefined) cancelAnimationFrame(this.#singleLayoutFrame);
    if (this.#continuousLayoutFrame !== undefined) cancelAnimationFrame(this.#continuousLayoutFrame);
    if (this.#sheetScrollFrame !== undefined) cancelAnimationFrame(this.#sheetScrollFrame);
    this.#singleLayoutFrame = undefined;
    this.#continuousLayoutFrame = undefined;
    this.#sheetScrollFrame = undefined;
    this.#singleLayoutPending = false;
    this.#zoomAnchor = undefined;
    this.#cancelPageAnimation();
    const stage = this.#root.querySelector<HTMLElement>(".stage");
    stage?.style.removeProperty("transform");
    stage?.style.removeProperty("transform-origin");
    stage?.style.removeProperty("will-change");
  }

  #animatePageTransition(incoming: HTMLCanvasElement, outgoing: HTMLCanvasElement, direction: number): void {
    const viewerStyle = getComputedStyle(this);
    const durationToken = viewerStyle.getPropertyValue("--dv-motion-state").trim();
    const parsedDuration = Number.parseFloat(durationToken);
    const duration = durationToken.endsWith("s") && !durationToken.endsWith("ms")
      ? parsedDuration * 1000
      : parsedDuration;
    const easing = viewerStyle.getPropertyValue("--dv-ease-out").trim() || "cubic-bezier(.22, 1, .36, 1)";
    const stateDuration = Number.isFinite(duration) ? Math.min(200, duration) : 190;
    incoming.style.willChange = "transform, opacity";
    outgoing.style.willChange = "transform, opacity";
    const incomingAnimation = incoming.animate([
      { transform: `translateY(${direction * 6}px)`, opacity: .55 },
      { transform: "none", opacity: 1 },
    ], { duration: stateDuration, easing });
    const outgoingAnimation = outgoing.animate([
      { transform: "none", opacity: 1 },
      { transform: `translateY(${-direction * 4}px)`, opacity: 0 },
    ], { duration: Math.min(150, stateDuration), easing: "cubic-bezier(.7, 0, .84, 0)" });
    const animations = [incomingAnimation, outgoingAnimation];
    this.#pageAnimations = animations;
    this.#outgoingPage = outgoing;
    let remaining = animations.length;
    const settle = (): void => {
      remaining -= 1;
      if (remaining !== 0 || this.#pageAnimations !== animations) return;
      this.#pageAnimations = [];
    };
    incomingAnimation.onfinish = (): void => {
      incoming.style.removeProperty("will-change");
      settle();
    };
    incomingAnimation.oncancel = settle;
    outgoingAnimation.onfinish = (): void => {
      outgoing.remove();
      if (this.#outgoingPage === outgoing) this.#outgoingPage = undefined;
      settle();
    };
    outgoingAnimation.oncancel = settle;
  }

  #cancelPageAnimation(): void {
    const animations = this.#pageAnimations;
    this.#pageAnimations = [];
    for (const animation of animations) animation.cancel();
    this.#outgoingPage?.remove();
    this.#outgoingPage = undefined;
    this.#root.querySelector<HTMLElement>(".surface")?.style.removeProperty("will-change");
  }

  #fitScale(unit: UnitDescriptor): number {
    const workspace = this.#workspace();
    const horizontalPadding = workspace.clientWidth <= 720 ? 24 : 64;
    const verticalPadding = workspace.clientWidth <= 720 ? 36 : 64;
    const sheetLayout = unit.type === "sheet" ? this.#sheetLayout(unit) : undefined;
    const width = sheetLayout?.columns.total ?? unit.width;
    const height = sheetLayout?.rows.total ?? unit.height;
    const widthScale = Math.max(.1, (workspace.clientWidth - horizontalPadding) / width);
    const heightScale = Math.max(.1, (workspace.clientHeight - verticalPadding) / height);
    const scale = unit.type === "sheet" ? widthScale : Math.min(widthScale, heightScale, 1.5);
    return Math.min(MAX_ZOOM, Math.max(.1, scale));
  }

  #supportsContinuous(): boolean {
    const units = this.#info?.units ?? [];
    return units.length > 1 && units.every((unit) => unit.type !== "sheet");
  }

  #continuousScale(): number {
    if (!this.#fit) return this.#zoom;
    const workspace = this.#workspace();
    const horizontalPadding = workspace.clientWidth <= 720 ? 24 : 64;
    const width = (this.#info?.units ?? []).reduce((width, unit) => Math.max(width, unit.width), 1);
    return Math.min(1.5, Math.max(.1, (workspace.clientWidth - horizontalPadding) / width));
  }

  async #setPageMode(mode: ViewerPageMode): Promise<void> {
    if (this.#document === undefined) return;
    const next = mode === "continuous" && this.#supportsContinuous() ? "continuous" : "single";
    if (next === this.#pageMode) return;
    this.#cancelPinchZoom();
    this.#zoomAnchor = undefined;
    this.#pageMode = next;
    this.#hideSelection();
    this.#updatePageModeControls();
    if (next === "continuous") {
      await this.#buildContinuousPages();
      this.#emitState();
    } else {
      this.#clearContinuousPages();
      this.#workspace().scrollTo({ top: 0, left: 0 });
      await this.#renderCurrent();
    }
  }

  #updatePageModeControls(): void {
    const continuous = this.#pageMode === "continuous";
    const supported = this.#supportsContinuous();
    const disabled = this.#document === undefined || !supported || this.#status === "loading";
    this.#root.querySelectorAll<HTMLButtonElement>('[data-action="page-mode"]').forEach((button) => {
      button.hidden = !this.#features.pageMode || !supported;
      button.disabled = disabled;
      button.setAttribute("aria-pressed", String(continuous));
    });
  }

  #previewContinuousNavigationLayout(): void {
    const units = this.#info?.units ?? [];
    const workspace = this.#workspace();
    this.#element<HTMLElement>(".continuous-view").style.width = `${workspace.clientWidth}px`;
    const scale = this.#continuousScale();
    for (let unitIndex = this.#unitIndex - 1; unitIndex <= this.#unitIndex + 1; unitIndex += 1) {
      const unit = units[unitIndex];
      const record = this.#continuousPages.get(unitIndex);
      if (unit === undefined || record === undefined) continue;
      record.scale = scale;
      record.element.style.width = `${unit.width * scale}px`;
      record.element.style.height = `${unit.height * scale}px`;
      const textFactor = record.renderedScale > 0 ? scale / record.renderedScale : 1;
      record.textLayer.style.transformOrigin = "top left";
      record.textLayer.style.transform = `scale(${textFactor})`;
    }
    this.#refreshObjectHighlights();
    this.#zoom = this.#continuousPages.get(this.#unitIndex)?.scale ?? this.#zoom;
  }

  #queueContinuousLayout(preserveAnchor = false): void {
    if (this.#pinchZoomTimer !== undefined) return;
    if (this.#continuousLayoutFrame !== undefined) return;
    this.#continuousLayoutFrame = requestAnimationFrame(() => {
      this.#continuousLayoutFrame = undefined;
      if (this.#pageMode !== "continuous") return;
      const anchorElement = preserveAnchor ? this.#continuousPages.get(this.#unitIndex)?.element : undefined;
      const anchorRect = anchorElement?.getBoundingClientRect();
      this.#clearPinchPreview();
      const units = this.#info?.units ?? [];
      const scale = this.#continuousScale();
      for (const unit of units) {
        const record = this.#continuousPages.get(unit.index);
        if (record === undefined) continue;
        record.scale = scale;
        record.element.style.width = `${unit.width * scale}px`;
        record.element.style.height = `${unit.height * scale}px`;
        if (record.visible || Math.abs(unit.index - this.#unitIndex) <= 1) {
          void this.#renderContinuousPage(unit.index);
        }
      }
      this.#element<HTMLElement>(".continuous-view").style.removeProperty("width");
      if (anchorElement !== undefined && anchorRect !== undefined) {
        const settledRect = anchorElement.getBoundingClientRect();
        const workspace = this.#workspace();
        workspace.scrollLeft += settledRect.left - anchorRect.left;
        workspace.scrollTop += settledRect.top - anchorRect.top;
      }
      this.#refreshObjectHighlights();
      this.#zoom = this.#continuousPages.get(this.#unitIndex)?.scale ?? this.#zoom;
      this.#updateControls();
    });
  }

  async #buildContinuousPages(): Promise<void> {
    const officeDocument = this.#document;
    const units = this.#info?.units ?? [];
    if (officeDocument === undefined || this.#pageMode !== "continuous" || !this.#supportsContinuous()) return;
    this.#clearContinuousPages();
    const revision = this.#continuousRevision;
    this.#renderRevision += 1;
    this.#closeRenderedFrame();
    this.#element<HTMLElement>(".stage-wrap").hidden = true;
    const view = this.#element<HTMLElement>(".continuous-view");
    view.hidden = false;

    const fragment = globalThis.document.createDocumentFragment();
    const scale = this.#continuousScale();
    for (const unit of units) {
      const page = globalThis.document.createElement("section");
      page.className = "continuous-page";
      page.dataset.unitIndex = String(unit.index);
      page.dataset.rendered = "false";
      page.style.width = `${unit.width * scale}px`;
      page.style.height = `${unit.height * scale}px`;
      page.ariaLabel = unit.name || formatMessage(this.#messages.pagePosition, { current: unit.index + 1, total: units.length });
      const canvas = globalThis.document.createElement("canvas");
      canvas.width = 1;
      canvas.height = 1;
      canvas.ariaLabel = this.#messages.documentCanvas;
      const textLayer = globalThis.document.createElement("div");
      textLayer.className = "text-layer";
      textLayer.hidden = this.#features.interactionMode !== "text";
      const hover = globalThis.document.createElement("div");
      hover.className = "object-hover";
      hover.hidden = true;
      const selection = globalThis.document.createElement("div");
      selection.className = "continuous-selection";
      selection.hidden = true;
      page.append(canvas, textLayer, hover, selection);
      this.#continuousPages.set(unit.index, {
        element: page,
        canvas,
        textLayer,
        hover,
        selection,
        scale,
        renderedScale: 0,
        rendered: false,
        rendering: false,
        visible: false,
        pixels: 0,
        lastUsed: 0,
        diagnostics: [],
        textFragments: undefined,
      });
      fragment.append(page);
    }
    view.append(fragment);
    this.#refreshObjectHighlights();

    if (typeof IntersectionObserver === "function") {
      this.#continuousObserver = new IntersectionObserver((entries) => {
        if (revision !== this.#continuousRevision || this.#pageMode !== "continuous") return;
        for (const entry of entries) {
          const unitIndex = Number((entry.target as HTMLElement).dataset.unitIndex);
          const record = this.#continuousPages.get(unitIndex);
          if (record === undefined) continue;
          record.visible = entry.isIntersecting;
          record.element.style.contentVisibility = entry.isIntersecting ? "visible" : "auto";
          if (entry.isIntersecting) {
            record.lastUsed = performance.now();
            void this.#renderContinuousPage(unitIndex, revision);
          }
        }
        this.#evictContinuousPages();
      }, { root: this.#workspace(), rootMargin: "400px 0px", threshold: .01 });

      this.#continuousNavigationObserver = new IntersectionObserver((entries) => {
        if (revision !== this.#continuousRevision || this.#pageMode !== "continuous") return;
        for (const entry of entries) {
          const unitIndex = Number((entry.target as HTMLElement).dataset.unitIndex);
          if (entry.isIntersecting) this.#continuousCenterPages.add(unitIndex);
          else this.#continuousCenterPages.delete(unitIndex);
        }
        const workspaceCenter = this.#workspace().getBoundingClientRect().top + this.#workspace().clientHeight / 2;
        const closest = [...this.#continuousCenterPages]
          .map((unitIndex) => this.#continuousPages.get(unitIndex))
          .filter((record): record is ContinuousPageRecord => record !== undefined)
          .sort((left, right) => Math.abs(left.element.getBoundingClientRect().top - workspaceCenter)
            - Math.abs(right.element.getBoundingClientRect().top - workspaceCenter))[0];
        if (closest !== undefined) this.#syncContinuousUnit(Number(closest.element.dataset.unitIndex));
      }, { root: this.#workspace(), rootMargin: "-45% 0px -45% 0px", threshold: .01 });

      for (const record of this.#continuousPages.values()) {
        this.#continuousObserver.observe(record.element);
        this.#continuousNavigationObserver.observe(record.element);
      }
    }

    this.#zoom = scale;
    this.#updateControls();
    await Promise.all([
      this.#renderContinuousPage(this.#unitIndex, revision),
      this.#renderContinuousPage(this.#unitIndex + 1, revision),
    ]);
    if (revision !== this.#continuousRevision || this.#pageMode !== "continuous") return;
    await this.#revealContinuousUnit(this.#unitIndex, false);
    this.#zoom = this.#continuousPages.get(this.#unitIndex)?.scale ?? this.#zoom;
    this.#setStatus("ready");
    this.#setMessage(this.#info?.format.toUpperCase() ?? "");
    this.#renderDiagnostics();
    this.#updateControls();
  }

  async #renderContinuousPage(unitIndex: number, revision = this.#continuousRevision): Promise<void> {
    const document = this.#document;
    const record = this.#continuousPages.get(unitIndex);
    const unit = this.#info?.units[unitIndex];
    if (document === undefined || unit === undefined || record === undefined
      || this.#pinchZoomTimer !== undefined || record.rendering
      || (record.rendered && Math.abs(record.renderedScale - record.scale) < .005)
      || revision !== this.#continuousRevision || this.#pageMode !== "continuous") return;
    record.rendering = true;
    if (!record.rendered) this.#showThumbnailPreview(unitIndex, record.canvas);
    const scale = record.scale;
    let completed = false;
    let frame: RenderResult | undefined;
    try {
      frame = await document.render({
        unitIndex,
        scale,
        pixelRatio: Math.min(2, Math.max(1, devicePixelRatio || 1)),
        background: "#ffffff",
        includeTextFragments: true,
        ...this.#watermarkRequest(),
      }, { priority: "visible", supersedeKey: `viewer-continuous:${unitIndex}` });
      if (revision !== this.#continuousRevision || this.#pageMode !== "continuous" || document !== this.#document) return;
      record.canvas.width = frame.pixelWidth;
      record.canvas.height = frame.pixelHeight;
      const context = record.canvas.getContext("2d");
      if (context === null) throw new Error("Canvas 2D is unavailable");
      context.drawImage(frame.bitmap, 0, 0);
      delete record.canvas.dataset.preview;
      this.#drawThumbnail(unitIndex, record.canvas, frame.viewport);
      await this.#renderTextLayer(
        document,
        unit,
        frame.viewport,
        scale,
        frame.textFragments,
        record.textLayer,
        () => revision === this.#continuousRevision && document === this.#document && this.#pageMode === "continuous",
      );
      this.#continuousPixels -= record.pixels;
      record.pixels = frame.pixelWidth * frame.pixelHeight;
      this.#continuousPixels += record.pixels;
      record.diagnostics = frame.diagnostics;
      record.textFragments = frame.textFragments;
      record.renderedScale = scale;
      record.rendered = true;
      completed = true;
      record.lastUsed = performance.now();
      record.element.dataset.rendered = "true";
      this.#evictContinuousPages();
      this.#renderDiagnostics();
      this.#emit("docviewkit-diagnostic", { diagnostics: this.#allDiagnostics() });
    } catch (cause) {
      if (revision === this.#continuousRevision && this.#pageMode === "continuous") {
        record.element.dataset.error = "true";
        this.#emit("docviewkit-error", { cause });
      }
    } finally {
      record.rendering = false;
      frame?.bitmap.close();
      if (completed && Math.abs(record.renderedScale - record.scale) >= .005
        && (record.visible || Math.abs(unitIndex - this.#unitIndex) <= 1)) {
        void this.#renderContinuousPage(unitIndex, revision);
      }
    }
  }

  async #revealContinuousUnit(unitIndex: number, animate = true): Promise<void> {
    const record = this.#continuousPages.get(unitIndex);
    if (record === undefined) return;
    await this.#renderContinuousPage(unitIndex);
    const workspace = this.#workspace();
    const top = Math.max(0, record.element.offsetTop - (workspace.clientWidth <= 720 ? 18 : 28));
    const near = Math.abs(top - workspace.scrollTop) <= workspace.clientHeight * CONTINUOUS_SMOOTH_SCROLL_VIEWPORTS;
    workspace.scrollTo({ top, left: 0, behavior: animate && near ? "smooth" : "auto" });
  }

  #syncContinuousUnit(unitIndex: number): void {
    if (this.#pinchZoomTimer !== undefined) return;
    if (unitIndex === this.#unitIndex || !this.#continuousPages.has(unitIndex)) return;
    this.#unitIndex = unitIndex;
    this.#zoom = this.#continuousPages.get(unitIndex)?.scale ?? this.#zoom;
    this.#updateNavigationSelection();
    this.#emitState();
  }

  #evictContinuousPages(): void {
    if (this.#continuousPixels <= CONTINUOUS_CACHE_PIXELS) return;
    const candidates = [...this.#continuousPages]
      .filter(([unitIndex, record]) => unitIndex !== this.#unitIndex
        && !record.visible && !record.rendering && record.rendered)
      .map(([, record]) => record)
      .sort((left, right) => left.lastUsed - right.lastUsed);
    for (const record of candidates) {
      this.#continuousPixels -= record.pixels;
      record.pixels = 0;
      record.rendered = false;
      record.renderedScale = 0;
      record.diagnostics = [];
      record.textFragments = undefined;
      record.canvas.width = 1;
      record.canvas.height = 1;
      record.textLayer.replaceChildren();
      record.element.dataset.rendered = "false";
      if (this.#continuousPixels <= CONTINUOUS_CACHE_PIXELS) break;
    }
  }

  #clearContinuousPages(): void {
    if (this.#continuousLayoutFrame !== undefined) cancelAnimationFrame(this.#continuousLayoutFrame);
    this.#continuousLayoutFrame = undefined;
    this.#continuousRevision += 1;
    this.#continuousObserver?.disconnect();
    this.#continuousNavigationObserver?.disconnect();
    this.#continuousObserver = undefined;
    this.#continuousNavigationObserver = undefined;
    this.#continuousCenterPages.clear();
    for (const record of this.#continuousPages.values()) {
      record.canvas.width = 1;
      record.canvas.height = 1;
    }
    this.#continuousPages.clear();
    this.#continuousPixels = 0;
    this.#root.querySelector(".continuous-view")?.replaceChildren();
    const view = this.#root.querySelector<HTMLElement>(".continuous-view");
    if (view !== null) {
      view.style.removeProperty("width");
      view.hidden = true;
    }
  }

  #allDiagnostics(): readonly Diagnostic[] {
    return [
      ...(this.#document?.diagnostics() ?? []),
      ...(this.#pageMode === "continuous"
        ? [...this.#continuousPages.values()].flatMap((record) => record.diagnostics)
        : (this.#rendered?.frame.diagnostics ?? [])),
    ];
  }

  #canRenderAtZoom(scale: number): boolean {
    const info = this.#info;
    if (info === undefined) return false;
    const pixelRatio = Math.min(2, Math.max(1, devicePixelRatio || 1));
    const firstUnit = this.#pageMode === "continuous" ? 0 : this.#unitIndex;
    const endUnit = this.#pageMode === "continuous" ? info.units.length : firstUnit + 1;
    for (let unitIndex = firstUnit; unitIndex < endUnit; unitIndex += 1) {
      const unit = info.units[unitIndex];
      if (unit === undefined) return false;
      let width = unit.width;
      let height = unit.height;
      if (unit.type === "sheet") {
        const workspace = this.#workspace();
        const layout = this.#sheetLayout(unit);
        width = Math.min(layout.columns.total,
          Math.max(1, (workspace.clientWidth - SHEET_ROW_HEADER_WIDTH) / scale));
        height = Math.min(layout.rows.total,
          Math.max(1, (workspace.clientHeight - SHEET_COLUMN_HEADER_HEIGHT) / scale));
      }
      const pixels = Math.ceil(width * scale * pixelRatio) * Math.ceil(height * scale * pixelRatio);
      if (!Number.isSafeInteger(pixels) || pixels > this.#renderPixelLimit) return false;
    }
    return true;
  }

  #captureZoomAnchor(
    targetScale: number,
    focalPoint?: Readonly<{ clientX: number; clientY: number }>,
  ): ZoomAnchor | undefined {
    const workspace = this.#workspace();
    const workspaceRect = workspace.getBoundingClientRect();
    const preferredX = focalPoint?.clientX ?? workspaceRect.left + workspaceRect.width / 2;
    if (this.#pageMode === "continuous") {
      const anchoredUnit = this.#pinchVerticalAnchor?.unitIndex ?? this.#unitIndex;
      const record = this.#continuousPages.get(anchoredUnit);
      if (record === undefined || record.scale <= 0) return undefined;
      const rect = record.element.getBoundingClientRect();
      const clientX = Math.min(rect.right, Math.max(rect.left, preferredX));
      return {
        unitIndex: anchoredUnit,
        documentX: (clientX - rect.left) / record.scale,
        ...(this.#pinchVerticalAnchor !== undefined
          ? {
              documentY: this.#pinchVerticalAnchor.documentY,
              clientY: this.#pinchVerticalAnchor.clientY,
            }
          : {}),
        clientX,
        targetScale,
      };
    }
    const rendered = this.#rendered;
    if (rendered === undefined || rendered.unit.index !== this.#unitIndex || rendered.scale <= 0) return undefined;
    const stage = this.#element<HTMLElement>(".stage");
    stage.style.removeProperty("transform");
    stage.style.removeProperty("transform-origin");
    stage.style.removeProperty("will-change");
    const rect = stage.getBoundingClientRect();
    const displayedScale = this.#pinchPreviewScale > 0 ? this.#pinchPreviewScale : rendered.scale;
    const clientX = Math.min(rect.right, Math.max(rect.left, preferredX));
    return {
      unitIndex: rendered.unit.index,
      documentX: rendered.frame.viewport.x + (clientX - rect.left) / displayedScale,
      ...(this.#pinchVerticalAnchor !== undefined
        ? {
            documentY: this.#pinchVerticalAnchor.documentY,
            clientY: this.#pinchVerticalAnchor.clientY,
          }
        : {}),
      clientX,
      targetScale,
    };
  }

  #capturePinchVerticalAnchor(event: WheelEvent): PinchVerticalAnchor | undefined {
    if (this.#pageMode === "continuous") {
      let unitIndex = this.#unitIndex;
      for (const [candidateIndex, record] of this.#continuousPages) {
        const rect = record.element.getBoundingClientRect();
        if (event.clientY >= rect.top && event.clientY <= rect.bottom) {
          unitIndex = candidateIndex;
          break;
        }
      }
      const record = this.#continuousPages.get(unitIndex);
      if (record === undefined || record.scale <= 0) return undefined;
      const rect = record.element.getBoundingClientRect();
      const clientY = Math.min(rect.bottom, Math.max(rect.top, event.clientY));
      return {
        unitIndex,
        documentY: (clientY - rect.top) / record.scale,
        clientY,
      };
    }
    const rendered = this.#rendered;
    if (rendered === undefined || rendered.scale <= 0) return undefined;
    const rect = this.#element<HTMLElement>(".stage").getBoundingClientRect();
    const displayedScale = this.#pinchPreviewScale > 0 ? this.#pinchPreviewScale : rendered.scale;
    const clientY = Math.min(rect.bottom, Math.max(rect.top, event.clientY));
    return {
      unitIndex: rendered.unit.index,
      documentY: rendered.frame.viewport.y + (clientY - rect.top) / displayedScale,
      clientY,
    };
  }

  #applySingleZoomAnchor(frame: RenderResult, scale: number): void {
    const anchor = this.#zoomAnchor;
    if (anchor === undefined || anchor.unitIndex !== this.#unitIndex
      || Math.abs(anchor.targetScale - scale) >= .005) return;
    this.#zoomAnchor = undefined;
    const workspace = this.#workspace();
    const stageRect = this.#element<HTMLElement>(".stage").getBoundingClientRect();
    if (workspace.scrollWidth <= workspace.clientWidth) {
      workspace.scrollLeft = 0;
    } else {
      workspace.scrollLeft += stageRect.left + (anchor.documentX - frame.viewport.x) * scale - anchor.clientX;
    }
    if (anchor.documentY !== undefined && anchor.clientY !== undefined) {
      workspace.scrollTop += stageRect.top + (anchor.documentY - frame.viewport.y) * scale - anchor.clientY;
    }
  }

  #animatePinchSettle(stage: HTMLElement): void {
    const previous = this.#pinchSettleRect;
    this.#pinchSettleRect = undefined;
    this.#pinchSettleAnimation?.cancel();
    this.#pinchSettleAnimation = undefined;
    if (previous === undefined || matchMedia("(prefers-reduced-motion: reduce)").matches) return;
    const current = stage.getBoundingClientRect();
    if (current.width <= 0) return;
    const translateX = previous.left - current.left;
    if (Math.abs(translateX) < .5) return;
    const animation = stage.animate([
      { transform: `translateX(${translateX}px)` },
      { transform: "none" },
    ], {
      duration: 160,
      easing: "cubic-bezier(.22, 1, .36, 1)",
    });
    this.#pinchSettleAnimation = animation;
    const clear = (): void => {
      if (this.#pinchSettleAnimation === animation) this.#pinchSettleAnimation = undefined;
    };
    animation.onfinish = clear;
    animation.oncancel = clear;
  }

  #applyContinuousZoomAnchor(): void {
    const anchor = this.#zoomAnchor;
    if (anchor === undefined) return;
    const record = this.#continuousPages.get(anchor.unitIndex);
    if (record === undefined || Math.abs(anchor.targetScale - record.scale) >= .005) return;
    this.#zoomAnchor = undefined;
    const workspace = this.#workspace();
    const rect = record.element.getBoundingClientRect();
    if (anchor.documentY !== undefined && anchor.clientY !== undefined) {
      workspace.scrollTop += rect.top + anchor.documentY * record.scale - anchor.clientY;
    }
    if (workspace.scrollWidth <= workspace.clientWidth) {
      workspace.scrollLeft = 0;
      return;
    }
    workspace.scrollLeft += rect.left + anchor.documentX * record.scale - anchor.clientX;
  }

  #setZoom(
    value: number,
    snap = true,
    focalPoint?: Readonly<{ clientX: number; clientY: number }>,
    deferRender = false,
  ): number | undefined {
    if (this.#document === undefined) return undefined;
    if (!deferRender && this.#pinchZoomTimer !== undefined) this.#cancelPinchZoom();
    const zoom = Math.min(MAX_ZOOM, Math.max(MIN_ZOOM,
      snap ? Math.round(value * 4) / 4 : Math.round(value * 100) / 100));
    if (!this.#fit && Math.abs(zoom - this.#zoom) < .005) return undefined;
    if (zoom > this.#zoom && !this.#canRenderAtZoom(zoom)) return undefined;
    this.#clearNativeSelection();
    this.#zoomAnchor = deferRender && this.#pinchZoomTimer !== undefined && this.#zoomAnchor !== undefined
      ? { ...this.#zoomAnchor, targetScale: zoom }
      : this.#captureZoomAnchor(zoom, focalPoint);
    this.#fit = false;
    this.#zoom = zoom;
    if (!deferRender) {
      if (this.#pageMode === "continuous") this.#queueContinuousLayout();
      else this.#queueSingleLayout();
    }
    this.#updateControls();
    return zoom;
  }

  #commitZoomInput(): void {
    const input = this.#element<HTMLInputElement>(".zoom-input");
    if (Number.isFinite(input.valueAsNumber)) this.#setZoom(input.valueAsNumber / 100, false);
    this.#syncZoomInput();
  }

  #syncZoomInput(): void {
    this.#element<HTMLInputElement>(".zoom-input").value = String(Math.round(this.#zoom * 100));
  }

  #fitView(): void {
    if (this.#document === undefined) return;
    this.#cancelPinchZoom();
    this.#zoomAnchor = undefined;
    this.#fit = true;
    if (this.#pageMode === "continuous") void this.#buildContinuousPages();
    else void this.#renderCurrent();
  }

  async #selectUnit(unitIndex: number): Promise<void> {
    if (this.#document === undefined) return;
    if (unitIndex < 0 || unitIndex >= (this.#info?.units.length ?? 0)) return;
    if (unitIndex === this.#unitIndex) {
      if (this.#pageMode === "continuous") await this.#revealContinuousUnit(unitIndex);
      return;
    }
    const transitionDirection = Math.sign(unitIndex - this.#unitIndex);
    this.#zoomAnchor = undefined;
    this.#unitIndex = unitIndex;
    this.#updateFormulaBar();
    this.#updateNavigationSelection();
    if (this.#pageMode === "continuous") await this.#revealContinuousUnit(unitIndex);
    else {
      this.#workspace().scrollTo({ top: 0, left: 0 });
      await this.#renderCurrent(transitionDirection);
    }
  }

  #assertUnit(unitIndex: number): void {
    const unit = this.#info?.units[unitIndex];
    if (!Number.isInteger(unitIndex) || unitIndex < 0 || unit === undefined || unit.index !== unitIndex) {
      throw new RangeError(`Document has no unit at index ${unitIndex}`);
    }
  }

  #buildNavigation(): void {
    const list = this.#element<HTMLElement>(".unit-list");
    const tabs = this.#element<HTMLElement>(".sheet-tabs");
    list.replaceChildren();
    tabs.replaceChildren();
    this.#thumbnailObserver?.disconnect();
    this.#thumbnailObserver = undefined;
    const spreadsheet = this.#info?.kind === "spreadsheet";
    this.#element<HTMLElement>(".shell").dataset.spreadsheet = String(spreadsheet);
    if (spreadsheet) {
      for (const unit of this.#info?.units ?? []) {
        const tab = document.createElement("button");
        tab.type = "button";
        tab.className = "sheet-tab";
        if (unit.type === "sheet" && unit.tabColor !== undefined) {
          tab.dataset.tabColor = "true";
          tab.style.setProperty("--dv-sheet-tab-color", `#${unit.tabColor.toString(16).padStart(8, "0")}`);
        }
        tab.dataset.unitIndex = String(unit.index);
        tab.setAttribute("role", "tab");
        tab.textContent = unit.name || String(unit.index + 1);
        tab.addEventListener("click", () => void this.#selectUnit(unit.index));
        tabs.append(tab);
      }
      this.#updateNavigationSelection();
      this.#updateNavigationVisibility();
      return;
    }
    this.#navigationMode = (this.#info?.outline.length ?? 0) > 0 ? "outline" : "thumbnails";
    this.#outlineSelection = -1;
    this.#thumbnailObserver = new IntersectionObserver((entries) => {
      for (const entry of entries) {
        const preview = entry.target as HTMLElement;
        preview.dataset.visible = String(entry.isIntersecting);
        if (!entry.isIntersecting) continue;
        const button = preview.closest<HTMLElement>("[data-unit-index]");
        const unitIndex = Number(button?.dataset.unitIndex);
        if (Number.isInteger(unitIndex)) void this.#renderThumbnail(unitIndex);
      }
    }, { root: this.#element<HTMLElement>(".navigation"), rootMargin: "160px 0px" });
    this.#navigationStart = -1;
    this.#navigationEnd = -1;
    this.#renderNavigationMode();
    this.#updateNavigationSelection();
    this.#updateNavigationVisibility();
  }

  #setNavigationMode(mode: ViewerNavigationMode): void {
    if (mode === "outline" && (this.#info?.outline.length ?? 0) === 0) return;
    if (mode === this.#navigationMode) return;
    this.#navigationMode = mode;
    this.#renderNavigationMode();
    this.#updateNavigationSelection();
  }

  #renderNavigationMode(): void {
    const hasOutline = (this.#info?.outline.length ?? 0) > 0;
    const switcher = this.#element<HTMLElement>(".navigation-switcher");
    switcher.hidden = !hasOutline;
    for (const mode of ["outline", "thumbnails"] as const) {
      const button = this.#element<HTMLButtonElement>(`[data-action="navigation-${mode}"]`);
      button.setAttribute("aria-selected", String(this.#navigationMode === mode));
      button.tabIndex = this.#navigationMode === mode ? 0 : -1;
    }
    this.#element<HTMLElement>(".unit-list").dataset.navigationMode = this.#navigationMode;
    if (this.#navigationMode === "outline") this.#renderOutline();
    else this.#renderNavigationWindow(this.#unitIndex, true);
  }

  #renderOutline(): void {
    this.#thumbnailObserver?.disconnect();
    const list = this.#element<HTMLElement>(".unit-list");
    const fragment = document.createDocumentFragment();
    for (const [index, item] of (this.#info?.outline ?? []).entries()) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "outline-button";
      button.dataset.unitIndex = String(item.unitIndex);
      button.dataset.outlineIndex = String(index);
      button.setAttribute("role", "option");
      button.style.setProperty("--dv-outline-level", String(item.level));
      button.textContent = item.title;
      button.title = item.title;
      button.addEventListener("click", () => {
        this.#outlineSelection = index;
        this.#updateNavigationSelection();
        void this.#selectUnit(item.unitIndex);
      });
      fragment.append(button);
    }
    list.replaceChildren(fragment);
  }

  #renderNavigationWindow(center: number, force = false): void {
    if (this.#info?.kind === "spreadsheet" || this.#navigationMode !== "thumbnails") return;
    const units = this.#info?.units ?? [];
    const count = units.length;
    const size = Math.min(count, NAVIGATION_WINDOW_SIZE);
    const start = count <= NAVIGATION_WINDOW_SIZE
      ? 0
      : Math.min(count - size, Math.max(0, Math.floor(center) - Math.floor(size / 2)));
    const end = start + size;
    if (!force && start === this.#navigationStart && end === this.#navigationEnd) return;
    this.#navigationStart = start;
    this.#navigationEnd = end;
    this.#thumbnailObserver?.disconnect();
    const list = this.#element<HTMLElement>(".unit-list");
    const retainedButtons = new Map(
      Array.from(list.querySelectorAll<HTMLElement>(".unit-button"), (button) => [
        Number(button.dataset.unitIndex),
        button,
      ]),
    );
    const fragment = document.createDocumentFragment();
    const spacer = (items: number) => {
      const element = document.createElement("div");
      element.setAttribute("role", "presentation");
      element.style.blockSize = `${items * NAVIGATION_ITEM_HEIGHT}px`;
      return element;
    };
    if (start > 0) fragment.append(spacer(start));
    for (const unit of units.slice(start, end)) {
      const retained = retainedButtons.get(unit.index);
      if (retained !== undefined) {
        fragment.append(retained);
        continue;
      }
      const button = document.createElement("button");
      button.type = "button";
      button.className = "unit-button";
      button.dataset.unitIndex = String(unit.index);
      button.setAttribute("role", "option");
      button.setAttribute("aria-posinset", String(unit.index + 1));
      button.setAttribute("aria-setsize", String(count));
      button.ariaLabel = unit.name || String(unit.index + 1);
      const number = document.createElement("span");
      number.className = "unit-number";
      number.textContent = String(unit.index + 1);
      const preview = document.createElement("span");
      preview.className = "thumbnail";
      const aspect = unit.width / Math.max(1, unit.height);
      preview.style.setProperty("--dv-thumbnail-aspect", String(aspect));
      preview.style.setProperty("--dv-thumbnail-max-width", `${96 * aspect}px`);
      const canvas = document.createElement("canvas");
      canvas.ariaHidden = "true";
      preview.append(canvas);
      button.append(number, preview);
      button.addEventListener("click", () => void this.#selectUnit(unit.index));
      fragment.append(button);
    }
    if (end < count) fragment.append(spacer(count - end));
    list.replaceChildren(fragment);
    list.querySelectorAll<HTMLElement>(".thumbnail").forEach((preview) => this.#thumbnailObserver?.observe(preview));
    list.querySelectorAll<HTMLElement>(".unit-button").forEach((button) => {
      const unitIndex = Number(button.dataset.unitIndex);
      if (this.#rendered?.unit.index === unitIndex || this.#continuousPages.get(unitIndex)?.rendered) {
        void this.#renderThumbnail(unitIndex);
      }
      const current = unitIndex === this.#unitIndex;
      button.setAttribute("aria-current", current ? "page" : "false");
      button.setAttribute("aria-selected", String(current));
      button.tabIndex = current ? 0 : -1;
    });
  }

  async #renderThumbnail(unitIndex: number): Promise<void> {
    const document = this.#document;
    const unit = this.#info?.units[unitIndex];
    if (document === undefined || unit === undefined) return;
    const existingCanvas = this.#root.querySelector<HTMLCanvasElement>(`.unit-button[data-unit-index="${unitIndex}"] canvas`);
    if (existingCanvas === null || existingCanvas.dataset.rendered === "true") return;
    const rendered = this.#rendered;
    if (rendered?.unit.index === unitIndex
      && this.#drawThumbnail(unitIndex, this.#element<HTMLCanvasElement>(".surface"), rendered.frame.viewport)) return;
    const record = this.#continuousPages.get(unitIndex);
    if (record?.rendered && this.#drawThumbnail(unitIndex, record.canvas,
      { x: 0, y: 0, width: unit.width, height: unit.height })) return;
    if (this.#thumbnailRequests.has(unitIndex)) return;
    const controller = new AbortController();
    this.#thumbnailRequests.set(unitIndex, controller);
    const scale = Math.min(.22, 120 / Math.max(1, unit.width), 96 / Math.max(1, unit.height));
    let frame: RenderResult | undefined;
    let retry = false;
    try {
      frame = await document.render({
        unitIndex,
        scale,
        pixelRatio: 1,
        background: "#ffffff",
        ...this.#watermarkRequest(),
      }, {
        priority: "prefetch",
        supersedeKey: `viewer-thumbnail:${unitIndex}`,
        signal: controller.signal,
      });
      if (document !== this.#document) return;
      if (controller.signal.aborted) return;
      this.#drawThumbnail(unitIndex, frame.bitmap, frame.viewport);
    } catch (cause) {
      if (cause instanceof OfficeEngineError
        && cause.code === "OPERATION_ABORTED"
        && document === this.#document && !controller.signal.aborted) {
        const canvas = this.#root.querySelector<HTMLCanvasElement>(`.unit-button[data-unit-index="${unitIndex}"] canvas`);
        retry = canvas?.closest<HTMLElement>(".thumbnail")?.dataset.visible === "true";
      }
      // Thumbnail failure must not block the primary reading surface.
    } finally {
      frame?.bitmap.close();
      if (this.#thumbnailRequests.get(unitIndex) === controller) this.#thumbnailRequests.delete(unitIndex);
      if (retry && document === this.#document) void this.#renderThumbnail(unitIndex);
    }
  }

  #drawThumbnail(unitIndex: number, source: CanvasImageSource, viewport: Rect): boolean {
    const unit = this.#info?.units[unitIndex];
    const canvas = this.#root.querySelector<HTMLCanvasElement>(`.unit-button[data-unit-index="${unitIndex}"] canvas`);
    if (unit === undefined || unit.type === "sheet" || canvas === null
      || viewport.x !== 0 || viewport.y !== 0
      || Math.abs(viewport.width - unit.width) > .01 || Math.abs(viewport.height - unit.height) > .01) return false;
    if (canvas.dataset.rendered === "true") return true;
    const context = canvas.getContext("2d");
    if (context === null) return false;
    const scale = Math.min(.22, 120 / Math.max(1, unit.width), 96 / Math.max(1, unit.height));
    canvas.width = Math.max(1, Math.ceil(viewport.width * scale));
    canvas.height = Math.max(1, Math.ceil(viewport.height * scale));
    context.imageSmoothingQuality = "high";
    context.drawImage(source, 0, 0, canvas.width, canvas.height);
    const aspect = viewport.width / Math.max(1, viewport.height);
    const preview = canvas.closest<HTMLElement>(".thumbnail");
    preview?.style.setProperty("--dv-thumbnail-aspect", String(aspect));
    preview?.style.setProperty("--dv-thumbnail-max-width", `${96 * aspect}px`);
    canvas.dataset.rendered = "true";
    this.#thumbnailRequests.get(unitIndex)?.abort();
    return true;
  }

  #showThumbnailPreview(unitIndex: number, canvas: HTMLCanvasElement): boolean {
    const thumbnail = this.#root.querySelector<HTMLCanvasElement>(`.unit-button[data-unit-index="${unitIndex}"] canvas`);
    if (thumbnail?.dataset.rendered !== "true") return false;
    const context = canvas.getContext("2d");
    if (context === null) return false;
    canvas.width = thumbnail.width;
    canvas.height = thumbnail.height;
    context.drawImage(thumbnail, 0, 0);
    canvas.dataset.preview = "true";
    return true;
  }

  #toggleNavigation(force?: boolean): void {
    if (this.#info?.kind === "spreadsheet" || !this.#features.navigation || this.#config.navigation === "hidden") return;
    const shell = this.#element<HTMLElement>(".shell");
    const visible = force ?? shell.dataset.navigation !== "true";
    const revision = ++this.#navigationTransitionRevision;
    const animated = this.clientWidth > 720 && !matchMedia("(prefers-reduced-motion: reduce)").matches;
    this.#navigationTransitioning = animated;
    const finish = (): void => {
      if (revision !== this.#navigationTransitionRevision) return;
      this.#navigationTransitioning = false;
      if (!this.#fit || this.#document === undefined) return;
      if (this.#pageMode === "continuous") this.#queueContinuousLayout(true);
      else this.#queueSingleLayout();
    };
    if (animated) {
      const onTransitionEnd = (event: TransitionEvent): void => {
        if (event.propertyName !== "grid-template-columns") return;
        shell.removeEventListener("transitionend", onTransitionEnd);
        finish();
      };
      shell.addEventListener("transitionend", onTransitionEnd);
    }
    shell.dataset.navigation = String(visible);
    this.#root.querySelector<HTMLButtonElement>('[data-action="navigation"]')?.setAttribute("aria-pressed", String(visible));
    if (!animated) finish();
  }

  #updateNavigationVisibility(): void {
    const preference = this.#config.navigation ?? "auto";
    const visible = this.#info?.kind !== "spreadsheet" && this.#features.navigation
      && preference !== "hidden"
      && (preference === "visible" || (this.clientWidth > 720
        && ((this.#info?.units.length ?? 0) > 1 || (this.#info?.outline.length ?? 0) > 0)));
    this.#element<HTMLElement>(".shell").dataset.navigation = String(visible);
  }

  #updateNavigationSelection(): void {
    if (this.#info?.kind !== "spreadsheet" && this.#navigationMode === "thumbnails"
      && (this.#unitIndex < this.#navigationStart || this.#unitIndex >= this.#navigationEnd)) {
      this.#renderNavigationWindow(this.#unitIndex);
    }
    this.#root.querySelectorAll<HTMLElement>(".unit-button").forEach((button) => {
      const current = Number(button.dataset.unitIndex) === this.#unitIndex;
      button.setAttribute("aria-current", current ? "page" : "false");
      button.setAttribute("aria-selected", String(current));
      button.tabIndex = current ? 0 : -1;
      if (current) this.#revealNavigationButton(button);
    });
    const outlineButtons = Array.from(this.#root.querySelectorAll<HTMLElement>(".outline-button"));
    const retainedOutline = outlineButtons[this.#outlineSelection];
    if (Number(retainedOutline?.dataset.unitIndex) !== this.#unitIndex) {
      this.#outlineSelection = outlineButtons.findIndex(
        (button) => Number(button.dataset.unitIndex) === this.#unitIndex,
      );
    }
    outlineButtons.forEach((button, index) => {
      const current = index === this.#outlineSelection;
      button.setAttribute("aria-current", current ? "page" : "false");
      button.setAttribute("aria-selected", String(current));
      button.tabIndex = current ? 0 : -1;
      if (current) this.#revealNavigationButton(button);
    });
    this.#root.querySelectorAll<HTMLElement>(".sheet-tab").forEach((tab) => {
      const current = Number(tab.dataset.unitIndex) === this.#unitIndex;
      tab.setAttribute("aria-selected", String(current));
      tab.tabIndex = current ? 0 : -1;
      if (current) tab.scrollIntoView({ block: "nearest", inline: "nearest" });
    });
    this.#updateControls();
  }

  #stepNavigationUnit(direction: -1 | 1): void {
    if (this.#navigationMode === "outline") {
      const items = Array.from(this.#root.querySelectorAll<HTMLElement>(".outline-button"));
      const current = items.findIndex((item) => item === this.#root.activeElement);
      const next = items[Math.min(items.length - 1, Math.max(0, current + direction))];
      if (next === undefined) return;
      next.focus({ preventScroll: true });
      this.#revealNavigationButton(next);
      return;
    }
    const unitCount = this.#info?.units.length ?? 0;
    const next = Math.min(unitCount - 1, Math.max(0, this.#unitIndex + direction));
    if (next === this.#unitIndex) return;
    void this.#selectUnit(next);
    this.#root.querySelector<HTMLElement>(`.unit-button[data-unit-index="${next}"]`)
      ?.focus({ preventScroll: true });
  }

  #revealNavigationButton(button: HTMLElement): void {
    const navigation = this.#element<HTMLElement>(".navigation");
    const navigationBounds = navigation.getBoundingClientRect();
    const buttonBounds = button.getBoundingClientRect();
    if (buttonBounds.top < navigationBounds.top) {
      navigation.scrollTop += buttonBounds.top - navigationBounds.top;
    } else if (buttonBounds.bottom > navigationBounds.bottom) {
      navigation.scrollTop += buttonBounds.bottom - navigationBounds.bottom;
    }
  }

  #openSearch(): void {
    if (!this.#features.search || this.#document === undefined) return;
    const search = this.#element<HTMLElement>(".search");
    search.hidden = false;
    const action = this.#root.querySelector<HTMLButtonElement>('[data-action="search"]');
    if (action !== null) {
      action.hidden = true;
      action.setAttribute("aria-pressed", "true");
    }
    this.#element<HTMLInputElement>(".search-input").focus();
  }

  #closeSearch(): void {
    this.#element<HTMLElement>(".search").hidden = true;
    const action = this.#root.querySelector<HTMLButtonElement>('[data-action="search"]');
    if (action !== null) {
      action.hidden = !this.#features.search;
      action.setAttribute("aria-pressed", "false");
    }
    this.#clearSearch();
    action?.focus();
  }

  #clearSearch(): void {
    if (this.#searchTimer !== undefined) clearTimeout(this.#searchTimer);
    this.#searchTimer = undefined;
    this.#searchQuery = "";
    this.#searchResults = [];
    this.#searchIndex = -1;
    const input = this.#root.querySelector<HTMLInputElement>(".search-input");
    if (input !== null) input.value = "";
    this.#clearSearchHighlight();
    this.#hideSelection();
    this.#updateSearchPosition();
  }

  async #performSearch(): Promise<void> {
    const document = this.#document;
    if (document === undefined) return;
    const query = this.#element<HTMLInputElement>(".search-input").value.trim();
    this.#searchQuery = query;
    if (query.length === 0) {
      this.#searchResults = [];
      this.#searchIndex = -1;
      this.#clearSearchHighlight();
      this.#hideSelection();
      this.#updateSearchPosition();
      return;
    }
    const results = await document.searchText({ query, limit: 10_000 });
    if (document !== this.#document || query !== this.#searchQuery) return;
    this.#searchResults = results;
    this.#searchIndex = results.length === 0 ? -1 : 0;
    this.#updateSearchPosition();
    if (results.length === 0) {
      this.#setMessage(this.#messages.noResults);
      this.#clearSearchHighlight();
      this.#hideSelection();
    } else {
      await this.#focusSearchResult();
    }
    this.#emitState();
  }

  async #stepSearch(offset: number): Promise<void> {
    if (this.#searchResults.length === 0) {
      await this.#performSearch();
      return;
    }
    this.#searchIndex = (this.#searchIndex + offset + this.#searchResults.length) % this.#searchResults.length;
    this.#updateSearchPosition();
    await this.#focusSearchResult();
    this.#emitState();
  }

  async #focusSearchResult(): Promise<void> {
    const result = this.#searchResults[this.#searchIndex];
    if (result === undefined) return;
    await this.#selectUnit(result.object.unitIndex);
    this.#hideSelection();
    if (!this.#refreshSearchHighlight(true)) {
      this.#showSelection(result.object.bounds, true, result.object.id, result.object.text);
    }
  }

  async #refreshTextLayers(): Promise<void> {
    const revision = this.#interactionRevision;
    const document = this.#document;
    if (this.#features.interactionMode !== "text" || document === undefined) return;
    const rendered = this.#rendered;
    if (this.#pageMode === "single" && rendered !== undefined) {
      await this.#renderTextLayer(
        document,
        rendered.unit,
        rendered.frame.viewport,
        rendered.scale,
        rendered.frame.textFragments,
        this.#element<HTMLElement>(".text-layer"),
        () => revision === this.#interactionRevision && document === this.#document && this.#pageMode === "single",
      );
      this.#refreshSearchHighlight();
      return;
    }
    await Promise.all([...this.#continuousPages.entries()]
      .filter(([, record]) => record.rendered)
      .map(async ([unitIndex, record]) => {
        const unit = this.#info?.units[unitIndex];
        if (unit === undefined) return;
        await this.#renderTextLayer(
          document,
          unit,
          { x: 0, y: 0, width: unit.width, height: unit.height },
          record.renderedScale,
          record.textFragments,
          record.textLayer,
          () => revision === this.#interactionRevision && document === this.#document && this.#pageMode === "continuous",
        );
      }));
    this.#refreshSearchHighlight();
  }

  #textFragmentElement(
    placement: RenderedTextFragment,
    viewport: Rect,
    scale: number,
    measure: CanvasRenderingContext2D | null,
  ): HTMLSpanElement | undefined {
    if (placement.text.length === 0 || placement.height <= 0
      || ![
        placement.width,
        placement.height,
        placement.letterSpacing,
        placement.transform.a,
        placement.transform.b,
        placement.transform.c,
        placement.transform.d,
        placement.transform.e,
        placement.transform.f,
      ].every(Number.isFinite)) return undefined;
    const run = globalThis.document.createElement("span");
    run.className = "text-layer-fragment";
    run.dataset.fontRun = "";
    if (placement.start !== undefined) run.dataset.textStart = String(placement.start);
    if (placement.end !== undefined) run.dataset.textEnd = String(placement.end);
    run.dataset.textLine = String(placement.line);
    run.style.font = placement.font;
    run.style.letterSpacing = `${placement.letterSpacing}px`;
    run.style.lineHeight = "1";
    run.style.height = `${placement.height}px`;
    run.style.direction = placement.direction;
    run.style.unicodeBidi = "plaintext";
    let measuredWidth = placement.width;
    if (measure !== null) {
      measure.font = placement.font;
      measure.letterSpacing = `${placement.letterSpacing}px`;
      const width = measure.measureText(placement.text).width;
      if (Number.isFinite(width) && width > 0) measuredWidth = width;
    }
    const widthScale = measuredWidth > 0 && placement.width >= 0
      ? placement.width / measuredWidth
      : 1;
    run.style.width = `${Math.max(0, measuredWidth)}px`;
    run.style.transform = `matrix(${
      placement.transform.a * scale * widthScale
    }, ${
      placement.transform.b * scale * widthScale
    }, ${
      placement.transform.c * scale
    }, ${
      placement.transform.d * scale
    }, ${
      (placement.transform.e - viewport.x) * scale
    }, ${
      (placement.transform.f - viewport.y) * scale
    })`;
    return run;
  }

  async #renderTextLayer(
    document: OfficeDocument,
    unit: UnitDescriptor,
    viewport: Rect,
    scale: number,
    textFragments: readonly RenderedTextFragment[] | undefined,
    target: HTMLElement,
    isCurrent: () => boolean,
  ): Promise<void> {
    if (this.#features.interactionMode !== "text") {
      target.hidden = true;
      target.replaceChildren();
      return;
    }
    if (textFragments !== undefined && textFragments.length > 0) {
      const items = new Map<string, HTMLElement>();
      const fragment = globalThis.document.createDocumentFragment();
      const measure = globalThis.document.createElement("canvas").getContext("2d");
      for (const placement of textFragments.slice(0, MAX_TEXT_LAYER_OBJECTS)) {
        const run = this.#textFragmentElement(placement, viewport, scale, measure);
        if (run === undefined) continue;
        let item = items.get(placement.objectId);
        if (item === undefined) {
          item = globalThis.document.createElement("span");
          item.className = "text-layer-item";
          item.dataset.objectId = placement.objectId;
          item.dataset.layoutSource = "render";
          items.set(placement.objectId, item);
          fragment.append(item);
        }
        run.textContent = placement.text;
        item.append(run);
      }
      if (!isCurrent() || this.#features.interactionMode !== "text") return;
      target.replaceChildren(fragment);
      target.hidden = false;
      return;
    }
    const layout = unit.type === "sheet" ? this.#sheetLayout(unit) : undefined;
    const objectViewport = layout === undefined ? viewport : {
      x: layout.columns.unmap(viewport.x),
      y: layout.rows.unmap(viewport.y),
      width: layout.columns.unmap(viewport.x + viewport.width) - layout.columns.unmap(viewport.x),
      height: layout.rows.unmap(viewport.y + viewport.height) - layout.rows.unmap(viewport.y),
    };
    const objects = await document.listObjects({
      unitIndex: unit.index,
      types: ["text-box", "paragraph", "shape", "cell"],
      textOnly: true,
      viewport: objectViewport,
      limit: MAX_TEXT_LAYER_OBJECTS,
    });
    if (!isCurrent() || this.#features.interactionMode !== "text") return;
    const visible = new Map<string, { object: DocumentObject; bounds: Rect }>();
    for (const object of objects) {
      const value = object.text;
      if (value === undefined || value.length === 0) continue;
      const bounds = layout === undefined ? object.bounds : {
        x: layout.columns.map(object.bounds.x),
        y: layout.rows.map(object.bounds.y),
        width: layout.columns.map(object.bounds.x + object.bounds.width) - layout.columns.map(object.bounds.x),
        height: layout.rows.map(object.bounds.y + object.bounds.height) - layout.rows.map(object.bounds.y),
      };
      if (bounds.x >= viewport.x + viewport.width || bounds.x + bounds.width <= viewport.x
        || bounds.y >= viewport.y + viewport.height || bounds.y + bounds.height <= viewport.y) continue;
      const key = `${bounds.x}:${bounds.y}:${bounds.width}:${bounds.height}:${value}`;
      const existing = visible.get(key);
      if (existing === undefined
        || (existing.object.fontRuns?.length ?? 0) === 0 && (object.fontRuns?.length ?? 0) > 0) {
        visible.set(key, { object, bounds });
      }
    }
    const fragment = globalThis.document.createDocumentFragment();
    for (const { object, bounds } of visible.values()) {
      const item = globalThis.document.createElement("span");
      item.className = "text-layer-item";
      item.dataset.objectId = object.id;
      item.style.left = `${(bounds.x - viewport.x) * scale}px`;
      item.style.top = `${(bounds.y - viewport.y) * scale}px`;
      item.style.width = `${Math.max(1, bounds.width * scale)}px`;
      item.style.height = `${Math.max(1, bounds.height * scale)}px`;
      item.style.fontSize = `${Math.max(6, Math.min(20, bounds.height * scale / 1.2))}px`;
      const value = object.text ?? "";
      let cursor = 0;
      for (const run of object.fontRuns ?? []) {
        if (run.start < cursor || run.end <= run.start || run.end > value.length) continue;
        if (run.start > cursor) item.append(globalThis.document.createTextNode(value.slice(cursor, run.start)));
        const span = globalThis.document.createElement("span");
        span.style.fontFamily = run.renderedFamily;
        span.textContent = value.slice(run.start, run.end);
        item.append(span);
        cursor = run.end;
      }
      if (cursor < value.length) item.append(globalThis.document.createTextNode(value.slice(cursor)));
      fragment.append(item);
    }
    target.replaceChildren(fragment);
    target.hidden = false;
  }

  #clearSearchHighlight(): void {
    this.#root.querySelectorAll(".search-highlight-fragment").forEach((highlight) => highlight.remove());
    if (this.#features.interactionMode !== "text") {
      this.#root.querySelectorAll<HTMLElement>(".text-layer").forEach((layer) => {
        layer.hidden = true;
        layer.replaceChildren();
      });
    }
  }

  #refreshSearchHighlight(reveal = false): boolean {
    this.#clearSearchHighlight();
    const result = this.#searchResults[this.#searchIndex];
    if (result === undefined) return false;
    let target: HTMLElement | undefined;
    let viewport: Rect | undefined;
    let scale: number | undefined;
    let textFragments: readonly RenderedTextFragment[] | undefined;
    if (this.#pageMode === "single" && this.#rendered?.unit.index === result.object.unitIndex) {
      target = this.#element<HTMLElement>(".stage .text-layer");
      viewport = this.#rendered.frame.viewport;
      scale = this.#rendered.scale;
      textFragments = this.#rendered.frame.textFragments;
    } else {
      const record = this.#continuousPages.get(result.object.unitIndex);
      const unit = this.#info?.units[result.object.unitIndex];
      if (record !== undefined && unit !== undefined) {
        target = record.textLayer;
        viewport = { x: 0, y: 0, width: unit.width, height: unit.height };
        scale = record.renderedScale;
        textFragments = record.textFragments;
      }
    }
    if (target === undefined || viewport === undefined || scale === undefined || textFragments === undefined) return false;
    const content = globalThis.document.createDocumentFragment();
    const measure = globalThis.document.createElement("canvas").getContext("2d");
    let firstMatch: HTMLElement | undefined;
    for (const placement of textFragments) {
      if (placement.objectId !== result.object.id || placement.start === undefined || placement.end === undefined) continue;
      const fragmentStart = placement.start;
      const ranges = result.ranges
        .map(([start, end]) => [
          Math.max(0, start - fragmentStart),
          Math.min(placement.text.length, end - fragmentStart),
        ] as const)
        .filter(([start, end]) => end > start);
      if (ranges.length === 0) continue;
      const run = this.#textFragmentElement(placement, viewport, scale, measure);
      if (run === undefined) continue;
      run.classList.add("search-highlight-fragment");
      run.ariaHidden = "true";
      let cursor = 0;
      for (const [start, end] of ranges) {
        if (start > cursor) run.append(placement.text.slice(cursor, start));
        const match = globalThis.document.createElement("span");
        match.className = "search-highlight-match";
        match.textContent = placement.text.slice(start, end);
        run.append(match);
        firstMatch ??= match;
        cursor = end;
      }
      if (cursor < placement.text.length) run.append(placement.text.slice(cursor));
      content.append(run);
    }
    if (firstMatch === undefined) return false;
    target.append(content);
    target.hidden = false;
    if (reveal) firstMatch.scrollIntoView({ block: "center", inline: "center", behavior: "smooth" });
    return true;
  }

  #objectPointAt(event: MouseEvent): Readonly<{ unitIndex: number; x: number; y: number }> | undefined {
    const target = event.composedPath()[0];
    if (!(target instanceof Element)) return undefined;
    const continuousPage = target.closest<HTMLElement>(".continuous-page");
    const stage = continuousPage === null ? target.closest<HTMLElement>(".stage") : null;
    if (continuousPage === null && stage === null) return undefined;
    const record = continuousPage === null
      ? undefined
      : this.#continuousPages.get(Number(continuousPage.dataset.unitIndex));
    if (record !== undefined && !record.rendered) return undefined;
    const rendered = continuousPage === null ? this.#rendered : undefined;
    const unitIndex = record === undefined ? rendered?.unit.index : Number(continuousPage?.dataset.unitIndex);
    const surface = record?.canvas ?? this.#root.querySelector<HTMLCanvasElement>(".stage .surface");
    if (surface === null || unitIndex === undefined) return undefined;
    const unit = this.#info?.units[unitIndex];
    if (unit === undefined) return undefined;
    const bounds = surface.getBoundingClientRect();
    if (bounds.width <= 0 || bounds.height <= 0) return undefined;
    const viewport = rendered?.frame.viewport ?? { x: 0, y: 0, width: unit.width, height: unit.height };
    let x = viewport.x + (event.clientX - bounds.left) * viewport.width / bounds.width;
    let y = viewport.y + (event.clientY - bounds.top) * viewport.height / bounds.height;
    if (unit.type === "sheet") {
      const layout = this.#sheetLayout(unit);
      x = layout.columns.unmap(x);
      y = layout.rows.unmap(y);
    }
    return { unitIndex, x, y };
  }

  #queueObjectHover(event: PointerEvent): void {
    if (event.pointerType === "touch" || this.#features.interactionMode !== "object"
      || this.#document === undefined) {
      this.#clearObjectHover();
      return;
    }
    const point = this.#objectPointAt(event);
    if (point === undefined) {
      this.#clearObjectHover();
      return;
    }
    this.#hoverPoint = point;
    if (this.#hoverTimer !== undefined) return;
    this.#hoverTimer = setTimeout(() => {
      this.#hoverTimer = undefined;
      const current = this.#hoverPoint;
      if (current !== undefined) void this.#updateObjectHover(current);
    }, 40);
  }

  async #updateObjectHover(point: Readonly<{ unitIndex: number; x: number; y: number }>): Promise<void> {
    const document = this.#document;
    if (document === undefined) return;
    const revision = ++this.#hoverRevision;
    let hits: readonly HitResult[];
    try {
      hits = await document.hitTest({ ...point, limit: 16 });
    } catch {
      if (revision === this.#hoverRevision) this.#hideObjectHover();
      return;
    }
    if (revision !== this.#hoverRevision || point !== this.#hoverPoint
      || document !== this.#document || this.#features.interactionMode !== "object") return;
    const hit = hits[0];
    if (hit === undefined) this.#hideObjectHover();
    else this.#showObjectHover(hit, point.unitIndex);
  }

  async #selectObjectAt(event: MouseEvent): Promise<void> {
    if (this.#document === undefined) return;
    const textModeSheet = this.#features.interactionMode === "text"
      && this.#info?.units[this.#unitIndex]?.type === "sheet";
    if (this.#features.interactionMode !== "object" && !textModeSheet) return;
    const textSelection = (this.#root as ShadowRoot & { getSelection?: () => Selection | null }).getSelection?.()
      ?? globalThis.document.getSelection();
    if (textModeSheet && (this.#pointerSelectionRectangles.length > 0 || textSelection?.isCollapsed === false)) return;
    const point = this.#objectPointAt(event);
    if (point === undefined) {
      this.#hitRevision += 1;
      this.#clearObjectHover();
      this.#hideSelection();
      return;
    }
    const document = this.#document;
    const revision = ++this.#hitRevision;
    let hits: readonly HitResult[];
    try {
      hits = await document.hitTest({ ...point, limit: 16 });
    } catch (cause) {
      if (revision === this.#hitRevision && document === this.#document) {
        this.#emit("docviewkit-error", { cause });
      }
      return;
    }
    if (revision !== this.#hitRevision || document !== this.#document) return;
    if (textModeSheet) {
      const cell = hits.find((hit) => hit.object.type === "cell");
      if (cell !== undefined) this.#updateFormulaBar(cell.object);
      return;
    }
    if (this.#features.hyperlinks) {
      const action = hits
        .flatMap((hit) => hit.object.actions ?? [])
        .find((candidate) => candidate.trigger === "click");
      if (action !== undefined && this.#activateAction(action)) {
        event.preventDefault();
        return;
      }
    }
    const hit = hits[0];
    if (hit === undefined) {
      this.#clearObjectHover();
      this.#hideSelection();
      return;
    }
    if (hit.object.id === this.#selectionObjectId && point.unitIndex === this.#selectionUnitIndex) {
      this.#hideSelection();
      this.#showObjectHover(hit, point.unitIndex);
      return;
    }
    this.#clearObjectHover();
    this.#unitIndex = point.unitIndex;
    this.#showSelection(hit.object.bounds, false, hit.object.id, hit.object.text);
    const clipboard = this.#element<HTMLTextAreaElement>(".object-clipboard");
    clipboard.value = hit.object.text ?? "";
    if (hit.object.text === undefined) this.#workspace().focus({ preventScroll: true });
    else {
      clipboard.focus({ preventScroll: true });
      clipboard.select();
    }
    this.#updateFormulaBar(hit.object);
    this.#emit("docviewkit-objectselect", hit);
  }

  #activateAction(action: DocumentAction): boolean {
    if (action.kind === "command" && action.action === "reveal" && action.target !== undefined) {
      void this.reveal({ kind: "object", objectId: action.target }).catch((cause: unknown) => {
        this.#emit("docviewkit-error", { cause });
      });
      return true;
    }
    if (action.kind === "command" && action.action === "navigate" && action.target !== undefined) {
      if (!/^(?:0|[1-9]\d*)$/u.test(action.target)) return false;
      const unitIndex = Number(action.target);
      if (!Number.isSafeInteger(unitIndex) || unitIndex < 0 || unitIndex >= (this.#info?.units.length ?? 0)) {
        return false;
      }
      void this.#selectUnit(unitIndex);
      return true;
    }
    if (action.kind !== "hyperlink" || action.target === undefined) return false;
    const target = safeHyperlink(action.target);
    if (target === undefined) return false;
    window.open(target, "_blank", "noopener,noreferrer");
    return true;
  }

  #showSelection(bounds: Rect, reveal = false, objectId?: string, text?: string): void {
    this.#selectionBounds = copyRect(bounds);
    this.#selectionUnitIndex = this.#unitIndex;
    this.#selectionObjectId = objectId;
    this.#selectionText = text;
    this.#refreshObjectHighlight("selection");
    const selection = this.#pageMode === "continuous"
      ? this.#continuousPages.get(this.#selectionUnitIndex)?.selection
      : this.#root.querySelector<HTMLElement>(".selection");
    if (reveal && selection?.hidden === false) {
      selection.scrollIntoView({ block: "center", inline: "center", behavior: "smooth" });
    }
  }

  #showObjectHover(hit: HitResult, unitIndex: number): void {
    this.#element<HTMLElement>(".shell").dataset.objectHover = "true";
    if (hit.object.id === this.#selectionObjectId && unitIndex === this.#selectionUnitIndex) {
      this.#hoverBounds = undefined;
      this.#hoverUnitIndex = -1;
    } else {
      this.#hoverBounds = copyRect(hit.object.bounds);
      this.#hoverUnitIndex = unitIndex;
    }
    this.#refreshObjectHighlight("hover");
  }

  #refreshObjectHighlights(singleScale?: number): void {
    this.#refreshObjectHighlight("selection", singleScale);
    this.#refreshObjectHighlight("hover", singleScale);
    if (singleScale === undefined) this.#refreshSearchHighlight();
  }

  #refreshObjectHighlight(kind: "selection" | "hover", singleScale?: number): void {
    const bounds = kind === "selection" ? this.#selectionBounds : this.#hoverBounds;
    const highlightedUnitIndex = kind === "selection" ? this.#selectionUnitIndex : this.#hoverUnitIndex;
    if (this.#pageMode === "continuous") {
      for (const [unitIndex, record] of this.#continuousPages) {
        const highlight = kind === "selection" ? record.selection : record.hover;
        highlight.hidden = bounds === undefined || unitIndex !== highlightedUnitIndex;
      }
      if (bounds === undefined) return;
      const record = this.#continuousPages.get(highlightedUnitIndex);
      if (record === undefined) return;
      const highlight = kind === "selection" ? record.selection : record.hover;
      highlight.style.insetInlineStart = `${bounds.x * record.scale}px`;
      highlight.style.insetBlockStart = `${bounds.y * record.scale}px`;
      highlight.style.width = `${Math.max(2, bounds.width * record.scale)}px`;
      highlight.style.height = `${Math.max(2, bounds.height * record.scale)}px`;
      highlight.hidden = false;
      return;
    }
    const highlight = this.#element<HTMLElement>(kind === "selection" ? ".selection" : ".stage .object-hover");
    const rendered = this.#rendered;
    if (bounds === undefined || rendered === undefined
      || rendered.unit.index !== highlightedUnitIndex) {
      highlight.hidden = true;
      return;
    }
    const { frame } = rendered;
    const scale = singleScale ?? rendered.scale;
    let displayedBounds = bounds;
    if (rendered.unit.type === "sheet") {
      const layout = this.#sheetLayout(rendered.unit);
      const x = layout.columns.map(bounds.x);
      const y = layout.rows.map(bounds.y);
      displayedBounds = {
        x,
        y,
        width: layout.columns.map(bounds.x + bounds.width) - x,
        height: layout.rows.map(bounds.y + bounds.height) - y,
      };
    }
    highlight.style.insetInlineStart = `${(displayedBounds.x - frame.viewport.x) * scale}px`;
    highlight.style.insetBlockStart = `${(displayedBounds.y - frame.viewport.y) * scale}px`;
    highlight.style.width = `${Math.max(2, displayedBounds.width * scale)}px`;
    highlight.style.height = `${Math.max(2, displayedBounds.height * scale)}px`;
    highlight.hidden = false;
  }

  #hideSelection(): void {
    this.#selectionBounds = undefined;
    this.#selectionUnitIndex = -1;
    this.#selectionObjectId = undefined;
    this.#selectionText = undefined;
    const clipboard = this.#element<HTMLTextAreaElement>(".object-clipboard");
    clipboard.value = "";
    if (this.#root.activeElement === clipboard) this.#workspace().focus({ preventScroll: true });
    this.#refreshObjectHighlight("selection");
  }

  #hideObjectHover(): void {
    this.#element<HTMLElement>(".shell").dataset.objectHover = "false";
    this.#hoverBounds = undefined;
    this.#hoverUnitIndex = -1;
    this.#refreshObjectHighlight("hover");
  }

  #clearObjectHover(): void {
    if (this.#hoverTimer !== undefined) clearTimeout(this.#hoverTimer);
    this.#hoverTimer = undefined;
    this.#hoverPoint = undefined;
    this.#hoverRevision += 1;
    this.#hideObjectHover();
  }

  #updateFormulaBar(object?: DocumentObject): void {
    const source = object?.source;
    let address = "A1";
    if (source?.format === "xlsx" && source.address !== undefined) {
      address = source.address;
    } else if (object?.type === "cell" && source !== undefined
      && "row" in source && "column" in source
      && source.row !== undefined && source.column !== undefined) {
      address = `${columnLabel(source.column)}${source.row + 1}`;
    }
    this.#element<HTMLOutputElement>(".cell-address").value = address;
    this.#element<HTMLOutputElement>(".formula-value").value = source?.format === "xlsx"
      ? source.formula ?? object?.text ?? ""
      : object?.text ?? "";
  }

  #updateSearchPosition(): void {
    const position = this.#element<HTMLOutputElement>(".search-position");
    position.value = this.#searchResults.length === 0
      ? "0 / 0"
      : formatMessage(this.#messages.resultPosition, {
          current: this.#searchIndex + 1,
          total: this.#searchResults.length,
        });
    for (const action of ["previous-result", "next-result"]) {
      const button = this.#root.querySelector<HTMLButtonElement>(`[data-action="${action}"]`);
      if (button !== null) button.disabled = this.#searchResults.length === 0;
    }
  }

  #toggleDiagnostics(force?: boolean): void {
    if (!this.#features.diagnostics) return;
    const panel = this.#element<HTMLElement>(".diagnostic-panel");
    panel.hidden = !(force ?? panel.hidden);
    this.#root.querySelector<HTMLButtonElement>('[data-action="diagnostics"]')?.setAttribute("aria-pressed", String(!panel.hidden));
  }

  #renderDiagnostics(): void {
    const diagnostics = this.#allDiagnostics();
    const list = this.#root.querySelector<HTMLElement>(".diagnostic-list");
    if (list === null) return;
    list.replaceChildren();
    if (diagnostics.length === 0) {
      const item = document.createElement("li");
      item.textContent = this.#messages.noDiagnostics;
      list.append(item);
      return;
    }
    for (const diagnostic of diagnostics.slice(0, 100)) list.append(this.#diagnosticItem(diagnostic));
  }

  #diagnosticItem(diagnostic: Diagnostic): HTMLLIElement {
    const item = document.createElement("li");
    item.className = "diagnostic-item";
    const code = document.createElement("span");
    code.className = "diagnostic-code";
    code.textContent = diagnostic.code;
    const message = document.createElement("span");
    message.textContent = diagnostic.message;
    item.append(code, message);
    return item;
  }

  async #printDocument(): Promise<void> {
    const document = this.#document;
    const info = this.#info;
    if (this.#printing || document === undefined || info === undefined) return;
    this.#clearPrintView();
    const controller = new AbortController();
    this.#printController = controller;
    this.#printing = true;
    const button = this.#root.querySelector<HTMLButtonElement>('[data-action="print"]');
    button?.setAttribute("aria-busy", "true");
    this.#updateControls();
    let printStarted = false;
    try {
      const printFrame = globalThis.document.createElement("iframe");
      printFrame.title = this.#messages.print;
      printFrame.style.cssText = "position:fixed;right:100%;bottom:100%;width:1px;height:1px;border:0";
      globalThis.document.body.append(printFrame);
      this.#printFrame = printFrame;
      const printDocument = printFrame.contentDocument;
      if (printDocument === null) throw new Error("Print document is unavailable");
      const style = printDocument.createElement("style");
      style.textContent = `
        @page { margin: 0; }
        html, body { margin: 0; background: #ffffff; }
        .print-page { display: flex; width: 100%; box-sizing: border-box; justify-content: center; break-after: page; break-inside: avoid-page; }
        .print-page:last-child { break-after: auto; }
        .print-page img { display: block; width: auto; height: auto; max-width: 99vw; max-height: 99vh; object-fit: contain; }
        .print-page-error { min-height: 80vh; align-items: center; color: #17191d; }
      `;
      printDocument.head.append(style);
      for (const unit of info.units) {
        if (document !== this.#document) throw new DOMException("Superseded", "AbortError");
        const sheetSizes = unit.type === "sheet" ? this.#sheetSizeRequest(unit) : undefined;
        const printPages = unit.type === "sheet" ? sheetPrintPages(unit, sheetSizes) : [undefined];
        for (const sheetPage of printPages) {
          const page = printDocument.createElement("section");
          page.className = "print-page";
          if (sheetPage !== undefined) {
            page.style.width = `${sheetPage.paper.width}px`;
            page.style.height = `${sheetPage.paper.height}px`;
            page.style.padding = `${sheetPage.margins.top}px ${sheetPage.margins.right}px ${sheetPage.margins.bottom}px ${sheetPage.margins.left}px`;
          }
          let frame: RenderResult | undefined;
          try {
            const width = sheetPage?.width ?? unit.width;
            const height = sheetPage?.height ?? unit.height;
            const scale = Math.max(.01, Math.min(2, PRINT_RENDER_LONG_EDGE / Math.max(1, width, height)));
            const canvas = globalThis.document.createElement("canvas");
            canvas.width = Math.max(1, Math.ceil(width * scale));
            canvas.height = Math.max(1, Math.ceil(height * scale));
            const context = canvas.getContext("2d");
            if (context === null) throw new Error("Canvas 2D is unavailable");
            context.fillStyle = "#ffffff";
            context.fillRect(0, 0, canvas.width, canvas.height);
            const fragments = sheetPage?.fragments ?? [{ viewport: undefined, x: 0, y: 0 }];
            for (const fragment of fragments) {
              frame = await document.render({
                unitIndex: unit.index,
                ...(fragment.viewport === undefined ? {} : { viewport: fragment.viewport }),
                scale, pixelRatio: 1, background: "#ffffff",
                ...this.#watermarkRequest(),
                ...(sheetSizes === undefined ? {} : { sheetSizes }),
              }, { signal: controller.signal, priority: "interactive", supersedeKey: "viewer-print" });
              context.drawImage(frame.bitmap, fragment.x * scale, fragment.y * scale,
                (fragment.viewport?.width ?? width) * scale, (fragment.viewport?.height ?? height) * scale);
              frame.bitmap.close();
              frame = undefined;
            }
            const blob = await new Promise<Blob>((resolve, reject) => canvas.toBlob((value) => {
              if (value === null) reject(new Error("Print image encoding failed"));
              else resolve(value);
            }, "image/png"));
            canvas.width = 1;
            canvas.height = 1;
            const url = URL.createObjectURL(blob);
            this.#printUrls.push(url);
            const image = printDocument.createElement("img");
            image.alt = "";
            image.src = url;
            if (sheetPage !== undefined) {
              image.style.width = `${sheetPage.width * sheetPage.scale}px`;
              image.style.height = `${sheetPage.height * sheetPage.scale}px`;
            }
            await image.decode();
            page.append(image);
          } catch (cause) {
            if (controller.signal.aborted) throw cause;
            page.classList.add("print-page-error");
            page.textContent = `${this.#messages.renderFailed} (${unit.index + 1})`;
            this.#emit("docviewkit-error", { cause, unitIndex: unit.index });
          } finally {
            frame?.bitmap.close();
          }
          printDocument.body.append(page);
        }
      }
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
      const printWindow = printFrame.contentWindow;
      if (printWindow === null) throw new Error("Print window is unavailable");
      printWindow.addEventListener("afterprint", () => this.#clearPrintView(), { once: true });
      printStarted = true;
      try {
        printWindow.focus();
        printWindow.print();
      } catch (cause) {
        printStarted = false;
        throw cause;
      }
    } catch (cause) {
      if (!(cause instanceof DOMException && cause.name === "AbortError")) this.#emit("docviewkit-error", { cause });
    } finally {
      if (!printStarted) this.#clearPrintView();
      if (this.#printController === controller) this.#printController = undefined;
      this.#printing = false;
      button?.removeAttribute("aria-busy");
      this.#updateControls();
    }
  }

  #clearPrintView(): void {
    for (const url of this.#printUrls) URL.revokeObjectURL(url);
    this.#printUrls = [];
    this.#printFrame?.remove();
    this.#printFrame = undefined;
  }

  async #toggleFullscreen(): Promise<void> {
    if (document.fullscreenElement === this) await document.exitFullscreen();
    else await this.requestFullscreen();
    this.#localize();
  }

  #setStatus(status: ViewerStatus, detail = ""): void {
    this.#status = status;
    const empty = this.#element<HTMLElement>(".empty");
    const loading = this.#element<HTMLElement>(".loading");
    const error = this.#element<HTMLElement>(".error");
    empty.hidden = status !== "idle" && status !== "destroyed";
    loading.hidden = status !== "loading";
    error.hidden = status !== "error";
    this.#element<HTMLElement>(".stage-wrap").hidden = status !== "ready" || this.#pageMode === "continuous";
    this.#element<HTMLElement>(".continuous-view").hidden = status !== "ready" || this.#pageMode !== "continuous";
    if (status === "error") {
      this.#element<HTMLElement>(".error .state-title").textContent = this.#messages.renderFailed;
      this.#element<HTMLElement>(".error .state-hint").textContent = detail || this.#messages.unsupported;
    }
    this.#setMessage(status === "loading" ? this.#messages.loading : detail);
    this.#updateControls();
  }

  #setMessage(value: string): void {
    this.#element<HTMLElement>(".status-message").textContent = value;
  }

  #updateControls(): void {
    const count = this.#info?.units.length ?? 0;
    const current = count === 0 ? 0 : this.#unitIndex + 1;
    const page = formatMessage(this.#messages.pagePosition, { current, total: count });
    this.#root.querySelectorAll<HTMLOutputElement>(".page-position").forEach((output) => {
      output.value = page;
    });
    const disabled = this.#document === undefined || this.#status === "loading";
    const zoomInput = this.#element<HTMLInputElement>(".zoom-input");
    if (this.#root.activeElement !== zoomInput) this.#syncZoomInput();
    zoomInput.disabled = disabled;
    for (const action of ["search", "zoom-out", "zoom-in", "fit", "print", "fullscreen", "diagnostics"]) {
      const button = this.#root.querySelector<HTMLButtonElement>(`[data-action="${action}"]`);
      if (button !== null) button.disabled = disabled || (action === "print" && this.#printing);
    }
    for (const action of ["previous", "mobile-previous"]) {
      const button = this.#root.querySelector<HTMLButtonElement>(`[data-action="${action}"]`);
      if (button !== null) button.disabled = disabled || this.#unitIndex <= 0;
    }
    for (const action of ["next", "mobile-next"]) {
      const button = this.#root.querySelector<HTMLButtonElement>(`[data-action="${action}"]`);
      if (button !== null) button.disabled = disabled || this.#unitIndex >= count - 1;
    }
    this.#updatePageModeControls();
  }

  #closeRenderedFrame(): void {
    const frame = this.#rendered?.frame;
    if (frame !== undefined) frame.bitmap.close();
    this.#rendered = undefined;
    this.#root.querySelector<HTMLElement>(".stage .text-layer")?.replaceChildren();
  }

  #clearThumbnails(): void {
    this.#thumbnailObserver?.disconnect();
    for (const controller of this.#thumbnailRequests.values()) controller.abort();
    this.#thumbnailRequests.clear();
    this.#prefetchedPages.clear();
    if (this.#pagePrefetchTimer !== undefined) clearTimeout(this.#pagePrefetchTimer);
    this.#pagePrefetchTimer = undefined;
    this.#navigationStart = -1;
    this.#navigationEnd = -1;
  }

  #workspace(): HTMLElement {
    return this.#element<HTMLElement>(".workspace");
  }

  #element<T extends Element>(selector: string): T {
    const value = this.#root.querySelector<T>(selector);
    if (value === null) throw new Error(`Viewer template is missing ${selector}`);
    return value;
  }

  #emit(type: string, detail: unknown): void {
    this.dispatchEvent(new CustomEvent(type, { detail, bubbles: true, composed: true }));
  }

  #emitState(): void {
    this.#emit("docviewkit-statechange", { state: this.state });
  }
}

export function defineDocViewKitViewer(tagName = "docviewkit-viewer"): void {
  if (customElements.get(tagName) === undefined) customElements.define(tagName, DocViewKitViewerElement);
}

if (typeof customElements !== "undefined") defineDocViewKitViewer();
