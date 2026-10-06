import { createHash } from "node:crypto";
import { createReadStream } from "node:fs";
import {
  copyFile,
  cp,
  lstat,
  mkdir,
  mkdtemp,
  open as openFile,
  readFile,
  readdir,
  rename,
  rm,
  stat,
  writeFile,
} from "node:fs/promises";
import { arch, platform, release } from "node:os";
import { basename, dirname, extname, resolve } from "node:path";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";

const RASTER_DPI = 96;
const COMMAND_OUTPUT_LIMIT = 1024 * 1024;
const NATIVE_EXPORT_TIMEOUT_MS = 15 * 60 * 1000;
const REVIEWER_ID_PATTERN = /^[A-Za-z0-9](?:[A-Za-z0-9._@:-]{0,126}[A-Za-z0-9])?$/u;

const APPLICATIONS = Object.freeze({
  powerpoint: {
    name: "Microsoft PowerPoint",
    oracleSuite: "microsoft-office",
    expectedBundleId: "com.microsoft.Powerpoint",
    path: "/Applications/Microsoft PowerPoint.app",
    captureSemantics: {
      unitType: "slide",
      description: "One PDF page per slide in source order, including hidden slides; slide-only, color, final static build state.",
    },
  },
  word: {
    name: "Microsoft Word",
    oracleSuite: "microsoft-office",
    expectedBundleId: "com.microsoft.Word",
    path: "/Applications/Microsoft Word.app",
    captureSemantics: {
      unitType: "page",
      description: "Paginated read-only PDF export with external-link and print-time field updates disabled.",
    },
  },
  excel: {
    name: "Microsoft Excel",
    oracleSuite: "microsoft-office",
    expectedBundleId: "com.microsoft.Excel",
    path: "/Applications/Microsoft Excel.app",
    printPages: true,
    captureSemantics: {
      unitType: "print-page",
      description: "One PDF per visible worksheet using its stored print settings and cached formula values.",
      warning: "Excel PDF pages are print pages, not an interactive worksheet viewport.",
    },
  },
  keynote: {
    name: "Keynote",
    oracleSuite: "apple-iwork",
    expectedBundleId: "com.apple.iWork.Keynote",
    path: "/Applications/Keynote.app",
    captureSemantics: {
      unitType: "slide",
      description: "Individual-slide PDF export in source order, including skipped slides, without notes, comments, borders, numbers, dates, or intermediate build stages.",
    },
  },
  pages: {
    name: "Pages",
    oracleSuite: "apple-iwork",
    expectedBundleId: "com.apple.iWork.Pages",
    path: "/Applications/Pages.app",
    captureSemantics: {
      unitType: "page",
      description: "Paginated PDF export at Best image quality, without comments or smart annotations.",
    },
  },
  numbers: {
    name: "Numbers",
    oracleSuite: "apple-iwork",
    expectedBundleId: "com.apple.iWork.Numbers",
    path: "/Applications/Numbers.app",
    printPages: true,
    captureSemantics: {
      unitType: "print-page",
      description: "Document-level PDF export at Best image quality, without comments.",
      warning: "Numbers PDF pages are print pages with no guaranteed page-to-sheet mapping, not an interactive sheet viewport.",
    },
  },
  "wps-writer": {
    name: "WPS Office",
    oracleSuite: "wps-office",
    expectedBundleId: "com.kingsoft.wpsoffice.mac",
    path: "/Applications/wpsoffice.app",
    manualOnly: true,
    captureSemantics: {
      unitType: "page",
      description: "Paginated WPS Writer PDF export reviewed in WPS Office before import.",
    },
  },
  "wps-spreadsheets": {
    name: "WPS Office",
    oracleSuite: "wps-office",
    expectedBundleId: "com.kingsoft.wpsoffice.mac",
    path: "/Applications/wpsoffice.app",
    manualOnly: true,
    printPages: true,
    captureSemantics: {
      unitType: "sheet",
      description: "Content-only WPS Spreadsheets viewport reviewed in WPS Office before import.",
    },
  },
  "wps-presentation": {
    name: "WPS Office",
    oracleSuite: "wps-office",
    expectedBundleId: "com.kingsoft.wpsoffice.mac",
    path: "/Applications/wpsoffice.app",
    manualOnly: true,
    captureSemantics: {
      unitType: "slide",
      description: "WPS Presentation PDF export reviewed in WPS Office before import.",
    },
  },
});

const EXTENSION_TO_APPLICATION = new Map([
  [".pptx", "powerpoint"],
  [".pptm", "powerpoint"],
  [".potx", "powerpoint"],
  [".potm", "powerpoint"],
  [".ppsx", "powerpoint"],
  [".ppsm", "powerpoint"],
  [".ppt", "powerpoint"],
  [".odp", "powerpoint"],
  [".docx", "word"],
  [".docm", "word"],
  [".dotx", "word"],
  [".dotm", "word"],
  [".doc", "word"],
  [".odt", "word"],
  [".ott", "word"],
  [".rtf", "word"],
  [".xlsx", "excel"],
  [".xlsm", "excel"],
  [".xltx", "excel"],
  [".xltm", "excel"],
  [".xls", "excel"],
  [".ods", "excel"],
  [".key", "keynote"],
  [".pages", "pages"],
  [".numbers", "numbers"],
  [".wps", "wps-writer"],
  [".et", "wps-spreadsheets"],
  [".dps", "wps-presentation"],
]);

const APPLE_SCRIPTS = Object.freeze({
  powerpoint: String.raw`
on run argv
  set inputFile to POSIX file (item 1 of argv)
  set outputPath to (POSIX file (item 2 of argv)) as text
  set previousSecurity to missing value
  set openedPresentation to missing value
  set presentationOpened to false
  set currentStage to "initialize"
  tell application "Microsoft PowerPoint"
    try
      with timeout of 270 seconds
        set my currentStage to "read-automation-security"
        set my previousSecurity to automation security
        set my currentStage to "set-automation-security"
        set automation security to msoAutomationSecurityForceDisable
        set my currentStage to "open"
        repeat with priorPresentation in (get presentations)
          set priorPath to full name of priorPresentation
          if priorPath is (item 1 of argv) or priorPath is (my inputFile as text) then close priorPresentation saving no
        end repeat
        set previousPresentationNames to name of every presentation
        set previousPresentationCount to count of presentations
        open (my inputFile)
        if (count of presentations) <= previousPresentationCount then error "Requested source did not create a presentation" number 8002
        set candidatePresentation to active presentation
        if (name of candidatePresentation) is in previousPresentationNames then error "Requested source did not activate a new presentation" number 8002
        set openedPath to full name of candidatePresentation
        -- ponytail: templates create unnamed presentations; only creation is observable through this API.
        set isTemplate to ((item 1 of argv) ends with ".potx") or ((item 1 of argv) ends with ".potm")
        if not isTemplate and openedPath is not (item 1 of argv) and openedPath is not (my inputFile as text) then error "Requested source identity mismatch: " & openedPath number 8003
        set my openedPresentation to candidatePresentation
        set my presentationOpened to true
        set my currentStage to "print-options"
        tell print options of my openedPresentation
          set output type to print slides
          set fit to page to false
          set frame slides to false
          set print hidden slides to true
          set print color type to print color
          set range type to print range all
        end tell
        set my currentStage to "save"
        save my openedPresentation in (my outputPath) as save as PDF
        set my currentStage to "close"
        close my openedPresentation saving no
        set my presentationOpened to false
        set my currentStage to "restore-automation-security"
        if my previousSecurity is not missing value then set automation security to my previousSecurity
      end timeout
    on error errorMessage number errorNumber
      set originalErrorMessage to "PowerPoint stage " & (my currentStage) & ": " & (errorMessage as text)
      set originalErrorNumber to errorNumber as integer
      try
        with timeout of 10 seconds
          if my presentationOpened then
            close my openedPresentation saving no
            set my presentationOpened to false
          end if
        end timeout
      end try
      try
        with timeout of 10 seconds
          if my previousSecurity is not missing value then set automation security to my previousSecurity
        end timeout
      end try
      error originalErrorMessage number originalErrorNumber
    end try
  end tell
end run
`,
  word: String.raw`
on run argv
  set inputFile to POSIX file (item 1 of argv)
  set outputPath to item 2 of argv
  set openedDocument to missing value
  set previousSecurity to missing value
  set previousAlerts to missing value
  set wordSettings to missing value
  set previousUpdateLinksAtOpen to missing value
  set previousUpdateFieldsAtPrint to missing value
  set previousUpdateLinksAtPrint to missing value
  tell application "Microsoft Word"
    try
      with timeout of 870 seconds
        set my previousSecurity to automation security
        set my previousAlerts to display alerts
        set my wordSettings to settings
        try
          set my previousUpdateLinksAtOpen to update links at open of my wordSettings
        end try
        try
          set my previousUpdateFieldsAtPrint to update fields at print of my wordSettings
        end try
        try
          set my previousUpdateLinksAtPrint to update links at print of my wordSettings
        end try
        set automation security to msoAutomationSecurityForceDisable
        set display alerts to alerts none
        try
          set update links at open of my wordSettings to false
        end try
        try
          set update fields at print of my wordSettings to false
        end try
        try
          set update links at print of my wordSettings to false
        end try
        repeat with priorDocument in (get documents)
          if (full name of priorDocument) is (my inputFile as text) then close priorDocument saving no
        end repeat
        open inputFile read only true add to recent files false confirm conversions false
        set candidateDocument to active document
        if (full name of candidateDocument) is not (my inputFile as text) then error "Requested source identity mismatch" number 8003
        set my openedDocument to candidateDocument
        save as my openedDocument file name (my outputPath) file format format PDF add to recent files false
        close my openedDocument saving no
        if my previousSecurity is not missing value then set automation security to my previousSecurity
        if my previousAlerts is not missing value then set display alerts to my previousAlerts
        try
          if my previousUpdateLinksAtOpen is not missing value then set update links at open of my wordSettings to my previousUpdateLinksAtOpen
        end try
        try
          if my previousUpdateFieldsAtPrint is not missing value then set update fields at print of my wordSettings to my previousUpdateFieldsAtPrint
        end try
        try
          if my previousUpdateLinksAtPrint is not missing value then set update links at print of my wordSettings to my previousUpdateLinksAtPrint
        end try
      end timeout
    on error errorMessage number errorNumber
      set originalErrorMessage to errorMessage as text
      set originalErrorNumber to errorNumber as integer
      try
        with timeout of 10 seconds
          if my openedDocument is not missing value then close active document saving no
        end timeout
      end try
      try
        with timeout of 10 seconds
          if my previousSecurity is not missing value then set automation security to my previousSecurity
          if my previousAlerts is not missing value then set display alerts to my previousAlerts
          try
            if my previousUpdateLinksAtOpen is not missing value then set update links at open of my wordSettings to my previousUpdateLinksAtOpen
          end try
          try
            if my previousUpdateFieldsAtPrint is not missing value then set update fields at print of my wordSettings to my previousUpdateFieldsAtPrint
          end try
          try
            if my previousUpdateLinksAtPrint is not missing value then set update links at print of my wordSettings to my previousUpdateLinksAtPrint
          end try
        end timeout
      end try
      error originalErrorMessage number originalErrorNumber
    end try
  end tell
end run
`,
  excel: String.raw`
on run argv
  set inputPath to item 1 of argv
  set inputFile to POSIX file inputPath
  set outputDirectory to item 2 of argv
  set currentStage to "initialize"
  set openedWorkbook to missing value
  set previousSecurity to missing value
  set previousAlerts to missing value
  set previousScreenUpdating to missing value
  set previousCalculation to missing value
  set previousAskToUpdateLinks to missing value
  tell application "Microsoft Excel"
    try
      with timeout of 270 seconds
        set previousSecurity to automation security
        set previousAlerts to display alerts
        set previousScreenUpdating to screen updating
        set previousCalculation to calculation
        set previousAskToUpdateLinks to ask to update links
        set automation security to msoAutomationSecurityForceDisable
        set display alerts to false
        set screen updating to false
        set calculation to calculation manual
        set ask to update links to false
        set currentStage to "cleanup"
        repeat with priorWorkbook in (get workbooks)
          set priorPath to full name of priorWorkbook
          if priorPath is inputPath or priorPath is (inputFile as text) then close priorWorkbook saving no
        end repeat
        set currentStage to "open"
        set candidateWorkbook to open workbook workbook file name inputPath update links do not update links read only true ignore read only recommended true editable true add to mru false
        set currentStage to "source-identity"
        set openedPath to full name of candidateWorkbook
        if openedPath is not inputPath and openedPath is not (inputFile as text) then error "Requested source identity mismatch" number 8003
        set openedWorkbook to candidateWorkbook
        set currentStage to "export"
        set exportedCount to 0
        repeat with sheetIndex from 1 to (count of worksheets of openedWorkbook)
          set currentSheet to worksheet sheetIndex of openedWorkbook
          if visible of currentSheet is sheet visible then
            activate object currentSheet
            set sheetPdfPath to outputDirectory & "/sheet-" & (sheetIndex as text) & ".pdf"
            save as currentSheet filename sheetPdfPath file format PDF file format
            set exportedCount to exportedCount + 1
          end if
        end repeat
        if exportedCount is 0 then error "Workbook has no visible worksheets" number 8001
        close openedWorkbook saving no
        set openedWorkbook to missing value
        if previousSecurity is not missing value then set automation security to previousSecurity
        if previousAlerts is not missing value then set display alerts to previousAlerts
        if previousScreenUpdating is not missing value then set screen updating to previousScreenUpdating
        if previousCalculation is not missing value then set calculation to previousCalculation
        if previousAskToUpdateLinks is not missing value then set ask to update links to previousAskToUpdateLinks
      end timeout
    on error errorMessage number errorNumber
      set originalErrorMessage to "Excel stage " & currentStage & ": " & (errorMessage as text)
      set originalErrorNumber to errorNumber as integer
      try
        with timeout of 10 seconds
          if openedWorkbook is not missing value then close openedWorkbook saving no
        end timeout
      end try
      try
        with timeout of 10 seconds
          if previousSecurity is not missing value then set automation security to previousSecurity
          if previousAlerts is not missing value then set display alerts to previousAlerts
          if previousScreenUpdating is not missing value then set screen updating to previousScreenUpdating
          if previousCalculation is not missing value then set calculation to previousCalculation
          if previousAskToUpdateLinks is not missing value then set ask to update links to previousAskToUpdateLinks
        end timeout
      end try
      error originalErrorMessage number originalErrorNumber
    end try
  end tell
end run
`,
  keynote: String.raw`
on run argv
  set inputFile to POSIX file (item 1 of argv)
  set outputFile to POSIX file (item 2 of argv)
  set openedDocument to missing value
  tell application "Keynote"
    try
      with timeout of 270 seconds
        set openedDocument to open inputFile
        export openedDocument to outputFile as PDF with properties {export style:IndividualSlides, all stages:false, skipped slides:true, borders:false, slide numbers:false, include comments:false, PDF image quality:Best}
        close openedDocument saving no
        set openedDocument to missing value
      end timeout
    on error errorMessage number errorNumber
      set originalErrorMessage to errorMessage as text
      set originalErrorNumber to errorNumber as integer
      try
        with timeout of 15 seconds
          if openedDocument is not missing value then close openedDocument saving no
        end timeout
      end try
      error originalErrorMessage number originalErrorNumber
    end try
  end tell
end run
`,
  pages: String.raw`
on run argv
  set inputFile to POSIX file (item 1 of argv)
  set outputFile to POSIX file (item 2 of argv)
  set openedDocument to missing value
  tell application "Pages"
    try
      with timeout of 270 seconds
        set openedDocument to open inputFile
        export openedDocument to outputFile as PDF with properties {image quality:Best, include comments:false, include annotations:false}
        close openedDocument saving no
        set openedDocument to missing value
      end timeout
    on error errorMessage number errorNumber
      set originalErrorMessage to errorMessage as text
      set originalErrorNumber to errorNumber as integer
      try
        with timeout of 15 seconds
          if openedDocument is not missing value then close openedDocument saving no
        end timeout
      end try
      error originalErrorMessage number originalErrorNumber
    end try
  end tell
end run
`,
  numbers: String.raw`
on run argv
  set inputFile to POSIX file (item 1 of argv)
  set outputFile to POSIX file (item 2 of argv)
  set openedDocument to missing value
  tell application "Numbers"
    try
      with timeout of 270 seconds
        set openedDocument to open inputFile
        export openedDocument to outputFile as PDF with properties {image quality:Best, include comments:false}
        close openedDocument saving no
        set openedDocument to missing value
      end timeout
    on error errorMessage number errorNumber
      set originalErrorMessage to errorMessage as text
      set originalErrorNumber to errorNumber as integer
      try
        with timeout of 15 seconds
          if openedDocument is not missing value then close openedDocument saving no
        end timeout
      end try
      error originalErrorMessage number originalErrorNumber
    end try
  end tell
end run
`,
});

function usage() {
  return `Usage:
  node scripts/export-native-reference.mjs <fixture> <output-directory> [--allow-print-pages]
  node scripts/export-native-reference.mjs <fixture> <output-directory> --sheet-viewport <reviewed.png> --range <A1:H40> --unit-index <0> --reviewed-by <stable-id>
  node scripts/export-native-reference.mjs <fixture> <output-directory> --native-pdf <reviewed.pdf> --reviewed-by <stable-id>

Builds a reviewed visual reference through the owning macOS application or from
a manually reviewed native PDF, then rasters every PDF page to PNG at ${RASTER_DPI} DPI.
Supported inputs:
  PPTX/PPT/ODP  Microsoft PowerPoint (also PPTM, POTX, POTM, PPSX, PPSM)
  DOCX/DOC/ODT/RTF  Microsoft Word (also DOCM, DOTX, DOTM, OTT)
  XLSX/XLS/ODS  Microsoft Excel (also XLSM, XLTX, XLTM) (requires --allow-print-pages)
  KEY           Keynote
  PAGES         Pages
  NUMBERS       Numbers (requires --allow-print-pages)
  WPS/DPS       WPS Office (reviewed --native-pdf import only)
  ET            WPS Office (reviewed --sheet-viewport import only)

Requires the corresponding installed application for version/build metadata.
Automated exports also require a logged-in macOS session with Automation
permission; PDF export/import paths require pdftoppm in PATH. Do not feed
malformed, malicious, or otherwise untrusted documents to native applications.

Spreadsheet PDF export is paginated print output, not a worksheet viewport.
Prefer an explicit viewport oracle. Use --allow-print-pages only for a reviewed
single-page print-area fixture or when print-page fidelity is the intended test.

--sheet-viewport imports an already reviewed Excel/Numbers/WPS Spreadsheets content-only PNG
verbatim. It is valid only for XLSX, XLS, ODS, NUMBERS, or ET, requires an explicit cell
range and zero-based sheet unit index, and never invokes native PDF export or
pdftoppm. --reviewed-by is required.

--native-pdf imports a PDF that the user has already exported and reviewed in
the corresponding native application. It is valid only for PPTX, PPT, DOCX,
DOC, ODP, ODT, RTF, KEY, PAGES, WPS, or DPS and does not automate or inspect that application. The script
checks the PDF signature and byte integrity but does not verify its provenance
or review. The installed application version/build is recorded as environment
metadata. This option cannot be combined with --sheet-viewport, --range, or
--allow-print-pages. --reviewed-by is required for both manual import modes and
must be a stable 1-128 character identifier using ASCII letters, digits, '.',
'_', '@', ':', or '-'. Its attestation is human-controlled; this script records
the identifier and hashes but does not verify reviewer identity or provenance.`;
}

function parseArguments(values) {
  if (values.includes("--help") || values.includes("-h")) return { help: true };
  const positional = [];
  let allowPrintPages = false;
  let sheetViewport;
  let nativePdf;
  let range;
  let unitIndex;
  let reviewedBy;
  for (let index = 0; index < values.length; index += 1) {
    const value = values[index];
    if (value === "--allow-print-pages") allowPrintPages = true;
    else if (value === "--sheet-viewport" || value === "--range" || value === "--native-pdf" || value === "--unit-index" || value === "--reviewed-by") {
      const optionValue = values[++index];
      if (optionValue === undefined || optionValue.startsWith("-")) throw new Error(`Missing ${value} value\n${usage()}`);
      if (value === "--sheet-viewport") sheetViewport = resolve(optionValue);
      else if (value === "--native-pdf") nativePdf = resolve(optionValue);
      else if (value === "--unit-index") unitIndex = Number(optionValue);
      else if (value === "--reviewed-by") {
        if (reviewedBy !== undefined) throw new Error("--reviewed-by may be provided only once");
        reviewedBy = optionValue;
      }
      else range = optionValue;
    } else if (value.startsWith("-")) throw new Error(`Unknown option ${value}\n${usage()}`);
    else positional.push(value);
  }
  if (positional.length !== 2) throw new Error(usage());
  if ((sheetViewport === undefined) !== (range === undefined) || (sheetViewport === undefined) !== (unitIndex === undefined)) {
    throw new Error("--sheet-viewport, --range, and --unit-index must be provided together");
  }
  if (unitIndex !== undefined && (!Number.isSafeInteger(unitIndex) || unitIndex < 0)) {
    throw new Error("--unit-index must be a non-negative integer");
  }
  if (sheetViewport !== undefined && allowPrintPages) {
    throw new Error("--sheet-viewport cannot be combined with --allow-print-pages");
  }
  if (nativePdf !== undefined && (sheetViewport !== undefined || range !== undefined || allowPrintPages)) {
    throw new Error("--native-pdf cannot be combined with --sheet-viewport, --range, or --allow-print-pages");
  }
  if (reviewedBy !== undefined && !REVIEWER_ID_PATTERN.test(reviewedBy)) {
    throw new Error("--reviewed-by must be a stable 1-128 character identifier using ASCII letters, digits, '.', '_', '@', ':', or '-', with no whitespace or leading/trailing punctuation");
  }
  const manualImport = nativePdf !== undefined || sheetViewport !== undefined;
  if (manualImport && reviewedBy === undefined) {
    throw new Error("--reviewed-by is required with --native-pdf or --sheet-viewport");
  }
  if (!manualImport && reviewedBy !== undefined) {
    throw new Error("--reviewed-by is valid only with --native-pdf or --sheet-viewport");
  }
  return {
    help: false,
    fixture: resolve(positional[0]),
    output: resolve(positional[1]),
    allowPrintPages,
    sheetViewport,
    nativePdf,
    range,
    unitIndex,
    reviewedBy,
  };
}

function normalizeSheetRange(value) {
  const match = /^([A-Za-z]{1,3})([1-9]\d{0,6}):([A-Za-z]{1,3})([1-9]\d{0,6})$/u.exec(value);
  if (match === null) throw new Error(`Invalid sheet range ${JSON.stringify(value)}; expected a rectangle such as A1:H40`);
  const columnNumber = (label) => [...label.toUpperCase()].reduce((number, character) => number * 26 + character.charCodeAt(0) - 64, 0);
  const startColumn = columnNumber(match[1]);
  const startRow = Number(match[2]);
  const endColumn = columnNumber(match[3]);
  const endRow = Number(match[4]);
  if (startColumn > 16_384 || endColumn > 16_384 || startRow > 1_048_576 || endRow > 1_048_576) {
    throw new Error(`Sheet range ${JSON.stringify(value)} exceeds XFD1048576`);
  }
  if (startColumn > endColumn || startRow > endRow) {
    throw new Error(`Sheet range ${JSON.stringify(value)} must run from its top-left cell to its bottom-right cell`);
  }
  return `${match[1].toUpperCase()}${startRow}:${match[3].toUpperCase()}${endRow}`;
}

async function sha256File(file) {
  const hash = createHash("sha256");
  for await (const chunk of createReadStream(file)) hash.update(chunk);
  return hash.digest("hex");
}

function runCommand(program, arguments_, { input, timeoutMs = 30_000 } = {}) {
  return new Promise((resolveCommand, rejectCommand) => {
    const child = spawn(program, arguments_, { stdio: ["pipe", "pipe", "pipe"] });
    let stdout = Buffer.alloc(0);
    let stderr = Buffer.alloc(0);
    let settled = false;
    let timer;
    const finish = (callback) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      callback();
    };
    const append = (current, chunk) => {
      const next = Buffer.concat([current, chunk]);
      if (next.length > COMMAND_OUTPUT_LIMIT) {
        child.kill("SIGTERM");
        finish(() => rejectCommand(new Error(`${program} exceeded the command output limit`)));
      }
      return next;
    };
    child.stdout.on("data", (chunk) => { stdout = append(stdout, chunk); });
    child.stderr.on("data", (chunk) => { stderr = append(stderr, chunk); });
    child.stdin.once("error", (cause) => finish(() => rejectCommand(cause)));
    child.once("error", (cause) => finish(() => rejectCommand(cause)));
    child.once("close", (code, signal) => finish(() => {
      const output = { stdout: stdout.toString("utf8"), stderr: stderr.toString("utf8") };
      if (code === 0) resolveCommand(output);
      else {
        const detail = output.stderr.trim() || output.stdout.trim();
        rejectCommand(new Error(`${program} failed with ${signal === null ? `exit ${code}` : `signal ${signal}`}${detail === "" ? "" : `: ${detail}`}`));
      }
    }));
    timer = setTimeout(() => {
      child.kill("SIGTERM");
      finish(() => rejectCommand(new Error(`${program} timed out after ${timeoutMs} ms`)));
    }, timeoutMs);
    child.stdin.end(input);
  });
}

async function plistValue(applicationPath, key) {
  const plist = resolve(applicationPath, "Contents/Info.plist");
  const { stdout } = await runCommand("/usr/bin/plutil", ["-extract", key, "raw", "-o", "-", plist]);
  const value = stdout.trim();
  if (value === "") throw new Error(`${basename(applicationPath)} has no ${key}`);
  return value;
}

async function applicationMetadata(application) {
  const applicationStat = await stat(application.path).catch(() => undefined);
  if (applicationStat?.isDirectory() !== true) throw new Error(`${application.name} is not installed at ${application.path}`);
  const [bundleId, version, build] = await Promise.all([
    plistValue(application.path, "CFBundleIdentifier"),
    plistValue(application.path, "CFBundleShortVersionString"),
    plistValue(application.path, "CFBundleVersion"),
  ]);
  if (bundleId !== application.expectedBundleId) {
    throw new Error(`${application.name} bundle identifier is ${bundleId}; expected ${application.expectedBundleId}`);
  }
  return { name: application.name, bundleId, version, build };
}

function environmentMetadata() {
  const resolvedLocale = Intl.DateTimeFormat().resolvedOptions();
  return {
    platform: platform(),
    osRelease: release(),
    architecture: arch(),
    locale: resolvedLocale.locale,
    timezone: resolvedLocale.timeZone,
  };
}

async function pdftoppmVersion() {
  const { stdout, stderr } = await runCommand("pdftoppm", ["-v"]);
  const firstLine = `${stdout}\n${stderr}`.split(/\r?\n/u).find((line) => line.trim() !== "")?.trim();
  if (firstLine === undefined) throw new Error("pdftoppm did not report its version");
  return firstLine.replace(/^pdftoppm version\s+/u, "");
}

async function assertOutputAbsent(output) {
  try {
    await lstat(output);
  } catch (cause) {
    if (cause?.code === "ENOENT") return;
    throw cause;
  }
  throw new Error(`Output path already exists: ${output}`);
}

async function normalizePdfs(artifactDirectory, applicationKey) {
  const names = await readdir(artifactDirectory);
  if (applicationKey !== "excel") {
    if (!names.includes("reference.pdf")) throw new Error("Native application did not create reference.pdf");
    return [{ file: "reference.pdf", path: resolve(artifactDirectory, "reference.pdf") }];
  }
  const sheets = names.flatMap((file) => {
    const match = /^sheet-(\d+)\.pdf$/u.exec(file);
    return match === null ? [] : [{ file, index: Number(match[1]) }];
  }).sort((left, right) => left.index - right.index);
  if (sheets.length === 0) throw new Error("Microsoft Excel did not export any visible worksheet PDF");
  const result = [];
  for (const sheet of sheets) {
    const file = `sheet-${String(sheet.index).padStart(4, "0")}.pdf`;
    const currentPath = resolve(artifactDirectory, sheet.file);
    const normalizedPath = resolve(artifactDirectory, file);
    if (currentPath !== normalizedPath) await rename(currentPath, normalizedPath);
    result.push({ file, path: normalizedPath, sheetIndex: sheet.index - 1 });
  }
  return result;
}

function pngDimensions(bytes, label) {
  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
  if (bytes.length < 24 || !bytes.subarray(0, 8).equals(signature) || bytes.subarray(12, 16).toString("ascii") !== "IHDR") {
    throw new Error(`${label} is not a valid PNG`);
  }
  const width = bytes.readUInt32BE(16);
  const height = bytes.readUInt32BE(20);
  if (width === 0 || height === 0) throw new Error(`${label} has invalid dimensions`);
  return { width, height };
}

async function assertPdfSignature(file, label) {
  const handle = await openFile(file, "r");
  try {
    const actual = Buffer.alloc(5);
    const { bytesRead } = await handle.read(actual, 0, actual.length, 0);
    if (bytesRead !== actual.length || !actual.equals(Buffer.from("%PDF-", "ascii"))) {
      throw new Error(`${label} does not start with the required %PDF- signature`);
    }
  } finally {
    await handle.close();
  }
}

async function rasterizePdf(pdf, artifactDirectory, rasterDirectory) {
  const stem = basename(pdf.file, ".pdf");
  const prefix = resolve(rasterDirectory, stem);
  await runCommand("pdftoppm", [
    "-png",
    "-r", String(RASTER_DPI),
    "-cropbox",
    "-hide-annotations",
    "-freetype", "yes",
    "-aa", "yes",
    "-aaVector", "yes",
    "-thinlinemode", "none",
    "-forcenum",
    pdf.path,
    prefix,
  ], { timeoutMs: NATIVE_EXPORT_TIMEOUT_MS });
  const rawPages = (await readdir(rasterDirectory)).flatMap((file) => {
    const match = new RegExp(`^${stem.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&")}-(\\d+)\\.png$`, "u").exec(file);
    return match === null ? [] : [{ file, index: Number(match[1]) }];
  }).sort((left, right) => left.index - right.index);
  if (rawPages.length === 0) throw new Error(`pdftoppm produced no pages for ${pdf.file}`);
  const pages = [];
  for (const rawPage of rawPages) {
    const file = pdf.sheetIndex === undefined
      ? `page-${String(rawPage.index).padStart(4, "0")}.png`
      : `sheet-${String(pdf.sheetIndex + 1).padStart(4, "0")}-page-${String(rawPage.index).padStart(4, "0")}.png`;
    const outputPath = resolve(artifactDirectory, file);
    await rename(resolve(rasterDirectory, rawPage.file), outputPath);
    const bytes = await readFile(outputPath);
    pages.push({
      file,
      pageIndex: rawPage.index - 1,
      ...(pdf.sheetIndex === undefined ? {} : { sheetIndex: pdf.sheetIndex }),
      ...pngDimensions(bytes, file),
      bytes: bytes.length,
      sha256: createHash("sha256").update(bytes).digest("hex"),
    });
  }
  return pages;
}

async function importNativePdf({ fixture, output, nativePdf, extension, application, reviewedBy }) {
  if (platform() !== "darwin") throw new Error("Native reference metadata requires macOS");
  const [fixtureStat, reviewedPdfStat] = await Promise.all([
    stat(fixture).catch(() => undefined),
    stat(nativePdf).catch(() => undefined),
  ]);
  if (fixtureStat?.isFile() !== true) throw new Error(`Fixture must be an existing single file: ${fixture}`);
  if (reviewedPdfStat?.isFile() !== true || reviewedPdfStat.size === 0) {
    throw new Error(`Reviewed native PDF must be a non-empty file: ${nativePdf}`);
  }
  await assertPdfSignature(nativePdf, basename(nativePdf));
  await assertOutputAbsent(output);

  const [sourceHash, reviewedPdfHash, app, rasterizerVersion] = await Promise.all([
    sha256File(fixture),
    sha256File(nativePdf),
    applicationMetadata(application),
    pdftoppmVersion(),
  ]);
  const outputParent = dirname(output);
  await mkdir(outputParent, { recursive: true });
  const temporary = await mkdtemp(resolve(outputParent, ".native-reference-"));
  const artifactDirectory = resolve(temporary, "artifacts");
  const rasterDirectory = resolve(temporary, "raster");
  const outputPdf = resolve(artifactDirectory, "reference.pdf");
  let failure;
  let published = false;
  try {
    await Promise.all([mkdir(artifactDirectory), mkdir(rasterDirectory)]);
    await copyFile(nativePdf, outputPdf);
    const [copiedPdfStat, copiedPdfHash] = await Promise.all([
      stat(outputPdf),
      sha256File(outputPdf),
    ]);
    if (!copiedPdfStat.isFile() || copiedPdfStat.size === 0 || copiedPdfHash !== reviewedPdfHash) {
      throw new Error("Imported native PDF does not match the reviewed PDF");
    }
    await assertPdfSignature(outputPdf, "reference.pdf");
    const pages = await rasterizePdf(
      { file: "reference.pdf", path: outputPdf },
      artifactDirectory,
      rasterDirectory,
    );
    const [sourceHashAfterImport, reviewedPdfHashAfterImport] = await Promise.all([
      sha256File(fixture),
      sha256File(nativePdf),
    ]);
    if (sourceHashAfterImport !== sourceHash) throw new Error("Original fixture changed during native-PDF import");
    if (reviewedPdfHashAfterImport !== reviewedPdfHash) throw new Error("Reviewed native PDF changed during import");

    const generatedAt = new Date().toISOString();
    const manifest = {
      schemaVersion: 1,
      oracleSuite: application.oracleSuite,
      generatedAt,
      source: {
        file: basename(fixture),
        format: extension.slice(1),
        bytes: fixtureStat.size,
        sha256: sourceHash,
        integrityVerifiedAfterImport: true,
      },
      referenceApplication: {
        ...app,
        metadataScope: "currently installed application; this script did not verify which application instance produced the reviewed PDF",
      },
      environment: environmentMetadata(),
      captureSemantics: {
        unitType: application.captureSemantics.unitType,
        capture: "reviewed-native-pdf-import",
        description: `The user supplied this PDF as already exported and reviewed in ${application.name}.`,
        provenanceControl: "human-controlled",
        provenanceVerification: "not performed by this script",
        nativeApplicationAutomation: false,
        sourceIsolation: "the original fixture and reviewed PDF were read only; no native application was invoked",
        pdfPageBox: "cropbox",
        rasterDpi: RASTER_DPI,
        annotations: "hidden during rasterization",
      },
      rasterizer: {
        name: "pdftoppm",
        version: rasterizerVersion,
        dpi: RASTER_DPI,
        options: ["cropbox", "hide-annotations", "freetype=yes", "aa=yes", "aaVector=yes", "thinlinemode=none"],
      },
      pdfs: [{
        file: "reference.pdf",
        importedFrom: basename(nativePdf),
        bytes: copiedPdfStat.size,
        sha256: copiedPdfHash,
      }],
      reviewAttestation: {
        status: "reviewed",
        reviewer: reviewedBy,
        reviewedAt: generatedAt,
        scope: "all-rendered-units-visual-fidelity-v1",
        oracleSuite: application.oracleSuite,
        application: application.name,
        sourceSha256: sourceHash,
        artifactSha256: copiedPdfHash,
      },
      pages,
    };
    await writeFile(resolve(artifactDirectory, "reference.json"), `${JSON.stringify(manifest, null, 2)}\n`, { mode: 0o600 });
    const [finalSourceHash, finalReviewedPdfHash] = await Promise.all([
      sha256File(fixture),
      sha256File(nativePdf),
    ]);
    if (finalSourceHash !== sourceHash) throw new Error("Original fixture changed before native-PDF publication");
    if (finalReviewedPdfHash !== reviewedPdfHash) throw new Error("Reviewed native PDF changed before publication");
    await rename(artifactDirectory, output);
    published = true;
  } catch (cause) {
    failure = cause;
  } finally {
    try {
      await rm(temporary, { recursive: true, force: true });
    } catch (cause) {
      failure = failure === undefined ? cause : new AggregateError([failure, cause], "Native-PDF import and temporary cleanup both failed");
    }
  }
  if (failure !== undefined) throw failure;
  if (!published) throw new Error("Native-PDF import did not publish an output directory");
  console.log(output);
}

async function importSheetViewport({ fixture, output, sheetViewport, range, unitIndex, extension, application, reviewedBy }) {
  if (platform() !== "darwin") throw new Error("Native reference metadata requires macOS");
  const normalizedRange = normalizeSheetRange(range);
  if (extname(sheetViewport).toLowerCase() !== ".png") throw new Error("--sheet-viewport must reference a .png file");
  const [fixtureStat, viewportStat] = await Promise.all([
    stat(fixture).catch(() => undefined),
    stat(sheetViewport).catch(() => undefined),
  ]);
  if (fixtureStat?.isFile() !== true) throw new Error(`Fixture must be an existing single file: ${fixture}`);
  if (viewportStat?.isFile() !== true || viewportStat.size === 0) {
    throw new Error(`Reviewed sheet viewport must be a non-empty PNG file: ${sheetViewport}`);
  }
  await assertOutputAbsent(output);

  const [sourceHash, reviewedPng, app] = await Promise.all([
    sha256File(fixture),
    readFile(sheetViewport),
    applicationMetadata(application),
  ]);
  const dimensions = pngDimensions(reviewedPng, basename(sheetViewport));
  const viewportHash = createHash("sha256").update(reviewedPng).digest("hex");
  const outputParent = dirname(output);
  await mkdir(outputParent, { recursive: true });
  const temporary = await mkdtemp(resolve(outputParent, ".native-reference-"));
  const artifactDirectory = resolve(temporary, "artifacts");
  let failure;
  let published = false;
  try {
    await mkdir(artifactDirectory);
    const outputPng = resolve(artifactDirectory, "sheet-viewport.png");
    await copyFile(sheetViewport, outputPng);
    const copiedPng = await readFile(outputPng);
    if (createHash("sha256").update(copiedPng).digest("hex") !== viewportHash) {
      throw new Error("Imported sheet viewport does not match the reviewed PNG");
    }
    if (await sha256File(fixture) !== sourceHash) throw new Error("Original fixture changed during sheet-viewport import");
    const generatedAt = new Date().toISOString();
    const manifest = {
      schemaVersion: 1,
      oracleSuite: application.oracleSuite,
      generatedAt,
      source: {
        file: basename(fixture),
        format: extension.slice(1),
        bytes: fixtureStat.size,
        sha256: sourceHash,
        integrityVerifiedAfterImport: true,
      },
      referenceApplication: {
        ...app,
        metadataScope: "currently installed application; this script did not verify which application instance produced the reviewed PNG",
      },
      environment: environmentMetadata(),
      captureSemantics: {
        unitType: "sheet",
        unitIndex,
        capture: "sheet-viewport",
        range: normalizedRange,
        description: `A content-only ${application.name} spreadsheet viewport PNG reviewed outside this script and imported byte-for-byte.`,
        nativePdfExport: false,
        provenanceControl: "human-controlled",
        provenanceVerification: "not performed by this script",
        nativeApplicationAutomation: false,
        rasterizer: null,
      },
      viewport: {
        file: "sheet-viewport.png",
        unitIndex,
        range: normalizedRange,
        ...dimensions,
        bytes: reviewedPng.length,
        sha256: viewportHash,
      },
      reviewAttestation: {
        status: "reviewed",
        reviewer: reviewedBy,
        reviewedAt: generatedAt,
        scope: "all-rendered-units-visual-fidelity-v1",
        oracleSuite: application.oracleSuite,
        application: application.name,
        sourceSha256: sourceHash,
        artifactSha256: viewportHash,
      },
    };
    await writeFile(resolve(artifactDirectory, "reference.json"), `${JSON.stringify(manifest, null, 2)}\n`, { mode: 0o600 });
    await rename(artifactDirectory, output);
    published = true;
  } catch (cause) {
    failure = cause;
  } finally {
    try {
      await rm(temporary, { recursive: true, force: true });
    } catch (cause) {
      failure = failure === undefined ? cause : new AggregateError([failure, cause], "Sheet-viewport import and temporary cleanup both failed");
    }
    try {
      if (await sha256File(fixture) !== sourceHash) {
        const cause = new Error("Original fixture changed; final sheet-viewport integrity verification failed");
        failure = failure === undefined ? cause : new AggregateError([failure, cause], "Sheet-viewport import failed source-integrity verification");
      }
    } catch (cause) {
      failure = failure === undefined ? cause : new AggregateError([failure, cause], "Sheet-viewport import could not verify source integrity");
    }
  }
  if (failure !== undefined) throw failure;
  if (!published) throw new Error("Sheet-viewport import did not publish an output directory");
  console.log(output);
}

// Keep the Office-facing directory inode stable so its sandbox grant survives batches.
export async function acquireNativeWorkspace(directory) {
  await mkdir(directory, { recursive: true });
  const lockPath = resolve(directory, ".export-lock");
  const lock = await openFile(lockPath, "wx").catch((cause) => {
    if (cause.code === "EEXIST") throw new Error(`Native export workspace is busy: ${directory}`);
    throw cause;
  });
  return async () => {
    try {
      for (const file of await readdir(directory)) {
        if (/^(?:source\.[a-z0-9]+|reference\.pdf|sheet-\d+\.pdf)$/u.test(file)) {
          await rm(resolve(directory, file), { force: true });
        }
      }
    } finally {
      await lock.close();
      await rm(lockPath, { force: true });
    }
  };
}

// Persistent reference library; independent of disposable batch output and Office workspace.
const NATIVE_LIBRARY = resolve(dirname(fileURLToPath(import.meta.url)), "../reference-library/native");
export function nativeReferenceCacheKey(manifest) {
  // ponytail: environment metadata does not fingerprint fonts; refresh the policy when reference fonts change.
  return createHash("sha256").update(JSON.stringify([
    1, manifest.source.sha256, manifest.source.format, manifest.referenceApplication,
    manifest.environment, manifest.rasterizer.version, manifest.captureSemantics.allowPrintPages,
  ])).digest("hex");
}
async function verifiedCachedReference(directory) {
  const manifest = JSON.parse(await readFile(resolve(directory, "reference.json"), "utf8"));
  if (!manifest.source.integrityVerifiedAfterExport || !manifest.captureSemantics.sourceIdentityCheck
    || manifest.captureSemantics.capture !== "pdf-export" || !manifest.pages?.length || !manifest.pdfs?.length) {
    throw new Error("Reference lacks verified native provenance");
  }
  for (const file of [...manifest.pdfs, ...manifest.pages]) {
    if (basename(file.file) !== file.file || !/^(reference\.pdf|sheet-\d+\.pdf|(?:sheet-\d+-)?page-\d+\.png)$/u.test(file.file)
      || await sha256File(resolve(directory, file.file)) !== file.sha256) throw new Error("Reference cache integrity failure: " + file.file);
  }
  return manifest;
}
export async function cacheNativeReference(directory, library = NATIVE_LIBRARY) {
  const manifest = await verifiedCachedReference(directory);
  const key = nativeReferenceCacheKey(manifest), target = resolve(library, key);
  await mkdir(library, { recursive: true });
  try {
    const existing = await verifiedCachedReference(target);
    if (nativeReferenceCacheKey(existing) !== key) throw new Error("Reference cache key mismatch");
    return key;
  } catch (error) { if (error.code !== "ENOENT") throw error; }
  const staging = await mkdtemp(resolve(library, ".staging-"));
  try {
    await cp(directory, resolve(staging, "entry"), { recursive: true });
    await verifiedCachedReference(resolve(staging, "entry"));
    try { await rename(resolve(staging, "entry"), target); }
    catch (error) { if (!["EEXIST", "ENOTEMPTY"].includes(error.code)) throw error; await verifiedCachedReference(target); }
  } finally { await rm(staging, { recursive: true, force: true }); }
  return key;
}
export async function restoreNativeReference(key, output, library = NATIVE_LIBRARY) {
  if (!/^[a-f0-9]{64}$/u.test(key)) throw new Error("Invalid reference cache key");
  const directory = resolve(library, key);
  let manifest;
  try { manifest = await verifiedCachedReference(directory); }
  catch (error) { if (error.code === "ENOENT") return false; throw error; }
  if (nativeReferenceCacheKey(manifest) !== key) throw new Error("Reference cache key mismatch");
  await assertOutputAbsent(output);
  await mkdir(dirname(output), { recursive: true });
  const staging = await mkdtemp(resolve(dirname(output), ".cached-reference-"));
  try {
    await cp(directory, resolve(staging, "entry"), { recursive: true });
    await verifiedCachedReference(resolve(staging, "entry"));
    await rename(resolve(staging, "entry"), output);
  } finally { await rm(staging, { recursive: true, force: true }); }
  return true;
}

async function exportReference({ fixture, output, allowPrintPages, sheetViewport, nativePdf, range, unitIndex, reviewedBy }) {
  const extension = extname(fixture).toLowerCase();
  const applicationKey = EXTENSION_TO_APPLICATION.get(extension);
  if (applicationKey === undefined) {
    throw new Error(`Unsupported native reference format ${extension || "(none)"}; expected PPTX, PPT, ODP, DOCX, DOC, ODT, RTF, XLSX, XLS, ODS, KEY, PAGES, NUMBERS, WPS, ET, or DPS`);
  }
  const application = APPLICATIONS[applicationKey];
  if (nativePdf !== undefined) {
    if (application.printPages === true) {
      throw new Error(`--native-pdf is valid only for PPTX, PPT, ODP, DOCX, DOC, ODT, RTF, KEY, PAGES, WPS, or DPS, not ${extension}`);
    }
    await importNativePdf({ fixture, output, nativePdf, extension, application, reviewedBy });
    return;
  }
  if (sheetViewport !== undefined) {
    if (application.printPages !== true) {
      throw new Error(`--sheet-viewport is valid only for XLSX, XLS, ODS, NUMBERS, or ET, not ${extension}`);
    }
    await importSheetViewport({ fixture, output, sheetViewport, range, unitIndex, extension, application, reviewedBy });
    return;
  }
  if (application.manualOnly === true) {
    throw new Error(`Refusing automated ${extension} capture: WPS references require an explicitly reviewed ${application.printPages === true ? "--sheet-viewport" : "--native-pdf"} import`);
  }
  if (application.printPages === true && !allowPrintPages) {
    throw new Error(`Refusing ${extension}: native PDF export has print-page semantics, not a worksheet viewport. Define an explicit viewport oracle, or pass --allow-print-pages only for a reviewed single-page print-area/print-page fixture.`);
  }
  if (platform() !== "darwin") throw new Error("Native reference export requires macOS");
  const fixtureStat = await stat(fixture).catch(() => undefined);
  if (fixtureStat?.isFile() !== true) throw new Error(`Fixture must be an existing single file: ${fixture}`);
  const sourceBytes = await readFile(fixture);
  if (/^\.(doc|dot|ppt|pps|pot|xls|xlt)[xm]$/u.test(extension)
    && sourceBytes.subarray(0, 8).equals(Buffer.from("d0cf11e0a1b11ae1", "hex"))
    && sourceBytes.includes(Buffer.from("EncryptedPackage", "utf16le"))) {
    throw new Error("Encrypted Office package requires credentials; no unattended native reference was generated");
  }
  await assertOutputAbsent(output);

  const sourceHash = await sha256File(fixture);
  const [app, rasterizerVersion] = await Promise.all([
    applicationMetadata(application),
    pdftoppmVersion(),
  ]);
  const cacheKey = nativeReferenceCacheKey({ source: { sha256: sourceHash, format: extension.slice(1) },
    referenceApplication: app, environment: environmentMetadata(), rasterizer: { version: rasterizerVersion },
    captureSemantics: { allowPrintPages } });
  if (await restoreNativeReference(cacheKey, output)) {
    if (await sha256File(fixture) !== sourceHash) throw new Error("Original fixture changed during cache restore");
    console.error("Reused verified native reference: " + cacheKey);
    console.log(output);
    return;
  }
  const outputParent = dirname(output);
  await mkdir(outputParent, { recursive: true });
  const temporary = await mkdtemp(resolve(outputParent, ".native-reference-"));
  const artifactDirectory = resolve(temporary, "artifacts");
  const rasterDirectory = resolve(temporary, "raster");
  const nativeDirectory = resolve(dirname(fileURLToPath(import.meta.url)), "../.cache/native-reference", applicationKey);
  const copiedFixture = resolve(nativeDirectory, `source${extension}`);
  let releaseWorkspace;
  let failure;
  let published = false;
  try {
    await Promise.all([mkdir(artifactDirectory), mkdir(rasterDirectory)]);
    releaseWorkspace = await acquireNativeWorkspace(nativeDirectory);
    await copyFile(fixture, copiedFixture);
    if (await sha256File(copiedFixture) !== sourceHash) throw new Error("Temporary source copy does not match the original fixture");

    const scriptDestination = applicationKey === "excel"
      ? nativeDirectory
      : resolve(nativeDirectory, "reference.pdf");
    await runCommand("/usr/bin/osascript", ["-", copiedFixture, scriptDestination], {
      input: APPLE_SCRIPTS[applicationKey],
      timeoutMs: NATIVE_EXPORT_TIMEOUT_MS,
    });
    await rm(copiedFixture);

    const nativePdfs = await normalizePdfs(nativeDirectory, applicationKey);
    for (const pdf of nativePdfs) await copyFile(pdf.path, resolve(artifactDirectory, pdf.file));
    await releaseWorkspace();
    releaseWorkspace = undefined;
    const pdfs = await normalizePdfs(artifactDirectory, applicationKey);
    const pages = [];
    const pdfRecords = [];
    for (const pdf of pdfs) {
      const pdfStat = await stat(pdf.path);
      if (!pdfStat.isFile() || pdfStat.size === 0) throw new Error(`${pdf.file} is empty`);
      pdfRecords.push({
        file: pdf.file,
        ...(pdf.sheetIndex === undefined ? {} : { sheetIndex: pdf.sheetIndex }),
        bytes: pdfStat.size,
        sha256: await sha256File(pdf.path),
      });
      pages.push(...await rasterizePdf(pdf, artifactDirectory, rasterDirectory));
    }

    const sourceHashAfterExport = await sha256File(fixture);
    if (sourceHashAfterExport !== sourceHash) throw new Error("Original fixture changed during native reference export");
    const manifest = {
      schemaVersion: 1,
      oracleSuite: application.oracleSuite,
      generatedAt: new Date().toISOString(),
      source: {
        file: basename(fixture),
        format: extension.slice(1),
        bytes: fixtureStat.size,
        sha256: sourceHash,
        integrityVerifiedAfterExport: true,
      },
      referenceApplication: app,
      environment: environmentMetadata(),
      captureSemantics: {
        ...application.captureSemantics,
        capture: "pdf-export",
        nativeApplicationAutomation: true,
        sourceIdentityCheck: applicationKey === "powerpoint"
          ? (extension === ".potx" || extension === ".potm" ? "new-template-instance; manual-source-review-required" : "new-presentation-and-source-path")
          : applicationKey === "word" ? "opened-document-source-path"
            : applicationKey === "excel" ? "opened-workbook-source-path" : "application-returned-document-object",
        producerObservation: `This script invoked ${application.name} and observed its PDF export; this operational observation is not cryptographic producer verification.`,
        sourceIsolation: "native application opened a temporary byte-identical copy; the original path was never passed to it",
        pdfPageBox: "cropbox",
        rasterDpi: RASTER_DPI,
        annotations: "hidden during rasterization",
        allowPrintPages,
      },
      rasterizer: {
        name: "pdftoppm",
        version: rasterizerVersion,
        dpi: RASTER_DPI,
        options: ["cropbox", "hide-annotations", "freetype=yes", "aa=yes", "aaVector=yes", "thinlinemode=none"],
      },
      pdfs: pdfRecords,
      pages,
    };
    await writeFile(resolve(artifactDirectory, "reference.json"), `${JSON.stringify(manifest, null, 2)}\n`, { mode: 0o600 });
    await rename(artifactDirectory, output);
    published = true;
  } catch (cause) {
    failure = cause;
  } finally {
    try {
      await releaseWorkspace?.();
    } catch (cause) {
      failure = failure === undefined ? cause : new AggregateError([failure, cause], "Native workspace cleanup failed");
    }
    try {
      await rm(temporary, { recursive: true, force: true });
    } catch (cause) {
      failure = failure === undefined ? cause : new AggregateError([failure, cause], "Native reference export and temporary cleanup both failed");
    }
    try {
      if (await sha256File(fixture) !== sourceHash) {
        const cause = new Error("Original fixture changed; final integrity verification failed");
        failure = failure === undefined ? cause : new AggregateError([failure, cause], "Native reference export failed source-integrity verification");
      }
    } catch (cause) {
      failure = failure === undefined ? cause : new AggregateError([failure, cause], "Native reference export could not verify source integrity");
    }
  }
  if (failure !== undefined) throw failure;
  if (!published) throw new Error("Native reference export did not publish an output directory");
  await cacheNativeReference(output);
  console.log(output);
}

async function main() {
  const options = parseArguments(process.argv.slice(2));
  if (options.help) {
    console.log(usage());
    return;
  }
  await exportReference(options);
}

if (process.argv[1] !== undefined && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) main().catch((cause) => {
  console.error(cause instanceof Error ? cause.message : String(cause));
  process.exitCode = 1;
});
