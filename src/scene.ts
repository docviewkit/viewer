import type { Diagnostic, DocumentInfo, DocumentObject, Rect } from "./types.js";
import type { FontStyle } from "./types.js";

export interface SceneAffineTransform {
  readonly a: number;
  readonly b: number;
  readonly c: number;
  readonly d: number;
  readonly e: number;
  readonly f: number;
}

export type SceneBlendMode =
  | "source-over"
  | "multiply"
  | "screen"
  | "overlay"
  | "darken"
  | "lighten"
  | "color-dodge"
  | "color-burn"
  | "hard-light"
  | "soft-light"
  | "difference"
  | "exclusion"
  | "hue"
  | "saturation"
  | "color"
  | "luminosity"
  | "destination-out";

export type ScenePathCommand =
  | { readonly kind: "moveTo"; readonly x: number; readonly y: number }
  | { readonly kind: "lineTo"; readonly x: number; readonly y: number }
  | {
      readonly kind: "quadraticCurveTo";
      readonly cpx: number;
      readonly cpy: number;
      readonly x: number;
      readonly y: number;
    }
  | {
      readonly kind: "bezierCurveTo";
      readonly cp1x: number;
      readonly cp1y: number;
      readonly cp2x: number;
      readonly cp2y: number;
      readonly x: number;
      readonly y: number;
    }
  | { readonly kind: "closePath" };

export type SceneGeometry =
  | "rectangle"
  | "ellipse"
  | "line"
  | {
      readonly kind: "rounded-rectangle";
      readonly radiusX: number;
      readonly radiusY: number;
    }
  | {
      readonly kind: "path";
      readonly fillRule: "nonzero" | "evenodd";
      readonly commands: readonly ScenePathCommand[];
    }
  | {
      readonly kind: "layered-path";
      readonly layers: readonly {
        readonly fillRule: "nonzero" | "evenodd";
        readonly fill: "normal" | "none" | "darken" | "darken-less" | "lighten" | "lighten-less";
        readonly stroke: boolean;
        readonly commands: readonly ScenePathCommand[];
      }[];
    };

export type SceneXpsColor =
  | { readonly kind: "rgba"; readonly color: number }
  | {
      readonly kind: "context";
      readonly alpha: number;
      readonly profile: Uint8Array;
      readonly channels: readonly number[];
    };

export type ScenePaint =
  | { readonly kind: "none" }
  | { readonly kind: "solid"; readonly color: number }
  | {
      readonly kind: "linear-gradient";
      readonly start: { readonly x: number; readonly y: number };
      readonly end: { readonly x: number; readonly y: number };
      readonly stops: readonly { readonly offset: number; readonly color: number }[];
    }
  | {
      readonly kind: "radial-gradient";
      readonly start: { readonly x: number; readonly y: number; readonly radius: number };
      readonly end: { readonly x: number; readonly y: number; readonly radius: number };
      readonly stops: readonly { readonly offset: number; readonly color: number }[];
    }
  | {
      readonly kind: "rect-gradient";
      readonly center: { readonly x: number; readonly y: number };
      readonly stops: readonly { readonly offset: number; readonly color: number }[];
    }
  | {
      readonly kind: "circle-gradient";
      readonly start: { readonly x: number; readonly y: number; readonly radius: number };
      readonly end: { readonly x: number; readonly y: number; readonly radius: number };
      readonly stops: readonly { readonly offset: number; readonly color: number }[];
    }
  | {
      readonly kind: "mapped-gradient";
      readonly paint: ScenePaint;
      readonly tile: { readonly left: number; readonly top: number; readonly right: number; readonly bottom: number };
      readonly flip: "none" | "tile" | "flip-x" | "flip-y" | "flip-xy";
      readonly rotateWithShape: boolean;
    }
  | {
      readonly kind: "shape-gradient";
      readonly focus: { readonly left: number; readonly top: number; readonly right: number; readonly bottom: number };
      readonly stops: readonly { readonly offset: number; readonly color: number }[];
    }
  | {
      readonly kind: "pattern";
      readonly preset: string;
      readonly foreground: number;
      readonly background: number;
    }
  | {
      readonly kind: "image";
      readonly mediaType: string;
      readonly bytes: Uint8Array;
      readonly cropLeft: number;
      readonly cropTop: number;
      readonly cropRight: number;
      readonly cropBottom: number;
      readonly tile: boolean;
      readonly tileWidth?: number;
      readonly tileHeight?: number;
      readonly mapping?: {
        readonly scaleX: number; readonly scaleY: number;
        readonly offsetX: number; readonly offsetY: number;
        readonly alignmentX: number; readonly alignmentY: number;
        readonly left: number; readonly top: number; readonly right: number; readonly bottom: number;
        readonly dpi: number; readonly flip: "none" | "tile" | "flip-x" | "flip-y" | "flip-xy";
        readonly rotateWithShape: boolean;
      };
    }
  | {
      readonly kind: "visual";
      readonly viewbox: { readonly x: number; readonly y: number; readonly width: number; readonly height: number };
      readonly viewport: { readonly x: number; readonly y: number; readonly width: number; readonly height: number };
      readonly viewboxRelative: boolean;
      readonly viewportRelative: boolean;
      readonly tileMode: "none" | "tile" | "flip-x" | "flip-y" | "flip-xy";
      readonly stretch: "none" | "fill" | "uniform" | "uniform-to-fill";
      readonly alignmentX: number;
      readonly alignmentY: number;
      readonly transform: SceneAffineTransform;
      readonly relativeTransform: SceneAffineTransform;
      readonly opacity: number;
      readonly children: readonly {
        readonly bounds: { readonly x: number; readonly y: number; readonly width: number; readonly height: number };
        readonly visual: SceneVisual;
      }[];
    }
  | {
      readonly kind: "xps-gradient";
      readonly radial: boolean;
      readonly start: { readonly x: number; readonly y: number };
      readonly end: { readonly x: number; readonly y: number };
      readonly radiusX: number;
      readonly radiusY: number;
      readonly relative: boolean;
      readonly spread: "pad" | "reflect" | "repeat";
      readonly linearRgb: boolean;
      readonly transform: SceneAffineTransform;
      readonly relativeTransform: SceneAffineTransform;
      readonly stops: readonly { readonly offset: number; readonly color: SceneXpsColor }[];
    };

export interface SceneShadow {
  readonly color: number;
  readonly blur: number;
  readonly offsetX: number;
  readonly offsetY: number;
}

export interface SceneOuterShadow extends SceneShadow {
  readonly scaleX: number;
  readonly scaleY: number;
  readonly skewX: number;
  readonly skewY: number;
  readonly alignment: number;
}

export interface SceneGlow {
  readonly color: number;
  readonly radius: number;
}

export interface SceneReflection {
  readonly startOpacity: number;
  readonly endOpacity: number;
  readonly startPosition: number;
  readonly endPosition: number;
  readonly directionDegrees: number;
  readonly blur: number;
  readonly distance: number;
  readonly scaleX: number;
  readonly scaleY: number;
}

export interface SceneStrokeStyle {
  readonly cap: "flat" | "round" | "square";
  readonly join: "miter" | "round" | "bevel";
  readonly compound: "single" | "double" | "thick-thin" | "thin-thick" | "triple";
  readonly alignment: "center" | "inset";
  readonly miterLimit: number;
  readonly dashOffset?: number;
  readonly dash: readonly number[];
}

export interface SceneTextLayout {
  readonly fillCharacter?: { readonly offset: number; readonly character: string };
  /** Rectangular exclusions relative to the inset text origin. */
  readonly wrapRegions?: readonly Rect[];
  /** Remove full-width punctuation whitespace only when needed to fit a line. */
  readonly compressPunctuation?: boolean;
  /** Center font boxes within fixed-height lines; bottom-align oversized fonts. */
  readonly fixedLineHeight?: boolean;
  /** The last visible paragraph continues in another frame. */
  readonly continuesAfter?: boolean;
  readonly direction: "auto" | "ltr" | "rtl";
  readonly orientation: "horizontal" | "vertical-rl" | "vertical-lr" | "rotated-90" | "rotated-270" | "stacked-rl" | "stacked-lr";
  readonly autoFit: "none" | "shrink" | "fit-frame";
  readonly verticalAlign: "top" | "center" | "bottom";
  readonly prefix?: string;
  readonly tabStops: readonly SceneTextTabStop[];
  readonly defaultTabStop: number;
  readonly hangingIndent: number;
  readonly paragraphSpacing?: number;
  readonly insetLeft: number;
  readonly insetRight: number;
  readonly insetTop: number;
  readonly insetBottom: number;
  readonly marginLeft: number;
  readonly marginRight: number;
  readonly firstLineIndent: number;
  readonly columnCount: number;
  readonly columnSpacing: number;
  readonly rotationDegrees: number;
  readonly fontScale: number;
  readonly lineSpacingReduction: number;
  readonly horizontalOverflow: "overflow" | "clip";
  readonly verticalOverflow: "overflow" | "clip" | "ellipsis";
  readonly wrap: boolean;
  readonly warp?: string;
  readonly textFill?: boolean;
  readonly textPaint?: ScenePaint;
  readonly textScaleToFit?: boolean;
  readonly textMatrixScaleToFit?: boolean;
  readonly lowResolutionSupersample?: boolean;
  readonly textStrokeColor?: number;
  readonly textStrokePaint?: ScenePaint;
  readonly textStrokeWidth?: number;
  readonly textBaseline?: number;
  readonly paragraphs?: readonly SceneTextParagraphLayout[];
  readonly minScale: number;
}

export interface SceneTextTabStop {
  readonly position: number;
  readonly align: "start" | "center" | "end";
  readonly leader: "none" | "dot" | "hyphen" | "underscore" | "middle-dot";
}

export type SceneTextAlign =
  | "start"
  | "center"
  | "end"
  | "justify"
  | "distribute"
  | "medium-kashida"
  | "high-kashida"
  | "low-kashida"
  | "thai-distribute";

export interface SceneTextParagraphLayout {
  readonly align: SceneTextAlign;
  readonly marginLeft: number;
  readonly marginRight: number;
  readonly firstLineIndent: number;
  readonly defaultTabStop: number;
  /** Absolute paragraph line height; zero inherits the text box line height. */
  readonly lineHeight?: number;
  readonly spaceBefore?: number;
  readonly spaceAfter?: number;
  /** Allows oversized Latin words to split at grapheme boundaries. */
  readonly latinLineBreak?: boolean;
  /** Allows terminal punctuation to extend beyond the paragraph's right edge. */
  readonly hangingPunctuation?: boolean;
  readonly ruleAbove?: {
    readonly color: number;
    readonly strokeWidth: number;
    readonly offsetX: number;
    readonly offsetY: number;
    readonly width: number;
  };
  readonly ruleBelow?: SceneTextParagraphLayout["ruleAbove"];
  readonly dropCap?: {
    readonly characters: number;
    readonly lines: number;
    readonly raisedLines: number;
    readonly padding: number;
    readonly outdent: number;
  };
}

export interface SceneTextRun {
  readonly paint?: ScenePaint;
  /** False disables East Asian punctuation restrictions for this run. */
  readonly eastAsianLineBreaks?: boolean;
  readonly text: string;
  /** Internal UTF-16 range in the owning object's logical text. */
  readonly sourceStart?: number;
  readonly sourceEnd?: number;
  readonly fontFamily: string;
  readonly fontSize: number;
  readonly color: number;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: boolean;
  readonly strikethrough: boolean;
  /** RGBA highlight; an alpha byte of zero means no highlight. */
  readonly highlight: number;
  /** Positive values raise the run, negative values lower it. */
  readonly baselineShift: number;
  readonly letterSpacing: number;
  /** Authored glyph width multiplier; 1 preserves the font's natural width. */
  readonly horizontalScale?: number;
  readonly shadow?: SceneShadow;
  readonly innerShadow?: SceneShadow;
  readonly textEffect?: SceneTextEffect;
}

export interface SceneTextEffect {
  readonly stroke?: ScenePaint;
  readonly strokeWidth?: number;
  readonly glow?: { readonly color: number; readonly radius: number };
  readonly fillToText?: boolean;
  readonly shadow?: SceneShadow;
  readonly innerShadow?: SceneShadow;
  readonly reflection?: SceneReflection;
  readonly wavyUnderline: boolean;
  readonly dottedUnderline: boolean;
  readonly heavyUnderline: boolean;
  readonly doubleUnderline: boolean;
  readonly dotDashUnderline: boolean;
  readonly doubleStrikethrough: boolean;
  readonly shadowScaleX: number;
  readonly shadowScaleY: number;
  readonly shadowSkewX: number;
  readonly shadowSkewY: number;
  readonly shadowAlignment: number;
}

export interface SceneBevel3D {
  readonly width: number;
  readonly height: number;
  readonly preset: string;
}

export interface SceneBackdrop3D {
  readonly anchorX: number;
  readonly anchorY: number;
  readonly anchorZ: number;
  readonly normalX: number;
  readonly normalY: number;
  readonly normalZ: number;
  readonly upX: number;
  readonly upY: number;
  readonly upZ: number;
}

export interface SceneThreeDStyle {
  readonly cameraPreset: string;
  readonly cameraFov: number;
  readonly cameraZoom: number;
  /** Resolved angles; the host parser already applied the camera preset defaults. */
  readonly cameraLatitude: number;
  readonly cameraLongitude: number;
  readonly cameraRevolution: number;
  readonly lightRig: string;
  readonly lightDirection: string;
  readonly lightLatitude: number;
  readonly lightLongitude: number;
  readonly lightRevolution: number;
  readonly z: number;
  readonly extrusionHeight: number;
  readonly contourWidth: number;
  readonly material: string;
  readonly bevelTop?: SceneBevel3D;
  readonly bevelBottom?: SceneBevel3D;
  readonly extrusionColor?: number;
  readonly contourColor?: number;
  readonly backdrop?: SceneBackdrop3D;
  readonly flatTextZ?: number;
  readonly appliesToText?: boolean;
}

export type SceneVisual =
  | { readonly kind: "none" }
  | {
      readonly kind: "shape";
      readonly geometry: SceneGeometry;
      readonly fill: number;
      readonly stroke: number;
      readonly strokeWidth: number;
    }
  | {
      readonly kind: "text";
      readonly geometry: SceneGeometry;
      readonly fill: number;
      readonly stroke: number;
      readonly strokeWidth: number;
      readonly fontFamily: string;
      readonly fontSize: number;
      readonly color: number;
      readonly bold: boolean;
      readonly italic: boolean;
      readonly align: SceneTextAlign;
    }
  | {
      readonly kind: "image";
      readonly mediaType: string;
      readonly bytes: Uint8Array;
      readonly fallback?: {
        readonly mediaType: string;
        readonly bytes: Uint8Array;
      };
      readonly alphaMask?: {
        readonly width: number;
        readonly height: number;
        readonly mediaType: string;
        readonly bytes: Uint8Array;
      };
      readonly cropLeft: number;
      readonly cropTop: number;
      readonly cropRight: number;
      readonly cropBottom: number;
    }
  | {
      readonly kind: "media";
      readonly mediaKind: "audio" | "video";
      readonly mediaType: string;
      readonly bytes: Uint8Array;
      readonly visual: SceneVisual;
    }
  | {
      readonly kind: "image-color-change";
      readonly from: number;
      readonly to: number;
      readonly useAlpha: boolean;
      readonly visual: SceneVisual;
    }
  | {
      readonly kind: "image-adjustment";
      readonly grayscale: boolean;
      readonly bilevelThreshold?: number;
      readonly brightness: number;
      readonly contrast: number;
      readonly duotone?: readonly [number, number];
      readonly visual: SceneVisual;
    }
  | {
      readonly kind: "group";
      readonly children: readonly { readonly bounds: Rect; readonly visual: SceneVisual }[];
    }
  | {
      readonly kind: "opacity-mask";
      readonly mask: ScenePaint;
      readonly visual: SceneVisual;
    }
  | {
      readonly kind: "color-managed-image";
      readonly sourceProfile: Uint8Array;
      readonly destinationProfile?: Uint8Array;
      readonly visual: SceneVisual;
    }
  | {
      readonly kind: "layer";
      readonly transform: SceneAffineTransform;
      readonly opacity: number;
      readonly blendMode?: SceneBlendMode;
      readonly visual: SceneVisual;
    }
  | {
      readonly kind: "painted-shape";
      readonly geometry: SceneGeometry;
      readonly fill: ScenePaint;
      readonly stroke: ScenePaint;
      readonly strokeWidth: number;
    }
  | {
      readonly kind: "rich-text";
      readonly geometry: SceneGeometry;
      readonly fill: ScenePaint;
      readonly stroke: ScenePaint;
      readonly strokeWidth: number;
      readonly align: SceneTextAlign;
      readonly lineHeight: number;
      readonly runs: readonly SceneTextRun[];
    }
  | {
      readonly kind: "effect";
      readonly shadow?: SceneShadow;
      readonly clip?: SceneGeometry;
      readonly visual: SceneVisual;
    }
  | {
      readonly kind: "text-layout";
      readonly layout: SceneTextLayout;
      readonly visual: SceneVisual;
    }
  | {
      readonly kind: "text-effects";
      readonly effects: readonly SceneTextEffect[];
      readonly visual: SceneVisual;
    }
  | {
      readonly kind: "stroke-style";
      readonly style: SceneStrokeStyle;
      readonly visual: SceneVisual;
    }
  | {
      readonly kind: "advanced-effect";
      readonly outerShadow?: SceneOuterShadow;
      readonly innerShadow?: SceneShadow;
      readonly glow?: SceneGlow;
      readonly reflection?: SceneReflection;
      readonly softEdge?: number;
      readonly threeD?: SceneThreeDStyle;
      readonly visual: SceneVisual;
    };

export interface SceneObject extends DocumentObject {
  readonly numericId: number;
  readonly parentNumericId?: number;
  readonly z: number;
  readonly visual: SceneVisual;
}

const IDENTITY_TRANSFORM: SceneAffineTransform = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };

export function multiplyTransform(
  left: SceneAffineTransform,
  right: SceneAffineTransform,
): SceneAffineTransform {
  return {
    a: left.a * right.a + left.c * right.b,
    b: left.b * right.a + left.d * right.b,
    c: left.a * right.c + left.c * right.d,
    d: left.b * right.c + left.d * right.d,
    e: left.a * right.e + left.c * right.f + left.e,
    f: left.b * right.e + left.d * right.f + left.f,
  };
}

export function nestedFontVisual(visual: SceneVisual): SceneVisual | undefined {
  switch (visual.kind) {
    case "layer":
    case "effect":
    case "text-layout":
    case "text-effects":
    case "stroke-style":
    case "advanced-effect":
    case "image-color-change":
    case "image-adjustment":
    case "media":
      return visual.visual;
    default:
      return undefined;
  }
}

export function sceneTextLayout(visual: SceneVisual): SceneTextLayout | undefined {
  for (let depth = 0; depth <= 64; depth += 1) {
    if (visual.kind === "text-layout") return visual.layout;
    const nested = nestedFontVisual(visual);
    if (nested === undefined) return undefined;
    visual = nested;
  }
  return undefined;
}

function leadingTransform(visual: SceneVisual): SceneAffineTransform {
  let transform = IDENTITY_TRANSFORM;
  let current = visual;
  for (let depth = 0; depth <= 64; depth += 1) {
    if (current.kind === "layer") {
      transform = multiplyTransform(transform, current.transform);
      current = current.visual;
    } else if (current.kind === "effect"
      || current.kind === "text-layout"
      || current.kind === "text-effects"
      || current.kind === "stroke-style"
      || current.kind === "advanced-effect"
      || current.kind === "image-color-change"
      || current.kind === "image-adjustment"
      || current.kind === "media") {
      current = current.visual;
    } else {
      break;
    }
  }
  return transform;
}

/** Public object bounds are document-space AABBs, including ancestor-group transforms. */
export function documentSpaceBounds(
  object: SceneObject,
  objectsByNumericId: ReadonlyMap<number, SceneObject>,
  includeObjectTransform = false,
): Rect {
  const ancestors: SceneObject[] = [];
  const visited = new Set<number>([object.numericId]);
  let parentId = object.parentNumericId;
  while (parentId !== undefined && ancestors.length < 128 && !visited.has(parentId)) {
    visited.add(parentId);
    const parent = objectsByNumericId.get(parentId);
    if (parent === undefined) break;
    ancestors.unshift(parent);
    parentId = parent.parentNumericId;
  }
  let transform = IDENTITY_TRANSFORM;
  for (const ancestor of ancestors) {
    if (ancestor.type === "group") {
      transform = multiplyTransform(transform, leadingTransform(ancestor.visual));
    }
  }
  if (includeObjectTransform) transform = multiplyTransform(transform, leadingTransform(object.visual));
  const { x, y, width, height } = object.bounds;
  const corners = [
    [x, y],
    [x + width, y],
    [x, y + height],
    [x + width, y + height],
  ].map(([pointX = 0, pointY = 0]) => ({
    x: transform.a * pointX + transform.c * pointY + transform.e,
    y: transform.b * pointX + transform.d * pointY + transform.f,
  }));
  const xs = corners.map((corner) => corner.x);
  const ys = corners.map((corner) => corner.y);
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const minY = Math.min(...ys);
  const maxY = Math.max(...ys);
  return { x: minX, y: minY, width: maxX - minX, height: maxY - minY };
}

export interface SceneEmbeddedFont {
  readonly family: string;
  readonly bytes: Uint8Array;
  readonly style: FontStyle;
  readonly weight: number;
}

export interface SceneDocument {
  readonly fatal: boolean;
  readonly info?: DocumentInfo;
  readonly diagnostics: readonly Diagnostic[];
  readonly embeddedFonts: readonly SceneEmbeddedFont[];
  readonly fontAlternateNames?: readonly { readonly family: string; readonly names: readonly string[] }[];
  readonly objects: readonly SceneObject[];
}
