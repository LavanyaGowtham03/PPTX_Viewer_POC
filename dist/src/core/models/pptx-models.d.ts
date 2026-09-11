/**
 * PPTX Viewer — Typed Data Models
 *
 * The single source of truth for the PPTX JSON format consumed by this
 * viewer. Derived from the actual sample JSON exported by the PPTX →
 * JSON converter.
 *
 * Hard rules encoded here:
 * - ALL coordinates (x, y, width, height) are in POINTS (px) — no EMU.
 * - width/height are ALWAYS read from the JSON — never assume 960×540.
 * - slides can contain ANY number of slides (1, 2, 4, 10+).
 * - Every optional field is marked "?" and must be null-safe at use sites.
 */
/** Root document as produced by the PPTX → JSON export. */
export interface PPTXFile {
    presentation: PptxPresentation;
}
export interface PptxPresentation {
    /** Slide width in points — ALWAYS read from JSON, never hardcoded. */
    width: number;
    /** Slide height in points — ALWAYS read from JSON, never hardcoded. */
    height: number;
    /** Any number of slides (1, 2, 4, 10+) — dynamic, no hardcoded limit. */
    slides: PptxSlide[];
}
export interface PptxSlide {
    /** e.g. "slide1" */
    id: string;
    /** e.g. "Slide 1" */
    name: string;
    /** Slide fill as a hex color string, e.g. "#FFFFFF" (parser defaults it). */
    background: string;
    /** Optional slide-level transition, played when navigating TO this slide. */
    transition?: SlideTransition;
    /** Any number of elements — renderers iterate all, skipping unknowns. */
    elements: PptxElement[];
}
/**
 * Fields common to every slide element.
 * All coordinates are in POINTS (px) — scale by slideScale at render time.
 */
export interface BaseElement {
    id: string;
    /**
     * Element kind. Known values: 'textbox' | 'shape' | 'picture' | 'table'
     * (lowercase in the JSON). Unknown values must be skipped gracefully by
     * the renderer with a console.warn.
     */
    type: string;
    x: number;
    y: number;
    width: number;
    height: number;
    /** Rotation in degrees; absent → no transform. */
    rotation?: number;
    /** Single animation object on the element (NOT an array). */
    animation?: ElementAnimation;
}
export interface TextBoxElement extends BaseElement {
    type: 'textbox';
    /** Flat text string; "\n" denotes line breaks. */
    text: string;
    style: ElementStyle;
}
export type ShapeType = 'oval' | 'triangle' | 'blockArrow' | 'flowchartTerminator' | 'calloutRectangle';
export interface ShapeElement extends BaseElement {
    type: 'shape';
    /** Geometry selector — separate field from `type`. */
    shapeType: ShapeType;
    /** Optional label centered inside the shape. */
    text?: string;
    style: ElementStyle;
    /** Shape fill as a hex color string, e.g. "#000000". */
    fill: string;
    /** Shape border color as a hex color string, e.g. "#000000". */
    stroke: string;
    /** Border width in px — defaults to 1 when absent. */
    strokeWidth?: number;
}
export interface PictureElement extends BaseElement {
    type: 'picture';
    /** base64 data URI ("data:image/...;base64,...") or an http(s) URL. */
    src: string;
    altText?: string;
}
export interface TableElement extends BaseElement {
    type: 'table';
    rows: TableRow[];
}
export interface TableRow {
    cells: TableCell[];
}
export interface TableCell {
    text?: string;
    rowSpan?: number;
    colSpan?: number;
    style?: ElementStyle;
    backgroundColor?: string;
    borderColor?: string;
}
/**
 * An element whose `type` (or a shape's `shapeType`) was not recognised.
 * The parser keeps it in the slide with `_unknown: true` so the renderer
 * can log a warning and skip it — never throwing, never hiding siblings.
 */
export interface UnknownElement extends BaseElement {
    /** The unrecognised `type` value exactly as it appeared in the JSON. */
    type: string;
    /** The raw `shapeType` when an unrecognised geometry was found. */
    shapeType?: string;
    /**
     * Raw shape colors preserved when an unknown shapeType is flagged —
     * ShapeRenderer uses them for the plain filled-rect fallback.
     */
    fill?: string;
    stroke?: string;
    strokeWidth?: number;
    /** Parser flag — renderers must skip these elements with a console.warn. */
    _unknown: true;
}
export type PptxElement = TextBoxElement | ShapeElement | PictureElement | TableElement | UnknownElement;
export interface ElementStyle {
    /** e.g. "Calibri", "Calibri Light" */
    fontFamily: string;
    /** Font size in POINTS, e.g. 44 → "44pt" */
    fontSize: number;
    /** Text color as a hex string, e.g. "#000000" */
    color: string;
    bold: boolean;
    italic: boolean;
    underline: boolean;
    strikethrough: boolean;
    alignment?: 'left' | 'center' | 'right' | 'justify';
}
export type AnimationType = 'fadeIn' | 'flyIn' | 'zoomIn' | 'bounceIn' | 'fadedSwivel';
export interface ElementAnimation {
    /** Unknown types fall back to fadeIn with a console.warn. */
    type: AnimationType;
    /** Seconds before the entry animation starts (e.g. 0). */
    delay: number;
    /** Seconds the entry animation runs (e.g. 1). */
    duration: number;
}
export type TransitionType = 'fade' | 'push' | 'wipe' | 'zoom';
export type TransitionDirection = 'left' | 'right' | 'up' | 'down';
export interface SlideTransition {
    /** Unknown types fall back to fade with a console.warn. */
    type: TransitionType;
    /** Total transition time in milliseconds (e.g. 400). */
    duration?: number;
    /** Applies to push and wipe only. */
    direction?: TransitionDirection;
}
