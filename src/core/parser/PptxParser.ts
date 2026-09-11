import type {
  AnimationType,
  BaseElement,
  ElementAnimation,
  ElementStyle,
  PictureElement,
  PptxElement,
  PptxPresentation,
  PptxSlide,
  ShapeElement,
  ShapeType,
  SlideTransition,
  TableElement,
  TableCell,
  TextBoxElement,
  TableRow,
  TransitionDirection,
  TransitionType,
  UnknownElement,
} from '../models/pptx-models';

/**
 * PptxParser — raw JSON text → typed, normalised PptxPresentation.
 *
 * Error philosophy:
 * - FATAL (throw, with a clear message): the file isn't a presentation
 *   at all — bad JSON syntax, or missing presentation.width/height/slides.
 * - SAFE (never throw): every optional field gets a default; unknown
 *   element types / shapeTypes are kept but flagged `_unknown: true` so
 *   the renderer can skip them with a console.warn.
 *
 * The slide count is ALWAYS dynamic — 1, 2, 4, 10+ slides all parse
 * identically; nothing here assumes a count or hardcoded dimensions.
 */

/** Shared message for any file that isn't a valid PPTX JSON export. */
export const INVALID_FILE_MESSAGE =
  'Invalid file. Please select a valid PPTX JSON file.';

// ─────────────────────────────────────────────────────────────────────────────
// Small unknown-narrowing helpers (no `any`)
// ─────────────────────────────────────────────────────────────────────────────

const asRecord = (value: unknown): Record<string, unknown> | null =>
  typeof value === 'object' && value !== null
    ? (value as Record<string, unknown>)
    : null;

const asNumber = (value: unknown): number | null =>
  typeof value === 'number' && Number.isFinite(value) ? value : null;

const asString = (value: unknown): string | null =>
  typeof value === 'string' ? value : null;

const asBoolean = (value: unknown, fallback: boolean): boolean =>
  typeof value === 'boolean' ? value : fallback;

const asArray = (value: unknown): unknown[] | null =>
  Array.isArray(value) ? value : null;

// ─────────────────────────────────────────────────────────────────────────────
// Style / animation / transition normalisation (all safe — never throw)
// ─────────────────────────────────────────────────────────────────────────────

const KNOWN_ANIMATIONS: ReadonlySet<string> = new Set([
  'fadeIn',
  'flyIn',
  'zoomIn',
  'bounceIn',
  'fadedSwivel',
]);

const KNOWN_TRANSITIONS: ReadonlySet<string> = new Set(['fade', 'push', 'wipe', 'zoom']);

const KNOWN_DIRECTIONS: ReadonlySet<string> = new Set(['left', 'right', 'up', 'down']);

const KNOWN_ALIGNMENTS: ReadonlySet<string> = new Set([
  'left',
  'center',
  'right',
  'justify',
]);

/** Style defaults: bold/italic/underline/strikethrough → false; alignment → left. */
const normalizeStyle = (raw: unknown): ElementStyle => {
  const rec = asRecord(raw);

  if (!rec) {
    // No style block at all — fully safe defaults.
    return {
      fontFamily: 'Calibri',
      fontSize: 18,
      color: '#000000',
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      alignment: 'left',
    };
  }

  const alignmentRaw = asString(rec['alignment']);
  const alignment = alignmentRaw !== null && KNOWN_ALIGNMENTS.has(alignmentRaw)
    ? alignmentRaw
    : 'left';

  return {
    fontFamily: asString(rec['fontFamily']) ?? 'Calibri',
    fontSize: asNumber(rec['fontSize']) ?? 18,
    color: asString(rec['color']) ?? '#000000',
    bold: asBoolean(rec['bold'], false),
    italic: asBoolean(rec['italic'], false),
    underline: asBoolean(rec['underline'], false),
    strikethrough: asBoolean(rec['strikethrough'], false),
    alignment: alignment as ElementStyle['alignment'],
  };
};

/** Animation is a single object; unknown types fall back to fadeIn. */
const normalizeAnimation = (raw: unknown): ElementAnimation | undefined => {
  const rec = asRecord(raw);

  if (!rec) {
    return undefined; // no animation — element simply renders visible
  }

  const typeRaw = asString(rec['type']);
  const type: AnimationType =
    typeRaw !== null && KNOWN_ANIMATIONS.has(typeRaw) ? (typeRaw as AnimationType) : 'fadeIn';

  if (typeRaw !== null && typeRaw !== type) {
    console.warn(
      `[PptxViewer] Unknown animation type "${typeRaw}" — falling back to fadeIn.`
    );
  }

  return {
    type,
    delay: asNumber(rec['delay']) ?? 0,
    duration: asNumber(rec['duration']) ?? 1,
  };
};

/** Transition duration is in ms; direction applies to push/wipe. */
const normalizeTransition = (raw: unknown): SlideTransition | undefined => {
  const rec = asRecord(raw);

  if (!rec) {
    return undefined;
  }

  const typeRaw = asString(rec['type']);
  const type: TransitionType =
    typeRaw !== null && KNOWN_TRANSITIONS.has(typeRaw) ? (typeRaw as TransitionType) : 'fade';

  if (typeRaw !== null && typeRaw !== type) {
    console.warn(
      `[PptxViewer] Unknown transition type "${typeRaw}" — falling back to fade.`
    );
  }

  const directionRaw = asString(rec['direction']);
  const direction: TransitionDirection | undefined =
    directionRaw !== null && KNOWN_DIRECTIONS.has(directionRaw)
      ? (directionRaw as TransitionDirection)
      : undefined;

  const duration = asNumber(rec['duration']);

  return {
    type,
    ...(duration !== null ? { duration } : {}),
    ...(direction !== undefined ? { direction } : {}),
  };
};

// ─────────────────────────────────────────────────────────────────────────────
// Element normalisation — one branch per known type, plus the unknown flag
// ─────────────────────────────────────────────────────────────────────────────

const KNOWN_SHAPE_TYPES: ReadonlySet<string> = new Set([
  'oval',
  'triangle',
  'blockArrow',
  'flowchartTerminator',
  'calloutRectangle',
]);

/** Extract the base fields (position/size/rotation/animation) safely. */
const normalizeBase = (rec: Record<string, unknown>): BaseElement => {
  const rotation = asNumber(rec['rotation']);

  return {
    id: asString(rec['id']) ?? 'element', // must exist for warn messages
    type: asString(rec['type']) ?? '',
    x: asNumber(rec['x']) ?? 0,
    y: asNumber(rec['y']) ?? 0,
    width: asNumber(rec['width']) ?? 0,
    height: asNumber(rec['height']) ?? 0,
    ...(rotation !== null && rotation !== 0 ? { rotation } : {}),
    animation: normalizeAnimation(rec['animation']),
  };
};

const normalizeTextBox = (rec: Record<string, unknown>): TextBoxElement => ({
  ...normalizeBase(rec),
  type: 'textbox',
  text: asString(rec['text']) ?? '', // absent text → empty content, no crash
  style: normalizeStyle(rec['style']),
});

const normalizeShape = (rec: Record<string, unknown>): ShapeElement | UnknownElement => {
  const shapeTypeRaw = asString(rec['shapeType']);

  // Unknown/missing shapeType → flag as unknown, PRESERVING the raw shape
  // colors (fill/stroke/strokeWidth) so ShapeRenderer can draw its
  // plain filled-rect fallback without throwing.
  if (shapeTypeRaw === null || !KNOWN_SHAPE_TYPES.has(shapeTypeRaw)) {
    const fill = asString(rec['fill']);
    const stroke = asString(rec['stroke']);
    const strokeWidth = asNumber(rec['strokeWidth']);

    return {
      ...normalizeBase(rec),
      ...(shapeTypeRaw !== null ? { shapeType: shapeTypeRaw } : {}),
      ...(fill !== null ? { fill } : {}),
      ...(stroke !== null ? { stroke } : {}),
      ...(strokeWidth !== null && strokeWidth > 0 ? { strokeWidth } : {}),
      _unknown: true,
    };
  }

  const strokeWidth = asNumber(rec['strokeWidth']);
  const text = asString(rec['text']);

  return {
    ...normalizeBase(rec),
    type: 'shape',
    shapeType: shapeTypeRaw as ShapeType,
    ...(text !== null ? { text } : {}),
    style: normalizeStyle(rec['style']),
    fill: asString(rec['fill']) ?? '#000000',
    stroke: asString(rec['stroke']) ?? '#000000',
    strokeWidth: strokeWidth !== null && strokeWidth > 0 ? strokeWidth : 1,
  };
};

/** Build a flagged UnknownElement that renders will skip with a warning. */
const makeUnknown = (base: BaseElement, shapeTypeRaw: string | null): UnknownElement => ({
  ...base,
  ...(shapeTypeRaw !== null ? { shapeType: shapeTypeRaw } : {}),
  _unknown: true,
});

const normalizePicture = (
  rec: Record<string, unknown>
): PictureElement => {
  const altText = asString(rec['altText']);
  return {
    ...normalizeBase(rec),
    type: 'picture',
    src: asString(rec['src']) ?? '', // empty src → placeholder, no crash
    ...(altText !== null ? { altText } : {}),
  };
};

const normalizeTableCell = (raw: unknown): TableCell => {
  const rec = asRecord(raw) ?? {};

  const rowSpan = asNumber(rec['rowSpan']);
  const colSpan = asNumber(rec['colSpan']);
  const style = normalizeStyle(rec['style']);
  const backgroundColor = asString(rec['backgroundColor']);
  const borderColor = asString(rec['borderColor']);

  return {
    ...(asString(rec['text']) !== null ? { text: asString(rec['text']) as string } : {}),
    ...(rowSpan !== null && rowSpan > 1 ? { rowSpan } : {}),
    ...(colSpan !== null && colSpan > 1 ? { colSpan } : {}),
    ...(rec['style'] !== undefined ? { style } : {}),
    ...(backgroundColor !== null ? { backgroundColor } : {}),
    ...(borderColor !== null ? { borderColor } : {}),
  };
};

const normalizeTable = (rec: Record<string, unknown>): TableElement => {
  const rawRows = asArray(rec['rows']) ?? []; // missing rows → empty table
  const rows: TableRow[] = rawRows.map((rawRow): TableRow => {
    const rowRec = asRecord(rawRow) ?? {};
    const rawCells = asArray(rowRec['cells']) ?? [];
    return { cells: rawCells.map(normalizeTableCell) };
  });

  return {
    ...normalizeBase(rec),
    type: 'table',
    rows,
  };
};

/**
 * Central element dispatch. Unknown types are KEPT as UnknownElement
 * (`_unknown: true`) — never dropped here, never thrown.
 */
const normalizeElement = (raw: unknown): PptxElement => {
  const rec = asRecord(raw);

  const type = rec !== null ? asString(rec['type']) : null;

  if (rec !== null && type === 'textbox') return normalizeTextBox(rec);
  if (rec !== null && type === 'shape') return normalizeShape(rec);
  if (rec !== null && type === 'picture') return normalizePicture(rec);
  if (rec !== null && type === 'table') return normalizeTable(rec);

  // Unknown/missing type, or the entry wasn't even an object —
  // flag it; the renderer warns + skips. element.id preserved for the warn.
  return makeUnknown(
    rec !== null ? normalizeBase(rec) : emptyBase(),
    asString(rec?.['shapeType'])
  );
};

/** Minimal placeholder used only when the JSON entry wasn't an object. */
const emptyBase = (): BaseElement => ({
  id: 'element',
  type: '',
  x: 0,
  y: 0,
  width: 0,
  height: 0,
  animation: undefined,
});

// ─────────────────────────────────────────────────────────────────────────────
// Slide normalisation
// ─────────────────────────────────────────────────────────────────────────────

const normalizeSlide = (raw: unknown, index: number): PptxSlide => {
  const rec = asRecord(raw) ?? {};

  return {
    id: asString(rec['id']) ?? `slide${index + 1}`,
    name: asString(rec['name']) ?? `Slide ${index + 1}`,
    background: asString(rec['background']) ?? '#FFFFFF', // safe default
    transition: normalizeTransition(rec['transition']),
    elements: (asArray(rec['elements']) ?? []).map(normalizeElement),
  };
};

// ─────────────────────────────────────────────────────────────────────────────
// Public API
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Parse the raw JSON text of a PPTX JSON export.
 *
 * @throws Error (with a user-presentable message) only for FATAL issues:
 *         broken JSON syntax, or missing presentation.width/height/slides.
 */
export function parse(jsonText: string): PptxPresentation {
  let parsed: unknown;

  try {
    parsed = JSON.parse(jsonText);
  } catch {
    throw new Error(INVALID_FILE_MESSAGE);
  }

  const root = asRecord(parsed);

  if (!root) {
    throw new Error(INVALID_FILE_MESSAGE);
  }

  const presentation = asRecord(root['presentation']);

  if (!presentation) {
    throw new Error(INVALID_FILE_MESSAGE);
  }

  // FATAL: dimensions must exist — they drive every scale computation.
  const width = asNumber(presentation['width']);
  const height = asNumber(presentation['height']);

  if (width === null || height === null) {
    throw new Error('Invalid PPTX JSON: missing presentation.width / presentation.height.');
  }

  // FATAL: slides must be an array. Any length (0+) is allowed — the
  // renderer and thumbnail panel are fully dynamic.
  const rawSlides = asArray(presentation['slides']);

  if (rawSlides === null) {
    throw new Error(INVALID_FILE_MESSAGE);
  }

  return {
    width,
    height,
    slides: rawSlides.map(normalizeSlide), // ALL slides, any count
  };
}