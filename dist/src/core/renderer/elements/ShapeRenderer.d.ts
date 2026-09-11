import type { ShapeElement, UnknownElement } from '../../models/pptx-models';
/**
 * ShapeRenderer — JSON type: "shape"
 *
 * Renders an inline SVG whose viewBox is the UNSCALED element box, so
 * geometry math uses raw JSON units (points) while CSS scales the whole
 * wrapper. The parser guarantees fill/stroke/strokeWidth (with safe
 * defaults), so nothing here can throw on them.
 */
export declare function renderShape(el: ShapeElement, scale: number): HTMLElement;
/**
 * Unknown shapeType fallback — used by the elements dispatcher when an
 * element arrives flagged `_unknown` with a shapeType. Draws a plain
 * filled <rect> using the preserved raw colors; never throws.
 */
export declare function renderShapeFallback(el: UnknownElement, scale: number): HTMLElement;
