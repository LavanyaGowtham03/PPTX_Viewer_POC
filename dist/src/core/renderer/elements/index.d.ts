import type { PptxElement } from '../../models/pptx-models';
import { renderTextBox } from './TextBoxRenderer';
import { renderShape, renderShapeFallback } from './ShapeRenderer';
import { renderPicture } from './PictureRenderer';
import { renderTable } from './TableRenderer';
/**
 * Element dispatcher — routes each slide element to its renderer.
 *
 * Graceful handling contract (agent spec, Step 3):
 * - Known type            → delegated to its renderer
 * - `_unknown` flag       → console.warn with the element id; skipped
 *                            (unless it carries a shapeType, in which case
 *                            ShapeRenderer draws its plain-rect fallback)
 * - Unknown `type` string → same console.warn + skip, return null
 * - Never throws — callers wrap calls in try/catch as a final safety net.
 *
 * @returns the rendered node, or null when the element is skipped.
 */
export declare function renderElement(el: PptxElement, scale: number): HTMLElement | SVGElement | null;
export { renderTextBox, renderShape, renderShapeFallback, renderPicture, renderTable };
