import type { PptxElement, UnknownElement } from '../../models/pptx-models';
import { renderTextBox } from './TextBoxRenderer';
import { renderShape, renderShapeFallback } from './ShapeRenderer';
import { renderPicture } from './PictureRenderer';
import { renderTable } from './TableRenderer';
import type { ShapeElement } from '../../models/pptx-models';

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
export function renderElement(el: PptxElement, scale: number): HTMLElement | SVGElement | null {
  // Parser-flagged unknowns. Shape unknowns keep their fallback path;
  // every other unknown is warned + skipped.
  if (isUnknown(el)) {
    if (el.shapeType !== undefined) {
      return renderShapeFallback(el, scale);
    }
    console.warn(
      `[PptxViewer] Unknown element type "${el.type}" (id: ${el.id ?? 'unnamed'}) — skipped.`
    );
    return null;
  }

  switch (el.type) {
    case 'textbox':
      return renderTextBox(el, scale);
    case 'shape':
      return renderShape(el as ShapeElement, scale);
    case 'picture':
      return renderPicture(el, scale);
    case 'table':
      return renderTable(el, scale);
    default: {
      // Defensive: a raw JSON path with an unlisted type string. After
      // the union narrowing TS sees `never` here, so read the raw fields
      // through the UnknownElement shape for the warning message.
      const raw = el as UnknownElement;
      console.warn(
        `[PptxViewer] Unknown element type "${String(raw.type)}" (id: ${String(raw.id ?? 'unnamed')}) — skipped.`
      );
      return null;
    }
  }
}

/** Type guard — parser flags unrecognised elements with `_unknown: true`. */
const isUnknown = (el: PptxElement): el is UnknownElement =>
  (el as UnknownElement)._unknown === true;

export { renderTextBox, renderShape, renderShapeFallback, renderPicture, renderTable };