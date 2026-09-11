import { applyPositionAndRotation, buildTextStyle } from './element-utils';
import { buildGeometry, renderUnknownShapeFallback } from './shape-geometry';
/**
 * ShapeRenderer — JSON type: "shape"
 *
 * Renders an inline SVG whose viewBox is the UNSCALED element box, so
 * geometry math uses raw JSON units (points) while CSS scales the whole
 * wrapper. The parser guarantees fill/stroke/strokeWidth (with safe
 * defaults), so nothing here can throw on them.
 */
export function renderShape(el, scale) {
    const wrapper = document.createElement('div');
    wrapper.className = `pptx-element pptx-shape pptx-shape--${el.shapeType}`;
    applyPositionAndRotation(wrapper, el, scale);
    // Unscaled box drives the viewBox and all geometry math below.
    const W = el.width;
    const H = el.height;
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.setAttribute('width', '100%');
    svg.setAttribute('height', '100%');
    svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
    const geometry = buildGeometry(el.shapeType, W, H, {
        fill: el.fill,
        stroke: el.stroke,
        strokeWidth: el.strokeWidth ?? 1, // parser default; extra safety here
    });
    for (const node of geometry) {
        svg.appendChild(node);
    }
    wrapper.appendChild(svg);
    // Optional centered label — same inline text style rules as textbox.
    if (el.text !== undefined) {
        const label = document.createElement('div');
        label.className = 'pptx-shape__label';
        label.style.cssText = buildTextStyle(el.style);
        label.textContent = el.text;
        wrapper.appendChild(label);
    }
    return wrapper;
}
/**
 * Unknown shapeType fallback — used by the elements dispatcher when an
 * element arrives flagged `_unknown` with a shapeType. Draws a plain
 * filled <rect> using the preserved raw colors; never throws.
 */
export function renderShapeFallback(el, scale) {
    const wrapper = document.createElement('div');
    wrapper.className =
        `pptx-element pptx-shape pptx-shape--${el.shapeType ?? 'unknown'}`;
    applyPositionAndRotation(wrapper, el, scale);
    console.warn(`[PptxViewer] Unknown shapeType "${el.shapeType ?? '(missing)'}" (id: ${el.id}) — ` +
        `rendered plain filled rectangle fallback.`);
    return renderUnknownShapeFallback(wrapper, {
        width: el.width,
        height: el.height,
        fill: el.fill ?? '#000000',
        stroke: el.stroke ?? '#000000',
        strokeWidth: el.strokeWidth ?? 1,
    });
}
