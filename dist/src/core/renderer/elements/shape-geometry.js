/**
 * Pure SVG geometry builders — one per known shapeType.
 * All coordinates are in the element's UNSCALED JSON units, matching
 * the svg viewBox set by ShapeRenderer.
 */
const SVG_NS = 'http://www.w3.org/2000/svg';
const attrs = (element, values) => {
    for (const [key, value] of Object.entries(values)) {
        element.setAttribute(key, String(value));
    }
};
/** oval → <ellipse> centered in the bounding box */
const buildOval = (W, H, paint) => {
    const ellipse = document.createElementNS(SVG_NS, 'ellipse');
    attrs(ellipse, {
        cx: W / 2,
        cy: H / 2,
        rx: W / 2,
        ry: H / 2,
        fill: paint.fill,
        stroke: paint.stroke,
        'stroke-width': paint.strokeWidth,
    });
    return ellipse;
};
/** triangle → isosceles: top-center, bottom-right, bottom-left */
const buildTriangle = (W, H, paint) => {
    const polygon = document.createElementNS(SVG_NS, 'polygon');
    attrs(polygon, {
        points: `${W / 2},0 ${W},${H} 0,${H}`,
        fill: paint.fill,
        stroke: paint.stroke,
        'stroke-width': paint.strokeWidth,
    });
    return polygon;
};
/**
 * blockArrow (UP) → 7-point thick upward arrow:
 *   head spans full width at the top, shaft is the middle ~half.
 */
const buildBlockArrow = (W, H, paint) => {
    const polygon = document.createElementNS(SVG_NS, 'polygon');
    attrs(polygon, {
        points: [
            `${W * 0.25},0`, // head left-top
            `${W * 0.75},0`, // head right-top
            `${W * 0.75},${H * 0.4}`, // shaft right
            `${W},${H * 0.4}`, // barb right
            `${W * 0.5},${H}`, // tip bottom (arrow points UP)
            `0,${H * 0.4}`, // barb left
            `${W * 0.25},${H * 0.4}`, // shaft left
        ].join(' '),
        fill: paint.fill,
        stroke: paint.stroke,
        'stroke-width': paint.strokeWidth,
    });
    return polygon;
};
/** flowchartTerminator → stadium/pill: rect with rx/ry = half the height */
const buildFlowchartTerminator = (W, H, paint) => {
    const rect = document.createElementNS(SVG_NS, 'rect');
    attrs(rect, {
        x: 0,
        y: 0,
        width: W,
        height: H,
        rx: H / 2,
        ry: H / 2,
        fill: paint.fill,
        stroke: paint.stroke,
        'stroke-width': paint.strokeWidth,
    });
    return rect;
};
/**
 * calloutRectangle → speech bubble: rounded rect body + triangular tail
 * at the bottom-left pointing down-left.
 */
const buildCalloutRectangle = (W, H, paint) => {
    const tailHeight = Math.min(H * 0.25, 24);
    const body = document.createElementNS(SVG_NS, 'rect');
    attrs(body, {
        x: 0,
        y: 0,
        width: W,
        height: H - tailHeight, // leave room for the tail
        rx: Math.min(8, W / 8),
        ry: Math.min(8, H / 8),
        fill: paint.fill,
        stroke: paint.stroke,
        'stroke-width': paint.strokeWidth,
    });
    const tail = document.createElementNS(SVG_NS, 'polygon');
    attrs(tail, {
        points: [
            `${W * 0.12},${H - tailHeight - 1}`, // top of tail (inside body)
            `${W * 0.3},${H - tailHeight - 1}`,
            `0,${H}`, // tip points down-left
        ].join(' '),
        fill: paint.fill,
        stroke: paint.stroke,
        'stroke-width': paint.strokeWidth,
    });
    return [body, tail];
};
/** Build the SVG geometry nodes for a known shapeType. Never throws. */
export function buildGeometry(shapeType, W, H, paint) {
    switch (shapeType) {
        case 'oval':
            return [buildOval(W, H, paint)];
        case 'triangle':
            return [buildTriangle(W, H, paint)];
        case 'blockArrow':
            return [buildBlockArrow(W, H, paint)];
        case 'flowchartTerminator':
            return [buildFlowchartTerminator(W, H, paint)];
        case 'calloutRectangle':
            return buildCalloutRectangle(W, H, paint);
        default: {
            // Exhaustiveness guard — unreachable with the ShapeType union.
            const never = shapeType;
            console.warn(`[PptxViewer] Missing geometry for "${String(never)}" — empty shape.`);
            return [];
        }
    }
}
/**
 * Plain filled <rect> fallback for unknown shapeTypes. Appends the rect
 * into the given wrapper's svg and returns the wrapper — used by the
 * dispatcher when an element is flagged _unknown.
 */
export function renderUnknownShapeFallback(wrapper, box) {
    const svg = document.createElementNS(SVG_NS, 'svg');
    svg.setAttribute('width', '100%');
    svg.setAttribute('height', '100%');
    svg.setAttribute('viewBox', `0 0 ${box.width} ${box.height}`);
    const rect = document.createElementNS(SVG_NS, 'rect');
    attrs(rect, {
        x: 0,
        y: 0,
        width: box.width,
        height: box.height,
        fill: box.fill,
        stroke: box.stroke,
        'stroke-width': box.strokeWidth,
    });
    svg.appendChild(rect);
    wrapper.appendChild(svg);
    return wrapper;
}
