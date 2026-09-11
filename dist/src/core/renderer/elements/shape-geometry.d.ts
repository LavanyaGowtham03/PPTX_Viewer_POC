import type { ShapeType } from '../../models/pptx-models';
export interface ShapePaint {
    fill: string;
    stroke: string;
    strokeWidth: number;
}
export interface FallbackBox extends ShapePaint {
    width: number;
    height: number;
}
/** Build the SVG geometry nodes for a known shapeType. Never throws. */
export declare function buildGeometry(shapeType: ShapeType, W: number, H: number, paint: ShapePaint): SVGElement[];
/**
 * Plain filled <rect> fallback for unknown shapeTypes. Appends the rect
 * into the given wrapper's svg and returns the wrapper — used by the
 * dispatcher when an element is flagged _unknown.
 */
export declare function renderUnknownShapeFallback(wrapper: HTMLElement, box: FallbackBox): HTMLElement;
