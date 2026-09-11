import type { PptxPresentation, PptxSlide } from '../models/pptx-models';
/**
 * Render a slide into a .pptx-slide container.
 *
 * @param slide             parsed slide (elements already normalised)
 * @param pres              presentation (source of width/height)
 * @param scale             slide scale — all element coords × scale
 * @param suppressAnimations pass true for thumbnails: elements render
 *                          statically (no .pptx-animated classes), while a
 *                          real slide keeps animation hooks for the engine.
 */
export declare function renderSlide(slide: PptxSlide, pres: PptxPresentation, scale: number, suppressAnimations?: boolean): HTMLElement;
/**
 * Push presentation dimensions into CSS custom properties on the viewer
 * root. Called once per JSON load — everything derived (slide canvas,
 * thumbnail scale/height) recalculates from these values.
 */
export declare function applyPresentationDimensions(root: HTMLElement, pres: PptxPresentation): void;
/**
 * Slide scale for the main render area: fit availableWidth, never upscale
 * past 1 (a tiny slide renders at natural size, not magnified).
 */
export declare function computeScale(availableWidth: number, pres: PptxPresentation): number;
