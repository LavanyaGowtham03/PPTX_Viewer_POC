import type { PptxPresentation } from '../core/models/pptx-models';
/**
 * ThumbnailPanel — left navigation column with one mini render per slide.
 *
 * Every thumbnail is a real renderSlide() output at NATURAL JSON size
 * (scale 1); CSS shrinks it via --pptx-thumbnail-scale (set on the viewer
 * root by applyPresentationDimensions). Thumbnails are STATIC:
 * suppressAnimations = true, no transition classes — previews only.
 *
 * Fully dynamic: 1, 2, 4, 10+ slides all populate identically.
 */
export declare class ThumbnailPanel {
    private readonly onSlideSelected;
    private readonly panel;
    private readonly thumbs;
    private pres;
    constructor(onSlideSelected: (index: number) => void);
    render(): HTMLElement;
    /** Slide count of the loaded presentation (0 before setSlides). */
    getSlideCount(): number;
    /** Rebuild the whole panel — called once per JSON load. */
    setSlides(pres: PptxPresentation): void;
    /** Move the is-active highlight to thumbnail `index` and scroll to it. */
    setActiveIndex(index: number): void;
    /**
     * Render a MAIN-VIEW slide (animations attached) at the given scale.
     * The panel owns the presentation data; the container asks for scaled
     * renders through this instead of reaching into the data again.
     */
    renderSlideAt(index: number, scale: number): HTMLElement | null;
    private buildThumbnail;
}
