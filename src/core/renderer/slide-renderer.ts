import type { PptxPresentation, PptxSlide } from '../models/pptx-models';
import { renderElement } from './elements';
import { applyAnimation } from './effects/AnimationEngine';

/**
 * SlideRenderer — renders one full slide.
 *
 * Dimensions ALWAYS come from pres.width / pres.height — never hardcoded.
 * Every element is dispatched inside try/catch so one bad element can
 * never break the rest of the slide.
 */

/** Fixed thumbnail column width in px — a UI constant, not from JSON. */
const THUMBNAIL_WIDTH = 160;

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
export function renderSlide(
  slide: PptxSlide,
  pres: PptxPresentation,
  scale: number,
  suppressAnimations: boolean = false
): HTMLElement {
  const container = document.createElement('div');
  container.className = 'pptx-slide';

  // Dimensions from JSON — the single source of truth for scaling.
  container.style.width = `${pres.width}px`;
  container.style.height = `${pres.height}px`;

  // Parser guarantees a background (default #FFFFFF); ?? is final safety.
  container.style.backgroundColor = slide.background ?? '#FFFFFF';

  container.style.transformOrigin = 'top left';
  container.style.transform = `scale(${scale})`;

  // EVERY element — any count. One failure must not stop the rest.
  for (const element of slide.elements) {
    try {
      const rendered = renderElement(element, scale);
      if (rendered !== null) {
        // Attach the entry animation unless this is a static render
        // (thumbnails). applyAnimation no-ops when animation is absent.
        if (!suppressAnimations) {
          applyAnimation(rendered, element.animation);
        }
        container.appendChild(rendered);
      }
    } catch (err) {
      console.warn(
        `[PptxViewer] Failed to render element "${element.id}" on slide ` +
          `"${slide.id}" — skipped (${String(err)})`
      );
      // Continue rendering the remaining elements.
    }
  }

  return container;
}

/**
 * Push presentation dimensions into CSS custom properties on the viewer
 * root. Called once per JSON load — everything derived (slide canvas,
 * thumbnail scale/height) recalculates from these values.
 */
export function applyPresentationDimensions(
  root: HTMLElement,
  pres: PptxPresentation
): void {
  root.style.setProperty('--pptx-slide-width', `${pres.width}px`);
  root.style.setProperty('--pptx-slide-height', `${pres.height}px`);
  root.style.setProperty('--pptx-thumbnail-scale', `${THUMBNAIL_WIDTH / pres.width}`);
  root.style.setProperty(
    '--pptx-thumbnail-height',
    `${Math.round(THUMBNAIL_WIDTH * (pres.height / pres.width))}px`
  );
  // JSON-derived aspect ratio for the main slide wrapper (fallback 16/9).
  root.style.setProperty('--pptx-slide-ratio', `${pres.width} / ${pres.height}`);
}

/**
 * Slide scale for the main render area: fit availableWidth, never upscale
 * past 1 (a tiny slide renders at natural size, not magnified).
 */
export function computeScale(availableWidth: number, pres: PptxPresentation): number {
  return Math.min(availableWidth / pres.width, 1);
}