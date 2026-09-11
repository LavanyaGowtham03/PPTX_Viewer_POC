import type { PictureElement } from '../../models/pptx-models';
/**
 * PictureRenderer — JSON type: "picture"
 *
 * Renders an <img> inside a positioned wrapper. `src` supports both
 * base64 data URIs and http(s) URLs — set directly on img.src. An empty
 * src or a load failure swaps in a styled placeholder (⚠️ + "Image not
 * found"); the renderer never throws and never leaves a broken-image icon.
 */
export declare function renderPicture(el: PictureElement, scale: number): HTMLElement;
