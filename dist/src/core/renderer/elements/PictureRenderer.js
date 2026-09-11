import { applyPositionAndRotation } from './element-utils';
/**
 * PictureRenderer — JSON type: "picture"
 *
 * Renders an <img> inside a positioned wrapper. `src` supports both
 * base64 data URIs and http(s) URLs — set directly on img.src. An empty
 * src or a load failure swaps in a styled placeholder (⚠️ + "Image not
 * found"); the renderer never throws and never leaves a broken-image icon.
 */
export function renderPicture(el, scale) {
    const wrapper = document.createElement('div');
    wrapper.className = 'pptx-element pptx-picture';
    applyPositionAndRotation(wrapper, el, scale);
    const buildPlaceholder = () => {
        const placeholder = document.createElement('div');
        placeholder.className = 'pptx-picture__placeholder';
        placeholder.textContent = 'Image not found';
        return placeholder;
    };
    // Empty/missing src → placeholder immediately (no doomed <img> load).
    if (!el.src) {
        wrapper.appendChild(buildPlaceholder());
        return wrapper;
    }
    const img = document.createElement('img');
    img.alt = el.altText ?? '';
    img.src = el.src; // data URI or https:// URL, set directly
    // Broken URL / bad base64 → swap for the same styled placeholder.
    img.onerror = () => {
        img.remove();
        wrapper.appendChild(buildPlaceholder());
    };
    wrapper.appendChild(img);
    return wrapper;
}
