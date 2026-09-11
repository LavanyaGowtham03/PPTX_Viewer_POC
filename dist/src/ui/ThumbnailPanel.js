import { renderSlide } from '../core/renderer/slide-renderer';
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
export class ThumbnailPanel {
    constructor(onSlideSelected) {
        this.onSlideSelected = onSlideSelected;
        this.panel = document.createElement('div');
        this.thumbs = [];
        this.pres = null;
        this.panel.className = 'pptx-thumbnail-panel';
    }
    render() {
        return this.panel;
    }
    /** Slide count of the loaded presentation (0 before setSlides). */
    getSlideCount() {
        return this.thumbs.length;
    }
    /** Rebuild the whole panel — called once per JSON load. */
    setSlides(pres) {
        this.pres = pres;
        this.panel.replaceChildren();
        this.thumbs.length = 0;
        // EVERY slide, any count — nothing is hardcoded.
        pres.slides.forEach((slide, index) => {
            this.panel.appendChild(this.buildThumbnail(slide, index, pres));
        });
    }
    /** Move the is-active highlight to thumbnail `index` and scroll to it. */
    setActiveIndex(index) {
        this.thumbs.forEach((thumb, i) => {
            const active = i === index;
            thumb.classList.toggle('is-active', active);
            if (active) {
                thumb.setAttribute('aria-current', 'true');
                // Optional call — not every DOM implementation provides it.
                thumb.scrollIntoView?.({ block: 'nearest' });
            }
            else {
                thumb.removeAttribute('aria-current');
            }
        });
    }
    /**
     * Render a MAIN-VIEW slide (animations attached) at the given scale.
     * The panel owns the presentation data; the container asks for scaled
     * renders through this instead of reaching into the data again.
     */
    renderSlideAt(index, scale) {
        if (this.pres === null) {
            return null;
        }
        const slide = this.pres.slides[index];
        if (slide === undefined) {
            return null;
        }
        return renderSlide(slide, this.pres, scale, false);
    }
    // ── Construction ─────────────────────────────────────────────────────────
    buildThumbnail(slide, index, pres) {
        const name = slide.name;
        const thumb = document.createElement('div');
        thumb.className = 'pptx-thumbnail';
        thumb.tabIndex = 0;
        thumb.setAttribute('role', 'button');
        thumb.setAttribute('aria-label', `Go to slide ${index + 1}: ${name}`);
        const canvas = document.createElement('div');
        canvas.className = 'pptx-thumbnail__canvas';
        // Mini slide at natural size — .pptx-thumbnail__slide scales it down
        // via the CSS variable; render is static (suppressAnimations = true).
        const mini = document.createElement('div');
        mini.className = 'pptx-thumbnail__slide';
        mini.appendChild(renderSlide(slide, pres, 1, true));
        const label = document.createElement('div');
        label.className = 'pptx-thumbnail__label';
        label.textContent = name; // parser defaults to "Slide N" if absent
        thumb.append(canvas, label);
        canvas.appendChild(mini);
        thumb.addEventListener('click', () => this.onSlideSelected(index));
        thumb.addEventListener('keydown', (ev) => {
            if (ev.key === 'Enter' || ev.key === ' ' || ev.key === 'Spacebar') {
                ev.preventDefault();
                this.onSlideSelected(index);
            }
        });
        this.thumbs.push(thumb);
        return thumb;
    }
}
