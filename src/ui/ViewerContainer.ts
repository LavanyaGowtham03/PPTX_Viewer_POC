import type { PPTXFile, PptxPresentation } from '../core/models/pptx-models';
import { INVALID_FILE_MESSAGE, parse } from '../core/parser/PptxParser';
import {
  applyPresentationDimensions,
  computeScale,
} from '../core/renderer/slide-renderer';
import {
  setAnimationsEnabled,
  setTransitionsEnabled,
  transitionTo,
} from '../core/renderer/effects';
import { Toolbar } from './Toolbar';
import { ThumbnailPanel } from './ThumbnailPanel';

/**
 * renderFullViewer — the complete viewer (toolbar + thumbnails + main
 * slide area), replacing the Prompt 1 stub. Only called with VALID,
 * already-parsed data; file (re)loading and validation stay in
 * initPPTXViewer's shared pipeline so both entry paths (welcome Browse
 * and in-viewer Search) reset through the exact same flow.
 *
 * Browse-first entry stays intact: initPPTXViewer always starts with
 * the welcome screen; renderFullViewer only runs after a valid load.
 */
export function renderFullViewer(container: HTMLElement, data: PPTXFile): void {
  const pres: PptxPresentation = data.presentation;

  // ── Root & shell ──────────────────────────────────────────────────────
  container.replaceChildren();

  const root = document.createElement('div');
  root.className = 'pptx-viewer';
  container.appendChild(root);

  // Error banner for failed Search loads — hidden until needed.
  const errorBanner = document.createElement('div');
  errorBanner.className = 'pptx-error-banner';
  errorBanner.setAttribute('role', 'alert');
  errorBanner.hidden = true;
  const icon = document.createElement('span');
  icon.className = 'pptx-error-banner__icon';
  icon.textContent = '⚠️';
  errorBanner.append(icon);

  // Dimensions FIRST — every derived scale reads these vars.
  applyPresentationDimensions(root, pres);

  // ── UI components ─────────────────────────────────────────────────────
  const thumbnails = new ThumbnailPanel((index) => goToSlide(index));
  const toolbar = new Toolbar({
    onAnimationsToggle: (enabled) => setAnimationsEnabled(root, enabled),
    onTransitionsToggle: (enabled) => setTransitionsEnabled(root, enabled),
    onFileSelected: () => hideErrorBanner(), // picker opens inside Toolbar
  });

  // Search-tab picks → SAME validation pipeline as the welcome screen.
  toolbar.onFileLoaded((file) => void loadViaSearch(file));

  // ── Layout: toolbar / body(panel + main) ──────────────────────────────
  const body = document.createElement('div');
  body.className = 'pptx-body';

  const main = document.createElement('div');
  main.className = 'pptx-main';

  const slideWrapper = document.createElement('div');
  slideWrapper.className = 'pptx-slide-wrapper';

  main.appendChild(slideWrapper);
  body.append(thumbnails.render(), main);
  root.append(toolbar.render(), errorBanner, body);

  // ── ResizeObserver: main-area width → slide scale ────────────────────
  let currentScale = 1;
  let activeSlideEl: HTMLElement | null = null;

  const rescale = (): void => {
    // contentRect accounts for padding — the space the slide can fill.
    const available = main.clientWidth;
    if (available > 0 && pres.width > 0) {
      currentScale = computeScale(available, pres);
      if (activeSlideEl !== null) {
        activeSlideEl.style.transform = `scale(${currentScale})`;
      }
      // Wrapper sized to the SCALED slide so centering is exact.
      slideWrapper.style.width = `${pres.width * currentScale}px`;
      slideWrapper.style.height = `${pres.height * currentScale}px`;
    }
  };

  const resizeObserver = new ResizeObserverSafe(() => rescale());
  resizeObserver.observe(main);

  // ── Navigation ────────────────────────────────────────────────────────
  let currentIndex = 0;
  let navLock = false; // ignore input while a transition is running

  const hideErrorBanner = (): void => {
    errorBanner.hidden = true;
    errorBanner.replaceChildren(icon);
  };

  const showErrorBanner = (message: string): void => {
    const text = document.createElement('span');
    text.textContent = message;
    errorBanner.replaceChildren(icon, text);
    errorBanner.hidden = false;
  };

  const showSlide = (index: number): void => {
    const fresh = thumbnails.renderSlideAt(index, currentScale);
    if (fresh === null) {
      console.warn(`[PptxViewer] Slide ${index + 1} unavailable — no render.`);
      return;
    }
    activeSlideEl = fresh;
    slideWrapper.replaceChildren(fresh);
    rescale();
  };

  const updateInfo = (): void => {
    const slide = pres.slides[currentIndex];
    toolbar.setSlideInfo(currentIndex, pres.slides.length, slide?.name ?? '');
    thumbnails.setActiveIndex(currentIndex);
  };

  const goToSlide = (index: number): void => {
    const total = pres.slides.length;
    if (total === 0) {
      return;
    }
    const clamped = Math.max(0, Math.min(index, total - 1));

    if (clamped === currentIndex && activeSlideEl !== null && !firstShow) {
      return; // already viewing this slide
    }

    // First appearance: no predecessor, just mount.
    if (firstShow || activeSlideEl === null) {
      firstShow = false;
      currentIndex = clamped;
      showSlide(clamped);
      updateInfo();
      return;
    }

    if (navLock) {
      return; // a transition is mid-flight
    }
    navLock = true;

    const next = thumbnails.renderSlideAt(clamped, currentScale);
    if (next === null) {
      navLock = false;
      return;
    }

    // INCOMING slide's transition governs (spec §6).
    const incoming = pres.slides[clamped]?.transition;

    void transitionTo(activeSlideEl as HTMLElement, next, incoming, root)
      .then(() => {
        currentIndex = clamped;
        activeSlideEl = next;
        updateInfo();
      })
      .finally(() => {
        navLock = false;
      });
  };

  let firstShow = true;

  // ── Mount: thumbnails, first slide, initial scale ────────────────────
  thumbnails.setSlides(pres);
  goToSlide(0);
  rescale(); // measure now; ResizeObserver handles later changes

  // ── Search-tab file load → shared validation + FULL reset ────────────
  const loadViaSearch = async (file: File): Promise<void> => {
    try {
      const text = await file.text();
      const presentation = parse(text); // throws with a clear message

      // FULL reset: nothing from the old presentation survives —
      // renderFullViewer rebuilds root, thumbnails, scales, tab state.
      const disposer = disposerMap.get(container);
      disposer?.(); // detach the old observer + keydown listener
      renderFullViewer(container, { presentation });
    } catch (err) {
      // Inline banner INSIDE the viewer; state is untouched.
      showErrorBanner(err instanceof Error ? err.message : INVALID_FILE_MESSAGE);
    }
  };

  // ── Keyboard navigation ───────────────────────────────────────────────
  const keydownHandler = (ev: KeyboardEvent): void => {
    if (ev.defaultPrevented) {
      return;
    }
    if (ev.key === 'ArrowRight') {
      goToSlide(currentIndex + 1);
    } else if (ev.key === 'ArrowLeft') {
      goToSlide(currentIndex - 1);
    }
  };
  document.addEventListener('keydown', keydownHandler);

  // This viewer's disposables — drained by the next full reset (or a
  // future unmount). Registry is per-container, so parallel viewers in
  // different containers never interfere.
  disposerMap.set(container, () => {
    resizeObserver.disconnect();
    document.removeEventListener('keydown', keydownHandler);
  });
}

// Per-container cleanup registry — a fresh renderFullViewer on the SAME
// container disposes the previous viewer's observer/listener first.
const disposerMap = new WeakMap<HTMLElement, () => void>();

/**
 * ResizeObserver wrapper — the real API in browsers, a no-op shim in
 * environments without it (jsdom). Fitting still happens via the initial
 * rescale() call; live window resizes only work where the API exists.
 */
class ResizeObserverSafe implements ResizeObserver {
  private readonly impl: ResizeObserver | null;
  private readonly fallbackCallback: (() => void) | null = null;

  constructor(callback: () => void) {
    if (typeof ResizeObserver !== 'undefined') {
      this.impl = new ResizeObserver(callback);
    } else {
      this.impl = null;
      this.fallbackCallback = callback;
    }
  }

  observe(target: Element): void {
    if (this.impl !== null) {
      this.impl.observe(target);
    }
    // Shim: deliver one initial tick so layout-math runs once.
    this.fallbackCallback?.();
  }

  unobserve(target: Element): void {
    this.impl?.unobserve(target);
  }

  disconnect(): void {
    this.impl?.disconnect();
  }
}