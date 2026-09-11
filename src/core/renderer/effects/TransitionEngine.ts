import type { SlideTransition, TransitionDirection, TransitionType } from '../../models/pptx-models';

/**
 * TransitionEngine — slide-to-slide transitions via CSS class swaps.
 *
 * All keyframes and class rules live in pptx-viewer.css:
 *   .pptx-slide--transition-<type>[-<direction>] + .pptx-slide-leave-active
 *   / .pptx-slide-enter-active. The engine never generates CSS.
 *
 * Slides keep their own inline transform: scale() (from renderSlide) —
 * transition animations run on WRAPPER divs around the slides so the
 * keyframe transforms never clobber the slide scale.
 */

const KNOWN_TYPES: ReadonlySet<string> = new Set(['fade', 'push', 'wipe', 'zoom']);
const KNOWN_DIRECTIONS: ReadonlySet<string> = new Set(['left', 'right', 'up', 'down']);

/** Default total transition time in ms (spec: 300 fade, 400 otherwise). */
const DEFAULT_DURATION: Record<TransitionType, number> = {
  fade: 300,
  push: 400,
  wipe: 400,
  zoom: 400,
};

const prefersReducedMotion = (): boolean =>
  typeof window !== 'undefined' &&
  typeof window.matchMedia === 'function' &&
  window.matchMedia('(prefers-reduced-motion: reduce)').matches;

/** Toolbar Transitions toggle — OFF makes every navigation instant. */
export function setTransitionsEnabled(root: HTMLElement, enabled: boolean): void {
  root.classList.toggle('pptx-transitions-off', !enabled);
}

/**
 * Swap `current` (attached) for `next` (fresh, unattached), playing the
 * INCOMING slide's transition. Resolves when the animation settles (or a
 * safety timeout), leaving no classes/DOM behind.
 */
export async function transitionTo(
  current: HTMLElement,
  next: HTMLElement,
  transition: SlideTransition | undefined,
  root: HTMLElement
): Promise<void> {
  const host = current.parentElement;

  if (host === null) {
    console.warn('[PptxViewer] transitionTo: current slide is detached — skipping swap.');
    return;
  }

  const instantSwap = (): void => {
    current.remove();
    host.appendChild(next);
  };

  // Instant paths: no transition declared, toggled OFF, or reduced motion.
  if (
    transition === undefined ||
    root.classList.contains('pptx-transitions-off') ||
    prefersReducedMotion()
  ) {
    instantSwap();
    return;
  }

  // Parser already normalises; defensive double-guard with fade fallback.
  const rawType = String(transition.type);
  const type: TransitionType = KNOWN_TYPES.has(rawType)
    ? (transition.type as TransitionType)
    : 'fade';
  if (rawType !== type) {
    console.warn(`[PptxViewer] Unknown transition type "${rawType}" — falling back to fade.`);
  }

  const rawDir = transition.direction;
  const direction: TransitionDirection | undefined =
    rawDir !== undefined && KNOWN_DIRECTIONS.has(rawDir)
      ? (rawDir as TransitionDirection)
      : undefined;

  const needsDirection = type === 'push' || type === 'wipe';
  const base =
    needsDirection && direction !== undefined
      ? `pptx-slide--transition-${type}-${direction}`
      : `pptx-slide--transition-${type}`;

  const duration = transition.duration ?? DEFAULT_DURATION[type];

  // Duration must be in place before the animation starts.
  root.style.setProperty('--pptx-slide-transition-duration', `${duration}ms`);

  // Wrappers: enter/leave classes position them (absolute, inset 0) and
  // the -active rules run the keyframes — on the WRAPPER, so the slide's
  // own scale transform is untouched.
  const leaveWrap = document.createElement('div');
  leaveWrap.className = `pptx-slide-leave ${base} pptx-slide-leave-active`;

  const enterWrap = document.createElement('div');
  enterWrap.className = `pptx-slide-enter ${base} pptx-slide-enter-active`;

  current.replaceWith(leaveWrap);
  leaveWrap.appendChild(current);
  enterWrap.appendChild(next);
  host.appendChild(enterWrap);

  await waitForSettled([leaveWrap, enterWrap], duration + 200);

  // Cleanup — unwrap the incoming slide, drop the outgoing one entirely.
  enterWrap.replaceWith(next);
  leaveWrap.remove();
  root.style.removeProperty('--pptx-slide-transition-duration');
}

/**
 * Resolve when every wrapper's animation ends, or on a safety timeout
 * (covers killed animations, e.g. reduced-motion CSS, or environments
 * without a running animation pipeline).
 */
const waitForSettled = (nodes: HTMLElement[], timeoutMs: number): Promise<void> =>
  new Promise<void>((resolve) => {
    let remaining = nodes.length;
    let settled = false;

    const finish = (): void => {
      if (!settled) {
        settled = true;
        resolve();
      }
    };

    for (const node of nodes) {
      node.addEventListener('animationend', (ev: Event) => {
        // Only the wrapper's own animation counts — element entry
        // animations bubble up from children and must be ignored.
        if (ev.target === ev.currentTarget) {
          remaining -= 1;
          if (remaining <= 0) {
            finish();
          }
        }
      });
    }

    setTimeout(finish, timeoutMs);
  });