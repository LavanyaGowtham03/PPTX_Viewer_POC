import type { AnimationType, ElementAnimation } from '../../models/pptx-models';

/**
 * AnimationEngine — element entry animations, CSS keyframes only.
 *
 * All @keyframes live in src/styles/pptx-viewer.css (pptx-fadeIn,
 * pptx-flyIn, pptx-flyIn-rotated, pptx-zoomIn, pptx-bounceIn,
 * pptx-fadedSwivel) — this engine only assigns classes and timing;
 * it never duplicates or generates keyframes.
 *
 * Element visibility model:
 * - .pptx-animated           → opacity: 0 until the animation begins
 * - .pptx-anim--<type>       → animation-name
 * - inline delay/duration    → seconds, straight from the JSON
 * - animations OFF (root.pptx-animations-off) → CSS !important shows all
 * - prefers-reduced-motion   → CSS disables everything, no JS involved
 */

const ANIMATED_CLASS = 'pptx-animated';
const ANIM_CLASS_PREFIX = 'pptx-anim--';

const KNOWN_ANIMATIONS: ReadonlySet<string> = new Set([
  'fadeIn',
  'flyIn',
  'zoomIn',
  'bounceIn',
  'fadedSwivel',
]);

/**
 * Attach the entry animation to a rendered element.
 *
 * @param el   the rendered element wrapper
 * @param anim parsed element.animation; undefined → no-op, the element
 *             simply stays fully visible
 */
export function applyAnimation(
  el: HTMLElement | SVGElement,
  anim: ElementAnimation | undefined
): void {
  if (anim === undefined) {
    return; // no animation in the JSON — nothing to attach
  }

  // Defensive even though the parser already normalises unknown types:
  // anything unrecognised at runtime falls back to fadeIn with a warning.
  const type: AnimationType = KNOWN_ANIMATIONS.has(String(anim.type))
    ? (anim.type as AnimationType)
    : 'fadeIn';

  if (String(anim.type) !== type) {
    console.warn(
      `[PptxViewer] Unknown animation type "${String(anim.type)}" — falling back to fadeIn.`
    );
  }

  el.classList.add(ANIMATED_CLASS, buildAnimClass(type, el));

  // JSON delay/duration are in SECONDS — CSS takes the "s" unit directly.
  el.style.animationDelay = `${anim.delay}s`;
  el.style.animationDuration = `${anim.duration}s`;
}

/**
 * Class name for the animation type. flyIn special-case: when the element
 * carries an inline rotate() (applyPositionAndRotation set it), the plain
 * pptx-flyIn keyframe would clobber that rotation via its own transform,
 * so the translate-based pptx-flyIn-rotated variant is used instead.
 */
const buildAnimClass = (type: AnimationType, el: HTMLElement | SVGElement): string => {
  if (type === 'flyIn' && el.style.transform !== '') {
    return `${ANIM_CLASS_PREFIX}flyIn-rotated`;
  }
  return `${ANIM_CLASS_PREFIX}${type}`;
};

/**
 * Restart every animated element on a slide — used when navigating back
 * to a slide so entry animations replay with their original timing
 * (each element keeps its own delay/duration, i.e. original order).
 *
 * Restart trick: strip the animation classes, force a reflow so the
 * running animation is discarded, then re-add the classes.
 */
export function replayAnimations(slideEl: HTMLElement): void {
  const animated = Array.from(slideEl.querySelectorAll<HTMLElement>(`.${ANIMATED_CLASS}`));

  if (animated.length === 0) {
    return; // nothing to replay — safe no-op
  }

  const saved: Array<{ el: HTMLElement; animClasses: string[] }> = [];

  for (const el of animated) {
    const animClasses = Array.from(el.classList).filter((cls) =>
      cls.startsWith(ANIM_CLASS_PREFIX)
    );
    el.classList.remove(ANIMATED_CLASS, ...animClasses);
    saved.push({ el, animClasses });
  }

  // One forced reflow for the whole batch — commits the removals.
  void slideEl.offsetWidth;

  for (const { el, animClasses } of saved) {
    el.classList.add(ANIMATED_CLASS, ...animClasses);
  }
}

/**
 * Toolbar Animations toggle. OFF adds .pptx-animations-off to the viewer
 * root — existing CSS (opacity: 1 !important; animation: none !important)
 * shows every .pptx-animated element instantly.
 */
export function setAnimationsEnabled(root: HTMLElement, enabled: boolean): void {
  root.classList.toggle('pptx-animations-off', !enabled);
}