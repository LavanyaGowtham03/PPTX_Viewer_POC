import type { ElementAnimation } from '../../models/pptx-models';
/**
 * Attach the entry animation to a rendered element.
 *
 * @param el   the rendered element wrapper
 * @param anim parsed element.animation; undefined → no-op, the element
 *             simply stays fully visible
 */
export declare function applyAnimation(el: HTMLElement | SVGElement, anim: ElementAnimation | undefined): void;
/**
 * Restart every animated element on a slide — used when navigating back
 * to a slide so entry animations replay with their original timing
 * (each element keeps its own delay/duration, i.e. original order).
 *
 * Restart trick: strip the animation classes, force a reflow so the
 * running animation is discarded, then re-add the classes.
 */
export declare function replayAnimations(slideEl: HTMLElement): void;
/**
 * Toolbar Animations toggle. OFF adds .pptx-animations-off to the viewer
 * root — existing CSS (opacity: 1 !important; animation: none !important)
 * shows every .pptx-animated element instantly.
 */
export declare function setAnimationsEnabled(root: HTMLElement, enabled: boolean): void;
