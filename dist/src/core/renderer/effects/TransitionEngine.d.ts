import type { SlideTransition } from '../../models/pptx-models';
/** Toolbar Transitions toggle — OFF makes every navigation instant. */
export declare function setTransitionsEnabled(root: HTMLElement, enabled: boolean): void;
/**
 * Swap `current` (attached) for `next` (fresh, unattached), playing the
 * INCOMING slide's transition. Resolves when the animation settles (or a
 * safety timeout), leaving no classes/DOM behind.
 */
export declare function transitionTo(current: HTMLElement, next: HTMLElement, transition: SlideTransition | undefined, root: HTMLElement): Promise<void>;
