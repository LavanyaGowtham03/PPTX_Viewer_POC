import type { BaseElement, ElementStyle } from '../../models/pptx-models';
/**
 * Shared helpers for all element renderers. Pure functions — no side
 * effects outside the returned/assigned elements.
 */
/**
 * Apply position/size (each JSON coordinate × slideScale) and rotation
 * to an element wrapper. Rotation uses transform-origin: center so the
 * element spins around its own middle, matching PowerPoint.
 */
export declare function applyPositionAndRotation(target: HTMLElement, el: BaseElement, scale: number): void;
/**
 * Build the inline CSS text styles from an ElementStyle. Every field is
 * defensively defaulted — a missing style block must NEVER crash a
 * renderer. Returns an array of `key: value` pairs for cssText assembly.
 *
 * Used by: TextBoxRenderer, ShapeRenderer (label), TableRenderer (cells).
 */
export declare function buildTextStyle(style: ElementStyle | undefined): string;
