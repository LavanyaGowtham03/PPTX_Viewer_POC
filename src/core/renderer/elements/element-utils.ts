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
export function applyPositionAndRotation(
  target: HTMLElement,
  el: BaseElement,
  scale: number
): void {
  target.style.left = `${el.x * scale}px`;
  target.style.top = `${el.y * scale}px`;
  target.style.width = `${el.width * scale}px`;
  target.style.height = `${el.height * scale}px`;

  if (el.rotation !== undefined && el.rotation !== 0) {
    target.style.transform = `rotate(${el.rotation}deg)`;
    target.style.transformOrigin = 'center center';
  }
}

/**
 * Build the inline CSS text styles from an ElementStyle. Every field is
 * defensively defaulted — a missing style block must NEVER crash a
 * renderer. Returns an array of `key: value` pairs for cssText assembly.
 *
 * Used by: TextBoxRenderer, ShapeRenderer (label), TableRenderer (cells).
 */
export function buildTextStyle(style: ElementStyle | undefined): string {
  const s: ElementStyle =
    style ??
    {
      fontFamily: 'Calibri',
      fontSize: 18,
      color: '#000000',
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
    };

  // Underline + strikethrough combine into one text-decoration value.
  const decorations: string[] = [];
  if (s.underline) decorations.push('underline');
  if (s.strikethrough) decorations.push('line-through');
  const textDecoration =
    decorations.length > 0 ? `text-decoration: ${decorations.join(' ')}` : '';

  return [
    `font-family: "${s.fontFamily}"`,
    `font-size: ${s.fontSize}pt`,
    `color: ${s.color}`,
    s.bold ? 'font-weight: bold' : '',
    s.italic ? 'font-style: italic' : '',
    textDecoration,
    s.alignment !== undefined ? `text-align: ${s.alignment}` : '',
  ]
    .filter(Boolean)
    .join('; ');
}