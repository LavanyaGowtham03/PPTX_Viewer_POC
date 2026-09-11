import type { TextBoxElement } from '../../models/pptx-models';
/**
 * TextBoxRenderer — JSON type: "textbox"
 *
 * Renders one positioned div; the flat `text` string is split on "\n"
 * into one <p> per line. Formatting from `style` is uniform across the
 * whole box (the JSON has no paragraphs/runs arrays for textboxes).
 *
 * Never throws: parser-guaranteed defaults make every field safe.
 */
export declare function renderTextBox(el: TextBoxElement, scale: number): HTMLElement;
