import { applyPositionAndRotation, buildTextStyle } from './element-utils';
/**
 * TextBoxRenderer — JSON type: "textbox"
 *
 * Renders one positioned div; the flat `text` string is split on "\n"
 * into one <p> per line. Formatting from `style` is uniform across the
 * whole box (the JSON has no paragraphs/runs arrays for textboxes).
 *
 * Never throws: parser-guaranteed defaults make every field safe.
 */
export function renderTextBox(el, scale) {
    const wrapper = document.createElement('div');
    wrapper.className = 'pptx-element pptx-textbox';
    applyPositionAndRotation(wrapper, el, scale);
    // Uniform inline text formatting for the entire box.
    wrapper.style.cssText += `; ${buildTextStyle(el.style)}`;
    // One <p> per line — "\n" denotes line breaks in the flat text.
    const lines = el.text.split('\n');
    for (const line of lines) {
        const p = document.createElement('p');
        p.textContent = line; // textContent — never innerHTML (untrusted data)
        wrapper.appendChild(p);
    }
    // The parser defaults animation to undefined; AnimationEngine (later
    // step) adds classes/timing when present. Nothing to do here yet.
    return wrapper;
}
