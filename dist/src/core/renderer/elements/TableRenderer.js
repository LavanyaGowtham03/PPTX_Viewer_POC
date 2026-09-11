import { applyPositionAndRotation, buildTextStyle } from './element-utils';
const DEFAULT_BORDER_COLOR = '#d1d5db';
/** Build one <td> — merges, colors, borders, and text formatting. */
const renderCell = (cell) => {
    const td = document.createElement('td');
    // Merged cells — only meaningful when > 1.
    if (cell.rowSpan !== undefined && cell.rowSpan > 1) {
        td.rowSpan = cell.rowSpan;
    }
    if (cell.colSpan !== undefined && cell.colSpan > 1) {
        td.colSpan = cell.colSpan;
    }
    // Per-cell paints — defaults keep the border visible when absent.
    if (cell.backgroundColor !== undefined) {
        td.style.backgroundColor = cell.backgroundColor;
    }
    td.style.borderColor = cell.borderColor ?? DEFAULT_BORDER_COLOR;
    // Same inline text formatting rules as the textbox renderer.
    if (cell.style !== undefined) {
        td.style.cssText += `; ${buildTextStyle(cell.style)}`;
    }
    // Cell text — "\n" handled via <br> (textContent alone would show a
    // literal newline without browser line-breaking inside table cells).
    if (cell.text !== undefined) {
        for (const [i, line] of cell.text.split('\n').entries()) {
            if (i > 0) {
                td.appendChild(document.createElement('br'));
            }
            td.appendChild(document.createTextNode(line));
        }
    }
    return td;
};
/**
 * TableRenderer — JSON type: "table"
 *
 * Renders a standard <table> filling the wrapper: border-collapse +
 * table-layout: fixed for stable merged-cell layout. Missing/empty rows
 * produce an empty <table> — never an exception.
 */
export function renderTable(el, scale) {
    const wrapper = document.createElement('div');
    wrapper.className = 'pptx-element pptx-table';
    applyPositionAndRotation(wrapper, el, scale);
    const table = document.createElement('table');
    const tbody = document.createElement('tbody');
    // rows is parser-guaranteed to be an array (default []); extra safety
    // here keeps the renderer total-callback-proof inside try/catch callers.
    const rows = Array.isArray(el.rows) ? el.rows : [];
    for (const row of rows) {
        const tr = document.createElement('tr');
        const cells = Array.isArray(row?.cells) ? row.cells : [];
        for (const cell of cells) {
            tr.appendChild(renderCell(cell ?? {}));
        }
        tbody.appendChild(tr);
    }
    table.appendChild(tbody);
    wrapper.appendChild(table);
    return wrapper;
}
