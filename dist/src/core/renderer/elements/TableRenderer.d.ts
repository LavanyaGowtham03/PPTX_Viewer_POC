import type { TableElement } from '../../models/pptx-models';
/**
 * TableRenderer — JSON type: "table"
 *
 * Renders a standard <table> filling the wrapper: border-collapse +
 * table-layout: fixed for stable merged-cell layout. Missing/empty rows
 * produce an empty <table> — never an exception.
 */
export declare function renderTable(el: TableElement, scale: number): HTMLElement;
