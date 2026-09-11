import type { PPTXFile } from '../core/models/pptx-models';
/**
 * renderFullViewer — the complete viewer (toolbar + thumbnails + main
 * slide area), replacing the Prompt 1 stub. Only called with VALID,
 * already-parsed data; file (re)loading and validation stay in
 * initPPTXViewer's shared pipeline so both entry paths (welcome Browse
 * and in-viewer Search) reset through the exact same flow.
 *
 * Browse-first entry stays intact: initPPTXViewer always starts with
 * the welcome screen; renderFullViewer only runs after a valid load.
 */
export declare function renderFullViewer(container: HTMLElement, data: PPTXFile): void;
