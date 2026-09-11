import './styles/pptx-viewer.css';
import type { PPTXFile } from './core/models/pptx-models';
/**
 * Mounts the viewer in "browse mode".
 *
 * There is NO preloaded data — the viewer always starts with the
 * file-browse welcome screen. When a valid PPTX JSON file is selected,
 * the welcome screen is replaced by the full viewer.
 */
export declare function initPPTXViewer(options: {
    container: HTMLElement;
}): void;
export type { PPTXFile };
