import type { PPTXFile } from '../core/models/pptx-models';
/**
 * PROMPT 3 — Browse-first welcome screen, validated via PptxParser.
 *
 * There is NO preloaded data. The user must select a .json file to load
 * the presentation.
 *
 * Flow:
 *   [📂 Browse JSON File] → hidden <input type="file" accept=".json"> →
 *   FileReader.readAsText → PptxParser.parse inside try/catch →
 *     valid   → onLoad(data)
 *     invalid → inline error inside the card, carrying the parser's
 *               own descriptive message.
 */
export declare function renderWelcomeScreen(container: HTMLElement, onLoad: (data: PPTXFile) => void): void;
