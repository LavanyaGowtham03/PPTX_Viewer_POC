import type { PptxPresentation } from '../models/pptx-models';
/**
 * PptxParser — raw JSON text → typed, normalised PptxPresentation.
 *
 * Error philosophy:
 * - FATAL (throw, with a clear message): the file isn't a presentation
 *   at all — bad JSON syntax, or missing presentation.width/height/slides.
 * - SAFE (never throw): every optional field gets a default; unknown
 *   element types / shapeTypes are kept but flagged `_unknown: true` so
 *   the renderer can skip them with a console.warn.
 *
 * The slide count is ALWAYS dynamic — 1, 2, 4, 10+ slides all parse
 * identically; nothing here assumes a count or hardcoded dimensions.
 */
/** Shared message for any file that isn't a valid PPTX JSON export. */
export declare const INVALID_FILE_MESSAGE = "Invalid file. Please select a valid PPTX JSON file.";
/**
 * Parse the raw JSON text of a PPTX JSON export.
 *
 * @throws Error (with a user-presentable message) only for FATAL issues:
 *         broken JSON syntax, or missing presentation.width/height/slides.
 */
export declare function parse(jsonText: string): PptxPresentation;
