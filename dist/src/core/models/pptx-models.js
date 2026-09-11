/**
 * PPTX Viewer — Typed Data Models
 *
 * The single source of truth for the PPTX JSON format consumed by this
 * viewer. Derived from the actual sample JSON exported by the PPTX →
 * JSON converter.
 *
 * Hard rules encoded here:
 * - ALL coordinates (x, y, width, height) are in POINTS (px) — no EMU.
 * - width/height are ALWAYS read from the JSON — never assume 960×540.
 * - slides can contain ANY number of slides (1, 2, 4, 10+).
 * - Every optional field is marked "?" and must be null-safe at use sites.
 */
export {};
