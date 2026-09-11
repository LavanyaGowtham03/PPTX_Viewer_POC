// Lets `tsc` accept CSS imports (e.g. "../styles/pptx-viewer.css").
// Vite handles the actual CSS injection at dev/build time.
declare module '*.css';