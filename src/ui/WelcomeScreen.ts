import type { PPTXFile } from '../core/models/pptx-models';
import { INVALID_FILE_MESSAGE, parse } from '../core/parser/PptxParser';

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
export function renderWelcomeScreen(
  container: HTMLElement,
  onLoad: (data: PPTXFile) => void
): void {
  container.innerHTML = '';

  // ── Card structure ────────────────────────────────────────────
  const screen = document.createElement('div');
  screen.className = 'pptx-welcome';

  const card = document.createElement('div');
  card.className = 'pptx-welcome__card';

  const icon = document.createElement('div');
  icon.className = 'pptx-welcome__icon';
  icon.textContent = '📂';

  const title = document.createElement('h1');
  title.className = 'pptx-welcome__title';
  title.textContent = 'PowerPoint Viewer';

  const subtitle = document.createElement('p');
  subtitle.className = 'pptx-welcome__subtitle';
  subtitle.textContent = 'Browse and open a PowerPoint JSON file to get started.';

  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'pptx-welcome__btn';
  button.textContent = '📂 Browse JSON File';

  // Hidden native file input, opened programmatically by the button.
  const fileInput = document.createElement('input');
  fileInput.type = 'file';
  fileInput.accept = '.json,application/json';
  fileInput.className = 'pptx-welcome__file-input';

  const error = document.createElement('p');
  error.className = 'pptx-welcome__error';
  error.setAttribute('role', 'alert');
  error.hidden = true;

  const hint = document.createElement('p');
  hint.className = 'pptx-welcome__hint';
  hint.textContent = 'Supported format: .json (PPTX JSON export)';

  card.append(icon, title, subtitle, button, error, hint);
  screen.append(card, fileInput);
  container.append(screen);

  // ── Validation + wiring ──────────────────────────────────────
  const showInlineError = (message: string): void => {
    error.textContent = message;
    error.hidden = false;
  };

  const validateAndLoad = (text: string): void => {
    try {
      // Full parse + normalisation: fatal issues throw with clear
      // messages; optional fields get safe defaults inside the parser.
      const presentation = parse(text);

      onLoad({ presentation });
    } catch (err) {
      // Parser errors carry user-presentable messages; anything
      // unexpected still degrades to the generic invalid-file text.
      showInlineError(err instanceof Error ? err.message : INVALID_FILE_MESSAGE);
    }
  };

  const readFile = (file: File): void => {
    const reader = new FileReader();
    reader.onload = (): void => validateAndLoad(String(reader.result ?? ''));
    reader.onerror = (): void => showInlineError(INVALID_FILE_MESSAGE);
    reader.readAsText(file);
  };

  button.addEventListener('click', () => fileInput.click());

  fileInput.addEventListener('change', () => {
    const file = fileInput.files?.[0];
    if (!file) {
      return; // picker cancelled — stay on the welcome screen
    }
    readFile(file);
    // Reset so the SAME file can be picked again later.
    fileInput.value = '';
  });
}