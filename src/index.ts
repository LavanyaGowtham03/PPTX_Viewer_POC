import './styles/pptx-viewer.css';

import type { PPTXFile } from './core/models/pptx-models';
import { renderWelcomeScreen } from './ui/WelcomeScreen';
import { renderFullViewer } from './ui/ViewerContainer';

/**
 * Mounts the viewer in "browse mode".
 *
 * There is NO preloaded data — the viewer always starts with the
 * file-browse welcome screen. When a valid PPTX JSON file is selected,
 * the welcome screen is replaced by the full viewer.
 */
export function initPPTXViewer(options: { container: HTMLElement }): void {
  const { container } = options;

  // Step 1: Render the welcome/browse screen inside the container
  renderWelcomeScreen(container, (data: PPTXFile) => {
    // Step 2: On valid JSON loaded, hide welcome screen and mount full viewer
    container.innerHTML = '';
    renderFullViewer(container, data);
  });
}

export type { PPTXFile };