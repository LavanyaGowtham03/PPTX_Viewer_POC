import './styles/pptx-viewer.css';
import { renderWelcomeScreen } from './ui/WelcomeScreen';
import { renderFullViewer } from './ui/ViewerContainer';
/**
 * Mounts the viewer in "browse mode".
 *
 * There is NO preloaded data — the viewer always starts with the
 * file-browse welcome screen. When a valid PPTX JSON file is selected,
 * the welcome screen is replaced by the full viewer.
 */
export function initPPTXViewer(options) {
    const { container } = options;
    // Step 1: Render the welcome/browse screen inside the container
    renderWelcomeScreen(container, (data) => {
        // Step 2: On valid JSON loaded, hide welcome screen and mount full viewer
        container.innerHTML = '';
        renderFullViewer(container, data);
    });
}
