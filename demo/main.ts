import { initPPTXViewer } from '../src/index';

// Mount the viewer in "browse mode" — no preloaded data.
// The viewer starts with the file-browse welcome screen.
initPPTXViewer({
  container: document.getElementById('pptx-root')!,
});