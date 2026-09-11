/**
 * Toolbar — POC tab bar (Home | Animations | Transitions | Search).
 *
 * Selecting a tab swaps ONLY the panel strip directly underneath;
 * slide content is never affected by tab switching. The Search tab is
 * always available so a different JSON can be loaded at any moment.
 *
 * No editing/ribbon features — POC scope by design.
 */

export interface ToolbarCallbacks {
  onAnimationsToggle: (enabled: boolean) => void;
  onTransitionsToggle: (enabled: boolean) => void;
  /** Opens the file picker for a new JSON (validation is wired by owner). */
  onFileSelected: () => void;
}

type TabId = 'home' | 'animations' | 'transitions' | 'search';

export class Toolbar {
  private readonly callbacks: ToolbarCallbacks;
  private readonly root = document.createElement('div');
  private readonly tabs = new Map<TabId, HTMLButtonElement>();
  private readonly panels = new Map<TabId, HTMLElement>();
  private readonly homeInfo = document.createElement('span');
  private animationsCheckbox: HTMLInputElement | null = null;
  private transitionsCheckbox: HTMLInputElement | null = null;

  constructor(callbacks: ToolbarCallbacks) {
    this.callbacks = callbacks;
    this.build();
  }

  /** Renders the toolbar: tab bar row + active-tab panel strip. */
  render(): HTMLElement {
    return this.root;
  }

  // ── Public state setters (called by ViewerContainer) ────────────────────

  /** Sync the Animations toggle (e.g. after a full viewer reset). */
  setAnimationsState(enabled: boolean): void {
    if (this.animationsCheckbox !== null) {
      this.animationsCheckbox.checked = enabled;
    }
  }

  /** Sync the Transitions toggle (e.g. after a full viewer reset). */
  setTransitionsState(enabled: boolean): void {
    if (this.transitionsCheckbox !== null) {
      this.transitionsCheckbox.checked = enabled;
    }
  }

  /** Update the Home panel read-only info. Navigation-safe: only the
   *  text changes — the active tab is NEVER switched here (spec resets
   *  to Home only on a full viewer reload). */
  setSlideInfo(index: number, total: number, name: string): void {
    this.homeInfo.textContent = `Slide ${index + 1} of ${total} — ${name}`;
  }

  /** Programmatically activate a tab (used by the Search flow). */
  activateTab(tab: TabId): void {
    this.selectTab(tab);
  }

  // ── Construction ─────────────────────────────────────────────────────────

  private build(): void {
    this.root.className = 'pptx-toolbar-root';

    // Tab row
    const bar = document.createElement('nav');
    bar.className = 'pptx-toolbar';
    bar.setAttribute('role', 'tablist');

    const homeTab = this.buildTab('home', 'Home');
    const animationsTab = this.buildTab('animations', 'Animations');
    const transitionsTab = this.buildTab('transitions', 'Transitions');
    const searchTab = this.buildTab('search', 'Search');

    bar.append(homeTab, animationsTab, transitionsTab, searchTab);

    // Panel strip — content swapped by active tab
    const panel = document.createElement('div');
    panel.className = 'pptx-toolbar__panel';
    panel.setAttribute('role', 'tabpanel');

    this.panels.set('home', this.buildHomePanel());
    this.panels.set('animations', this.buildTogglePanel(
      'Play animations', (checked) => this.callbacks.onAnimationsToggle(checked),
      (checkbox) => { this.animationsCheckbox = checkbox; }
    ));
    this.panels.set('transitions', this.buildTogglePanel(
      'Slide transitions', (checked) => this.callbacks.onTransitionsToggle(checked),
      (checkbox) => { this.transitionsCheckbox = checkbox; }
    ));
    this.panels.set('search', this.buildSearchPanel());

    this.root.append(bar, panel);
    this.selectTab('home'); // HOME is the default tab — info populated later
  }

  private buildTab(id: TabId, label: string): HTMLButtonElement {
    const tab = document.createElement('button');
    tab.type = 'button';
    tab.className = 'pptx-toolbar__tab';
    tab.textContent = label;
    tab.setAttribute('role', 'tab');
    tab.setAttribute('aria-selected', 'false');
    tab.addEventListener('click', () => this.selectTab(id));
    this.tabs.set(id, tab);
    return tab;
  }

  // ── Panels ───────────────────────────────────────────────────────────────

  /** HOME — current-slide info ("Slide X of N — name"), read-only. */
  private buildHomePanel(): HTMLElement {
    const wrap = document.createElement('div');
    wrap.className = 'pptx-toolbar__panel-info';

    this.homeInfo.textContent = 'No slide loaded';
    this.homeInfo.className = 'pptx-toolbar__panel-text';

    wrap.appendChild(this.homeInfo);
    return wrap;
  }

  /** ANIMATIONS / TRANSITIONS — one toggle each, starting ON. */
  private buildTogglePanel(
    label: string,
    onChange: (checked: boolean) => void,
    register: (checkbox: HTMLInputElement) => void
  ): HTMLElement {
    const toggle = document.createElement('label');
    toggle.className = 'pptx-toggle';

    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = true; // starts ON per spec
    checkbox.addEventListener('change', () => onChange(checkbox.checked));

    const text = document.createElement('span');
    text.textContent = label;

    toggle.append(checkbox, text);
    register(checkbox);

    const wrap = document.createElement('div');
    wrap.className = 'pptx-toolbar__panel-info';
    wrap.appendChild(toggle);
    return wrap;
  }

  /** SEARCH — "Load JSON File" opens the hidden file picker. */
  private buildSearchPanel(): HTMLElement {
    const wrap = document.createElement('div');
    wrap.className = 'pptx-toolbar__panel-info';

    const hint = document.createElement('span');
    hint.className = 'pptx-toolbar__panel-text';
    hint.textContent = 'Load a different PPTX JSON file — the viewer resets completely.';

    // Hidden native input, opened programmatically by the button.
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.json,application/json';
    fileInput.className = 'pptx-toolbar__file-input';

    // Reset after each pick so the SAME file can be re-selected later.
    fileInput.addEventListener('change', () => {
      const file = fileInput.files?.[0] ?? null;
      fileInput.value = '';
      if (file !== null && this.fileLoadedCallback !== null) {
        this.fileLoadedCallback(file);
      }
    });

    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'pptx-btn';
    button.textContent = '📂 Load JSON File';
    button.addEventListener('click', () => {
      this.callbacks.onFileSelected(); // "browse requested" notification
      fileInput.click();               // then open the OS picker
    });

    wrap.append(hint, button, fileInput);
    return wrap;
  }

  /** File picked via the Search tab — consumed by ViewerContainer. */
  private fileLoadedCallback: ((file: File) => void) | null = null;

  /** ViewerContainer registers here to receive Search-tab file picks. */
  onFileLoaded(callback: (file: File) => void): void {
    this.fileLoadedCallback = callback;
  }

  // ── Tab switching ────────────────────────────────────────────────────────

  private selectTab(id: TabId): void {
    for (const [tabId, tab] of this.tabs) {
      const isActive = tabId === id;
      tab.classList.toggle('is-active', isActive);
      tab.setAttribute('aria-selected', isActive ? 'true' : 'false');
    }

    const panelHost = this.root.querySelector<HTMLElement>('.pptx-toolbar__panel');
    if (panelHost !== null) {
      panelHost.replaceChildren(this.panels.get(id) as HTMLElement);
    }
  }
}