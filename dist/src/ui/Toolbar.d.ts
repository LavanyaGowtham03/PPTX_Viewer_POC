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
export declare class Toolbar {
    private readonly callbacks;
    private readonly root;
    private readonly tabs;
    private readonly panels;
    private readonly homeInfo;
    private animationsCheckbox;
    private transitionsCheckbox;
    constructor(callbacks: ToolbarCallbacks);
    /** Renders the toolbar: tab bar row + active-tab panel strip. */
    render(): HTMLElement;
    /** Sync the Animations toggle (e.g. after a full viewer reset). */
    setAnimationsState(enabled: boolean): void;
    /** Sync the Transitions toggle (e.g. after a full viewer reset). */
    setTransitionsState(enabled: boolean): void;
    /** Update the Home panel read-only info. Navigation-safe: only the
     *  text changes — the active tab is NEVER switched here (spec resets
     *  to Home only on a full viewer reload). */
    setSlideInfo(index: number, total: number, name: string): void;
    /** Programmatically activate a tab (used by the Search flow). */
    activateTab(tab: TabId): void;
    private build;
    private buildTab;
    /** HOME — current-slide info ("Slide X of N — name"), read-only. */
    private buildHomePanel;
    /** ANIMATIONS / TRANSITIONS — one toggle each, starting ON. */
    private buildTogglePanel;
    /** SEARCH — "Load JSON File" opens the hidden file picker. */
    private buildSearchPanel;
    /** File picked via the Search tab — consumed by ViewerContainer. */
    private fileLoadedCallback;
    /** ViewerContainer registers here to receive Search-tab file picks. */
    onFileLoaded(callback: (file: File) => void): void;
    private selectTab;
}
export {};
