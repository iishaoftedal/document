/**
 * Guard 19: Walters & Wally's look on the editor.
 *
 * The editor runs on the vendor's two modern interface themes, `theme-white`
 * by day and `theme-night` by night (see ../ui-theme.ts and ../../editor-theme.ts).
 * Both are nothing but CSS variables on `:root .theme-…`, so this sheet sets
 * the same variables one step more specifically and the whole interface takes
 * the brand: warm paper and ink by day, graphite by night, crimson only on the
 * one solid button and the active tab by day, cream for those by night (red is
 * never lightened at night; solid buttons go oxblood), and IBM Plex Sans for
 * every label, served from this site (`public/walters/plex-sans`, SIL OFL 1.1).
 *
 * Colour values are copied from the brand kit's `src/brand-tokens.css` in the
 * Wally and portal repos. If a value changes there, change it here too.
 * `canvas-*` values stay plain hex: the SDK reads them from the stylesheet to
 * paint the rulers, scroll bars and the ground around the page.
 *
 * The SDK reads those canvas colours once, when the theme is applied at boot,
 * which is before this sheet can land. So once the sheet is in, the colours
 * are handed to the SDK again the way the vendor does on a theme switch
 * (`Common.UI.Themes.getThemeColors()` into `asc_setSkin`). Later Day/Night
 * switches read the stylesheet with this sheet already in it.
 */

const FONT = '"IBM Plex Sans", "Segoe UI", Arial, sans-serif';

const DAY: Record<string, string> = {
  // Ground and sheets
  'background-normal': '#ffffff',
  'background-toolbar': '#ffffff',
  'background-toolbar-tab': '#ffffff',
  'background-toolbar-additional': '#faf8f3',
  'background-contrast-popover': '#ffffff',
  'background-pane': '#faf8f3',
  'background-statusbar': '#ffffff',
  'border-toolbar-active-panel-top': '#faf8f3',
  'toolbar-header-document': '#ffffff',
  'toolbar-header-spreadsheet': '#ffffff',
  'toolbar-header-presentation': '#ffffff',
  'toolbar-header-pdf': '#ffffff',
  'toolbar-header-visio': '#ffffff',
  // The one solid button: crimson, darker under the pointer
  'background-primary-dialog-button': '#a6121e',
  'background-accent-button': '#a6121e',
  'highlight-primary-dialog-button-hover': '#8f0f1a',
  'highlight-primary-dialog-button-pressed': '#7e1520',
  // Quiet fills: warm grey to hover, the crimson wash for "on"
  'highlight-button-hover': '#efebe3',
  'highlight-button-pressed': '#f6e7e8',
  'highlight-button-pressed-hover': '#efd9db',
  'highlight-header-button-hover': '#efebe3',
  'highlight-header-button-pressed': '#f6e7e8',
  'highlight-comment-hover': '#faf8f3',
  'highlight-comment-pressed': '#f6e7e8',
  'highlight-category-button-hover': 'rgba(166, 18, 30, 0.05)',
  'highlight-category-button-pressed': 'rgba(166, 18, 30, 0.1)',
  'highlight-text-select': '#a6121e',
  'chb-background-normal-hover': '#faf8f3',
  'chb-background-checked-hover': '#faf8f3',
  'rb-background-normal-hover': '#faf8f3',
  'rb-background-checked-hover': '#faf8f3',
  // Lines
  'border-toolbar': '#e3dfd7',
  'border-divider': '#ece8e0',
  'border-regular-control': '#ddd8ce',
  'border-preview-hover': '#d9a0a5',
  // Selected and focused: crimson
  'border-preview-select': '#a6121e',
  'border-control-focus': '#a6121e',
  'border-button-pressed-focus': '#a6121e',
  'border-fill-input-focused': '#a6121e',
  'chb-border-normal-focus': '#a6121e',
  'chb-border-checked-focus': '#a6121e',
  'rb-border-normal-focus': '#a6121e',
  'rb-border-checked-focus': '#a6121e',
  'slider-track-background-filled': '#a6121e',
  'slider-thumb-background-normal': '#a6121e',
  'slider-thumb-background-hover': '#8f0f1a',
  'slider-thumb-background-active': '#7e1520',
  'highlight-toolbar-tab-underline-document': '#a6121e',
  'highlight-toolbar-tab-underline-spreadsheet': '#a6121e',
  'highlight-toolbar-tab-underline-presentation': '#a6121e',
  'highlight-toolbar-tab-underline-pdf': '#a6121e',
  'highlight-toolbar-tab-underline-visio': '#a6121e',
  'highlight-header-tab-underline-document': '#a6121e',
  'highlight-header-tab-underline-spreadsheet': '#a6121e',
  'highlight-header-tab-underline-presentation': '#a6121e',
  'highlight-header-tab-underline-pdf': '#a6121e',
  'highlight-header-tab-underline-visio': '#a6121e',
  // Ink
  'text-normal': '#1f1d1a',
  'text-normal-pressed': '#1f1d1a',
  'text-secondary': '#6f6960',
  'text-tertiary': '#8f897f',
  'text-toolbar-header': '#1f1d1a',
  'text-link': '#a6121e',
  'text-link-hover': '#a6121e',
  'text-link-active': '#a6121e',
  'text-link-visited': '#a6121e',
  // Around the page
  'canvas-background': '#f5f2eb',
  'canvas-page-border': '#e0dbd1',
  'canvas-ruler-background': '#ffffff',
  'canvas-ruler-border': '#e3dfd7',
  'canvas-ruler-margins-background': '#efebe3',
  'canvas-scroll-thumb': '#faf8f3',
  'canvas-scroll-thumb-hover': '#e3dfd7',
  'canvas-scroll-thumb-pressed': '#d3cdc2',
  'canvas-scroll-thumb-border': '#e0dbd1',
  'canvas-scroll-thumb-border-hover': '#d3cdc2',
  'canvas-scroll-thumb-border-pressed': '#c4bdb1',
  'canvas-cell-title-background': '#faf8f3',
  'canvas-cell-title-background-hover': '#efebe3',
  'canvas-cell-title-background-selected': '#f6e7e8',
  'canvas-cell-title-border': '#e3dfd7',
  'canvas-cell-title-border-hover': '#d3cdc2',
  'canvas-cell-title-border-selected': '#a6121e',
  // Type
  'font-family-base': FONT,
};

const NIGHT: Record<string, string> = {
  // Ground and sheets: graphite
  'background-normal': '#474441',
  'background-toolbar': '#474441',
  'background-toolbar-tab': '#474441',
  'background-toolbar-additional': '#403e3b',
  'background-contrast-popover': '#474441',
  'background-pane': '#403e3b',
  'background-statusbar': '#474441',
  'border-toolbar-active-panel-top': '#403e3b',
  'toolbar-header-document': '#474441',
  'toolbar-header-spreadsheet': '#474441',
  'toolbar-header-presentation': '#474441',
  'toolbar-header-pdf': '#474441',
  'toolbar-header-visio': '#474441',
  // The one solid button: oxblood, never lightened
  'background-primary-dialog-button': '#7e1520',
  'background-accent-button': '#7e1520',
  'highlight-primary-dialog-button-hover': '#6e121c',
  'highlight-primary-dialog-button-pressed': '#5f0f18',
  // Quiet fills: a lighter graphite to hover, the cream wash for "on"
  'highlight-button-hover': '#4f4b47',
  'highlight-button-pressed': '#57534f',
  'highlight-button-pressed-hover': '#605c57',
  'highlight-header-button-hover': '#4f4b47',
  'highlight-header-button-pressed': '#57534f',
  'highlight-comment-hover': '#4f4b47',
  'highlight-comment-pressed': '#57534f',
  'highlight-category-button-hover': 'rgba(233, 223, 203, 0.05)',
  'highlight-category-button-pressed': 'rgba(233, 223, 203, 0.1)',
  'highlight-text-select': '#7e1520',
  'chb-background-normal-hover': '#4f4b47',
  'chb-background-checked-hover': '#4f4b47',
  'rb-background-normal-hover': '#4f4b47',
  'rb-background-checked-hover': '#4f4b47',
  // Lines
  'border-toolbar': '#57534f',
  'border-divider': '#504c48',
  'border-regular-control': '#5f5b56',
  'border-preview-hover': '#a39d93',
  // Selected and focused: cream
  'border-preview-select': '#e9dfcb',
  'border-control-focus': '#e9dfcb',
  'border-button-pressed-focus': '#e9dfcb',
  'border-fill-input-focused': '#e9dfcb',
  'chb-border-normal-focus': '#e9dfcb',
  'chb-border-checked-focus': '#e9dfcb',
  'rb-border-normal-focus': '#e9dfcb',
  'rb-border-checked-focus': '#e9dfcb',
  'slider-track-background-filled': '#e9dfcb',
  'slider-thumb-background-normal': '#e9dfcb',
  'slider-thumb-background-hover': '#e9dfcb',
  'slider-thumb-background-active': '#e9dfcb',
  'highlight-toolbar-tab-underline-document': '#e9dfcb',
  'highlight-toolbar-tab-underline-spreadsheet': '#e9dfcb',
  'highlight-toolbar-tab-underline-presentation': '#e9dfcb',
  'highlight-toolbar-tab-underline-pdf': '#e9dfcb',
  'highlight-toolbar-tab-underline-visio': '#e9dfcb',
  'highlight-header-tab-underline-document': '#e9dfcb',
  'highlight-header-tab-underline-spreadsheet': '#e9dfcb',
  'highlight-header-tab-underline-presentation': '#e9dfcb',
  'highlight-header-tab-underline-pdf': '#e9dfcb',
  'highlight-header-tab-underline-visio': '#e9dfcb',
  // Ink
  'text-normal': '#f1ede6',
  'text-normal-pressed': '#f1ede6',
  'text-secondary': '#c2bcb1',
  'text-tertiary': '#a39d93',
  'text-toolbar-header': '#f1ede6',
  'text-link': '#e9dfcb',
  'text-link-hover': '#e9dfcb',
  'text-link-active': '#e9dfcb',
  'text-link-visited': '#e9dfcb',
  // Around the page
  'canvas-background': '#3a3835',
  'canvas-page-border': '#57534f',
  'canvas-ruler-background': '#403e3b',
  'canvas-ruler-border': '#57534f',
  'canvas-ruler-margins-background': '#4f4b47',
  'canvas-scroll-thumb': '#4f4b47',
  'canvas-scroll-thumb-hover': '#5f5b56',
  'canvas-scroll-thumb-pressed': '#6b6661',
  'canvas-scroll-thumb-border': '#57534f',
  'canvas-scroll-thumb-border-hover': '#6b6661',
  'canvas-scroll-thumb-border-pressed': '#a39d93',
  'canvas-cell-title-background': '#403e3b',
  'canvas-cell-title-background-hover': '#4f4b47',
  'canvas-cell-title-background-selected': '#57534f',
  'canvas-cell-title-border': '#57534f',
  'canvas-cell-title-border-hover': '#6b6661',
  'canvas-cell-title-border-selected': '#e9dfcb',
  // Type
  'font-family-base': FONT,
};

const block = (selector: string, vars: Record<string, string>): string =>
  `${selector} {\n${Object.entries(vars)
    .map(([name, value]) => `  --${name}: ${value};`)
    .join('\n')}\n}`;

const face = (weight: number): string =>
  `@font-face { font-family: "IBM Plex Sans"; font-style: normal; font-weight: ${weight}; font-display: swap; ` +
  `src: url("/walters/plex-sans/ibm-plex-sans-latin-${weight}-normal.woff2") format("woff2"); }`;

/** The whole sheet. `:root body.theme-…` outranks the vendor's `:root .theme-…`. */
export const WALTERS_LOOK_CSS = [
  ...[400, 500, 600, 700].map(face),
  block(':root body.theme-white', DAY),
  block(':root body.theme-night', NIGHT),
  // Labels the vendor sets in its own font stack rather than the variable.
  'body, .btn, .btn-text-default, .dropdown-menu, .tooltip, input, textarea, select, ' +
    'label, .ribtab > a, .combobox, .listview, .asc-window { font-family: var(--font-family-base); }',
].join('\n');

type LookFrame = Window & {
  Common?: { UI?: { Themes?: { getThemeColors?: () => Record<string, string> } } };
  Asc?: { editor?: { asc_setSkin?: (colors: Record<string, string>) => void } };
  editor?: { asc_setSkin?: (colors: Record<string, string>) => void };
  __waltersSkinPushed?: boolean;
};

/**
 * Hands the recoloured canvas colours to the SDK, once per frame. Returns
 * true once done; until the SDK and the theme module have both landed it
 * does nothing and reports false, so the caller's retry timer brings it back.
 */
export function pushWaltersSkin(win: Window): boolean {
  const frame = win as LookFrame;
  if (frame.__waltersSkinPushed) return true;
  if (!frame.document?.getElementById('walters-look-css')) return false;
  const colors = frame.Common?.UI?.Themes?.getThemeColors?.();
  const api = frame.Asc?.editor ?? frame.editor;
  if (!colors || typeof api?.asc_setSkin !== 'function') return false;
  try {
    api.asc_setSkin(colors);
    frame.__waltersSkinPushed = true;
  } catch {
    // A frame mid-boot can throw; the next pass tries again.
  }
  return frame.__waltersSkinPushed === true;
}

/** Puts the sheet into an editor frame. Idempotent. */
export function injectWaltersLookCss(doc: Document): void {
  if (doc.getElementById('walters-look-css')) return;
  const style = doc.createElement('style');
  style.id = 'walters-look-css';
  style.textContent = WALTERS_LOOK_CSS;
  (doc.head || doc.documentElement).appendChild(style);
}
