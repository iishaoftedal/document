import { describe, expect, it, vi } from 'vitest';
import { WALTERS_LOOK_CSS, injectWaltersLookCss, pushWaltersSkin } from '../../lib/onlyoffice/guards/walters-look';
import { DEFAULT_UI_THEME } from '../../lib/onlyoffice/ui-theme';
import { DARK_UI_THEME } from '../../lib/editor-theme';

/** The variables one theme block of the sheet sets, by name. */
function varsOf(selector: string): Map<string, string> {
  const start = WALTERS_LOOK_CSS.indexOf(`${selector} {`);
  expect(start, `${selector} block`).toBeGreaterThanOrEqual(0);
  const body = WALTERS_LOOK_CSS.slice(start, WALTERS_LOOK_CSS.indexOf('}', start));
  return new Map([...body.matchAll(/--([\w-]+): ([^;]+);/g)].map((m) => [m[1], m[2]]));
}

describe('guard 19: the Walters & Wally look', () => {
  it('recolours exactly the two themes the editor is mounted with', () => {
    expect(DEFAULT_UI_THEME).toBe('theme-white');
    expect(DARK_UI_THEME).toBe('theme-night');
    // `:root body.theme-x` outranks the vendor's `:root .theme-x`.
    varsOf(`:root body.${DEFAULT_UI_THEME}`);
    varsOf(`:root body.${DARK_UI_THEME}`);
  });

  it('sets the same variables by day and by night', () => {
    const day = [...varsOf(':root body.theme-white').keys()].sort();
    const night = [...varsOf(':root body.theme-night').keys()].sort();
    expect(night).toEqual(day);
  });

  it('keeps crimson by day and never lightens red by night', () => {
    const day = varsOf(':root body.theme-white');
    const night = varsOf(':root body.theme-night');
    expect(day.get('background-primary-dialog-button')).toBe('#a6121e');
    expect(day.get('highlight-toolbar-tab-underline-document')).toBe('#a6121e');
    expect(night.get('background-primary-dialog-button')).toBe('#7e1520');
    expect(night.get('highlight-toolbar-tab-underline-document')).toBe('#e9dfcb');
  });

  it('paints canvas colours in plain hex (the SDK parses them)', () => {
    for (const vars of [varsOf(':root body.theme-white'), varsOf(':root body.theme-night')]) {
      for (const [name, value] of vars) {
        if (name.startsWith('canvas-')) expect(value, name).toMatch(/^#[0-9a-f]{6}$/);
      }
    }
  });

  it('serves IBM Plex Sans from this site, not a font service', () => {
    expect(WALTERS_LOOK_CSS).toContain('url("/walters/plex-sans/ibm-plex-sans-latin-400-normal.woff2")');
    expect(WALTERS_LOOK_CSS).not.toMatch(/fonts\.(googleapis|gstatic)\.com/);
  });

  it('injects once per frame', () => {
    const doc = document.implementation.createHTMLDocument('frame');
    injectWaltersLookCss(doc);
    injectWaltersLookCss(doc);
    expect(doc.querySelectorAll('style#walters-look-css')).toHaveLength(1);
  });
});

describe('guard 19: handing the canvas colours to the SDK', () => {
  type Fake = Window & Record<string, unknown>;
  const frameWith = (extra: Record<string, unknown>): Fake => {
    const doc = document.implementation.createHTMLDocument('frame');
    return { document: doc, ...extra } as unknown as Fake;
  };

  it('waits for the sheet, the theme module and the SDK', () => {
    const asc_setSkin = vi.fn();
    const early = frameWith({});
    expect(pushWaltersSkin(early)).toBe(false);
    const noSdk = frameWith({ Common: { UI: { Themes: { getThemeColors: () => ({}) } } } });
    injectWaltersLookCss(noSdk.document);
    expect(pushWaltersSkin(noSdk)).toBe(false);
    const sheetMissing = frameWith({
      Common: { UI: { Themes: { getThemeColors: () => ({}) } } },
      editor: { asc_setSkin },
    });
    expect(pushWaltersSkin(sheetMissing)).toBe(false);
    expect(asc_setSkin).not.toHaveBeenCalled();
  });

  it('pushes the computed colours once', () => {
    const asc_setSkin = vi.fn();
    const colors = { 'canvas-background': '#f5f2eb', type: 'light', name: 'theme-white' };
    const win = frameWith({ Common: { UI: { Themes: { getThemeColors: () => colors } } }, editor: { asc_setSkin } });
    injectWaltersLookCss(win.document);
    expect(pushWaltersSkin(win)).toBe(true);
    expect(pushWaltersSkin(win)).toBe(true);
    expect(asc_setSkin).toHaveBeenCalledTimes(1);
    expect(asc_setSkin).toHaveBeenCalledWith(colors);
  });
});
