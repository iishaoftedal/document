import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const { mockGetReadonlyMode } = vi.hoisted(() => ({ mockGetReadonlyMode: vi.fn().mockReturnValue(false) }));
vi.mock('../../lib/onlyoffice/readonly', () => ({ getReadonlyMode: mockGetReadonlyMode }));

import {
  readSelection,
  replaceSelection,
  SELECTION_LETTER_LIMIT,
  SelectionError,
  watchSelection,
} from '../../lib/embed-selection';

/** A stand-in for the Write/Slides editor: picked text, a look, and AddTextWithPr. */
function wordEditor(initial: { text: string; type?: string; slides?: boolean }) {
  const state = { text: initial.text, type: initial.type ?? (initial.text ? 'text' : 'none') };
  const added: Array<{ text: string; look: unknown }> = [];
  const callbacks = new Map<string, () => void>();
  class CAddTextSettings {
    look: unknown;
    SetTextPr(pr: unknown) {
      this.look = pr;
    }
  }
  const logic = {
    ...(initial.slides ? { Slides: [{}] } : {}),
    GetDirectTextPr: () => ({ Bold: true, Copy: () => ({ Bold: true, copied: true }) }),
    AddTextWithPr: vi.fn((text: string, settings: CAddTextSettings) => {
      added.push({ text, look: settings.look });
      state.text = '';
      state.type = 'none';
    }),
    MoveCursorLeft: vi.fn(() => {
      state.text = state.text.slice(0, -2);
    }),
    Document_UpdateSelectionState: vi.fn(),
  };
  const api = {
    WordControl: { m_oLogicDocument: logic },
    pluginMethod_GetSelectedText: () => state.text,
    pluginMethod_GetSelectionType: () => state.type,
    asc_registerCallback: (event: string, cb: () => void) => callbacks.set(event, cb),
  };
  return { api, logic, state, added, callbacks, Asc: {}, AscCommon: { CAddTextSettings } };
}

function sheetEditor(ranges: Array<{ r1: number; c1: number; r2: number; c2: number }>, editing = false) {
  return {
    api: {
      wb: {
        getCellEditMode: () => editing,
        getWorksheet: () => ({ model: { selectionRange: { ranges } } }),
      },
      pluginMethod_GetSelectedText: () => '',
      pluginMethod_GetSelectionType: () => 'text',
    },
    Asc: {},
    AscCommon: {},
  };
}

let frame: HTMLIFrameElement;
function mount(editor: { api: unknown; Asc: unknown; AscCommon: unknown }) {
  frame = document.createElement('iframe');
  frame.name = 'frameEditor';
  document.body.appendChild(frame);
  const win = frame.contentWindow as unknown as Record<string, unknown>;
  // The spreadsheet editor has no window.editor, only Asc.editor.
  if ((editor.api as { wb?: unknown }).wb) win.Asc = { ...(editor.Asc as object), editor: editor.api };
  else {
    win.editor = editor.api;
    win.Asc = editor.Asc;
  }
  win.AscCommon = editor.AscCommon;
  return win;
}

describe('embed-selection (added in this copy)', () => {
  beforeEach(() => {
    mockGetReadonlyMode.mockReturnValue(false);
  });
  afterEach(() => {
    frame?.remove();
    vi.useRealTimers();
  });

  describe('readSelection', () => {
    it('says nothing is picked when no editor is mounted', () => {
      expect(readSelection()).toEqual({ kind: 'none', letters: 0 });
    });

    it('gives the picked words with plain line ends', () => {
      mount(wordEditor({ text: 'One sentence.\r\nNext one.' }));
      expect(readSelection()).toEqual({ kind: 'text', letters: 23, text: 'One sentence.\nNext one.', paragraphs: 2 });
    });

    it('takes the kind from the words on a slide, which reports "slide"', () => {
      mount(wordEditor({ text: 'Spring', type: 'slide', slides: true }));
      expect(readSelection()).toMatchObject({ kind: 'text', text: 'Spring', letters: 6 });
      frame.remove();
      mount(wordEditor({ text: '', type: 'slide', slides: true }));
      expect(readSelection()).toEqual({ kind: 'none', letters: 0 });
    });

    it('calls a picked picture "other"', () => {
      mount(wordEditor({ text: '', type: 'drawing' }));
      expect(readSelection()).toEqual({ kind: 'other', letters: 0 });
    });

    it(`leaves the words out past ${SELECTION_LETTER_LIMIT} letters and says so`, () => {
      mount(wordEditor({ text: 'a'.repeat(SELECTION_LETTER_LIMIT + 1) }));
      const picked = readSelection();
      expect(picked).toMatchObject({ kind: 'text', letters: SELECTION_LETTER_LIMIT + 1, tooLong: true });
      expect(picked.text).toBeUndefined();
    });

    it('names a picked block of cells by its range', () => {
      mount(sheetEditor([{ r1: 1, c1: 1, r2: 13, c2: 1 }]));
      expect(readSelection()).toEqual({ kind: 'cells', letters: 0, cells: 13, range: 'B2:B14' });
    });

    it('names a block picked from the bottom up the same way', () => {
      mount(sheetEditor([{ r1: 4, c1: 27, r2: 2, c2: 26 }]));
      expect(readSelection()).toMatchObject({ kind: 'cells', range: 'AA3:AB5', cells: 6 });
    });

    it('treats the one cell a sheet always has as nothing picked', () => {
      mount(sheetEditor([{ r1: 0, c1: 0, r2: 0, c2: 0 }]));
      expect(readSelection()).toEqual({ kind: 'none', letters: 0 });
    });

    it('calls separate blocks, or typing in a cell, "other"', () => {
      mount(
        sheetEditor([
          { r1: 0, c1: 0, r2: 1, c2: 0 },
          { r1: 5, c1: 5, r2: 6, c2: 5 },
        ]),
      );
      expect(readSelection().kind).toBe('other');
      frame.remove();
      mount(sheetEditor([{ r1: 0, c1: 0, r2: 3, c2: 0 }], true));
      expect(readSelection().kind).toBe('other');
    });
  });

  describe('replaceSelection', () => {
    it('writes the new words with the look of the picked ones, through the editor', () => {
      const editor = wordEditor({ text: 'This bold sentence is far too long.' });
      mount(editor);
      expect(replaceSelection({ lines: ['Short and bold.'], expect: 'This bold sentence is far too long.' })).toEqual({
        letters: 15,
        text: 'Short and bold.',
      });
      expect(editor.added).toEqual([{ text: 'Short and bold.', look: { Bold: true, copied: true } }]);
    });

    it('changes nothing when the picked words are not the ones read', () => {
      const editor = wordEditor({ text: 'Something else' });
      mount(editor);
      expect(() => replaceSelection({ lines: ['New'], expect: 'Old words' })).toThrow(SelectionError);
      try {
        replaceSelection({ lines: ['New'], expect: 'Old words' });
      } catch (error) {
        expect((error as SelectionError).code).toBe('selection-changed');
      }
      expect(editor.logic.AddTextWithPr).not.toHaveBeenCalled();
    });

    it('leaves out the paragraph end of a whole picked paragraph', () => {
      const editor = wordEditor({ text: 'A whole paragraph.\r\n' });
      mount(editor);
      replaceSelection({ lines: ['Shorter.'], expect: 'A whole paragraph.\n' });
      expect(editor.logic.MoveCursorLeft).toHaveBeenCalledWith(true, false);
      expect(editor.added.map((a) => a.text)).toEqual(['Shorter.']);
    });

    it('refuses words that reach into a second paragraph, so none is merged away', () => {
      const editor = wordEditor({ text: 'First.\r\nSecond.' });
      mount(editor);
      expect(() => replaceSelection({ lines: ['One.'], expect: 'First.\nSecond.' })).toThrow(/one paragraph/);
      expect(editor.logic.AddTextWithPr).not.toHaveBeenCalled();
    });

    it('refuses cells, a read-only document, and bad requests', () => {
      mount(sheetEditor([{ r1: 0, c1: 0, r2: 3, c2: 0 }]));
      expect(() => replaceSelection({ lines: ['x'], expect: '' })).toThrow(/Cells/);
      frame.remove();
      const editor = wordEditor({ text: 'Words' });
      mount(editor);
      mockGetReadonlyMode.mockReturnValue(true);
      expect(() => replaceSelection({ lines: ['x'], expect: 'Words' })).toThrow(/read-only/);
      mockGetReadonlyMode.mockReturnValue(false);
      expect(() => replaceSelection({ lines: [], expect: 'Words' })).toThrow(SelectionError);
      expect(() => replaceSelection({ lines: ['   '], expect: 'Words' })).toThrow(SelectionError);
      expect(() => replaceSelection({ lines: [7], expect: 'Words' })).toThrow(SelectionError);
      expect(() => replaceSelection({ lines: ['x'] })).toThrow(SelectionError);
      expect(editor.logic.AddTextWithPr).not.toHaveBeenCalled();
    });

    it('joins several lines into one paragraph with spaces', () => {
      const editor = wordEditor({ text: 'Old' });
      mount(editor);
      replaceSelection({ lines: ['One', 'two\nthree'], expect: 'Old' });
      expect(editor.added[0].text).toBe('One two three');
    });
  });

  describe('watchSelection', () => {
    it('pushes only the kind and the count, only on a change, at most four times a second', () => {
      vi.useFakeTimers();
      const editor = wordEditor({ text: '' });
      mount(editor);
      const push = vi.fn();
      const stop = watchSelection(push, { slowCheckMs: 60_000 });
      vi.advanceTimersByTime(0);
      // Nothing is checked until something happens or the slow check runs.
      editor.callbacks.get('asc_onSelectionChanged')?.();
      vi.advanceTimersByTime(0);
      expect(push).toHaveBeenLastCalledWith({ kind: 'none', letters: 0 });
      editor.state.text = 'Hello there';
      editor.state.type = 'text';
      editor.callbacks.get('asc_onFocusObject')?.();
      editor.callbacks.get('asc_onFocusObject')?.();
      vi.advanceTimersByTime(100);
      expect(push).toHaveBeenCalledTimes(1);
      vi.advanceTimersByTime(150);
      expect(push).toHaveBeenCalledTimes(2);
      expect(push).toHaveBeenLastCalledWith({ kind: 'text', letters: 11 });
      // Same pick again: nothing pushed.
      editor.callbacks.get('asc_onFocusObject')?.();
      vi.advanceTimersByTime(300);
      expect(push).toHaveBeenCalledTimes(2);
      stop();
    });

    it('hears the pointer and keys inside the editor frame', () => {
      vi.useFakeTimers();
      const editor = wordEditor({ text: 'Picked' });
      const win = mount(editor) as unknown as Window & { Event: typeof Event };
      const push = vi.fn();
      const stop = watchSelection(push, { slowCheckMs: 60_000 });
      win.dispatchEvent(new win.Event('mouseup'));
      vi.advanceTimersByTime(0);
      expect(push).toHaveBeenLastCalledWith({ kind: 'text', letters: 6 });
      editor.state.text = '';
      editor.state.type = 'none';
      win.dispatchEvent(new win.Event('keyup'));
      vi.advanceTimersByTime(250);
      expect(push).toHaveBeenLastCalledWith({ kind: 'none', letters: 0 });
      stop();
    });

    it('finds the editor later with its slow check', () => {
      vi.useFakeTimers();
      const push = vi.fn();
      const stop = watchSelection(push, { slowCheckMs: 1000 });
      vi.advanceTimersByTime(1001);
      expect(push).toHaveBeenLastCalledWith({ kind: 'none', letters: 0 });
      mount(wordEditor({ text: 'Late' }));
      vi.advanceTimersByTime(1001);
      expect(push).toHaveBeenLastCalledWith({ kind: 'text', letters: 4 });
      stop();
    });
  });
});
