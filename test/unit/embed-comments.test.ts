import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const { mockGetReadonlyMode } = vi.hoisted(() => ({ mockGetReadonlyMode: vi.fn().mockReturnValue(false) }));
vi.mock('../../lib/onlyoffice/readonly', () => ({ getReadonlyMode: mockGetReadonlyMode }));

import { addComment, CommentError } from '../../lib/embed-comments';
import { editorUserFrom, getEditorUser, GUEST, setEditorUser, USER_NAME_LIMIT } from '../../lib/onlyoffice/editor-user';

/** A stand-in for the Write or Slides editor: words to find, and the comments added. */
function textEditor(words: string, opts: { noFind?: boolean } = {}) {
  const added: Array<Record<string, unknown>> = [];
  let searched: { text: string; matchCase: boolean; whole: boolean } | null = null;
  class CSearchSettings {
    text = '';
    matchCase = false;
    whole = true;
    put_Text(t: string) {
      this.text = t;
    }
    put_MatchCase(on: boolean) {
      this.matchCase = on;
    }
    put_WholeWords(on: boolean) {
      this.whole = on;
    }
  }
  class asc_CCommentDataWord {
    fields: Record<string, unknown> = {};
    asc_putText(t: string) {
      this.fields.text = t;
    }
    asc_putUserName(n: string) {
      this.fields.author = n;
    }
    asc_putUserId(id: string) {
      this.fields.userId = id;
    }
    asc_putTime(t: string) {
      this.fields.time = t;
    }
    asc_putSolved(s: boolean) {
      this.fields.solved = s;
    }
  }
  const api: Record<string, unknown> = {
    asc_addComment: vi.fn((data: asc_CCommentDataWord) => added.push({ ...data.fields, on: searched?.text ?? null })),
    asc_endFindText: vi.fn(),
    asc_RemoveSelection: vi.fn(),
  };
  if (!opts.noFind) {
    api.asc_findText = vi.fn((s: CSearchSettings) => {
      searched = { text: s.text, matchCase: s.matchCase, whole: s.whole };
      return words.split(s.text).length - 1;
    });
  }
  return { api, added, searched: () => searched, Asc: { asc_CCommentDataWord }, AscCommon: { CSearchSettings } };
}

let frame: HTMLIFrameElement;
function mount(editor: { api: unknown; Asc: unknown; AscCommon: unknown }) {
  frame = document.createElement('iframe');
  frame.name = 'frameEditor';
  document.body.appendChild(frame);
  const win = frame.contentWindow as unknown as Record<string, unknown>;
  const wb = (editor.api as { wb?: unknown }).wb;
  if (wb) win.Asc = { ...(editor.Asc as object), editor: editor.api };
  else {
    win.editor = editor.api;
    win.Asc = editor.Asc;
  }
  win.AscCommon = editor.AscCommon;
  (window as unknown as { editor?: unknown }).editor = {};
}

const LEASE = 'The rent is $2,400 a month. The tenant keeps the walk clear. The tenant keeps the walk clear of ice.';

describe('editor-user (added in this copy)', () => {
  afterEach(() => setEditorUser(null));

  it('opens as Guest when nobody is named', () => {
    expect(getEditorUser()).toEqual({ id: 'local-user', name: 'Guest' });
    expect(editorUserFrom(undefined)).toBeNull();
    expect(editorUserFrom({ name: '   ' })).toBeNull();
    expect(editorUserFrom('Sam')).toBeNull();
  });

  it('takes a name and an id, as one plain line', () => {
    expect(editorUserFrom({ id: 'owner-1', name: '  Sam\n Rivera\t' })).toEqual({ id: 'owner-1', name: 'Sam Rivera' });
    expect(editorUserFrom({ name: 'Summit Roofing' })).toEqual({ id: 'user-summit-roofing', name: 'Summit Roofing' });
    expect(editorUserFrom({ id: 'has spaces and <tags>', name: 'Pat' })?.id).toBe('user-pat');
    expect(editorUserFrom({ name: 'x'.repeat(200) })?.name).toHaveLength(USER_NAME_LIMIT);
  });

  it('keeps the person for the next open, and goes back to Guest after an open with nobody named', () => {
    setEditorUser(editorUserFrom({ id: 'o1', name: 'Sam Rivera' }));
    expect(getEditorUser()).toEqual({ id: 'o1', name: 'Sam Rivera' });
    setEditorUser(editorUserFrom(undefined));
    expect(getEditorUser()).toBe(GUEST);
  });
});

describe('embed-comments (added in this copy)', () => {
  beforeEach(() => {
    mockGetReadonlyMode.mockReturnValue(false);
    setEditorUser(editorUserFrom({ id: 'owner-1', name: 'Sam Rivera' }));
  });
  afterEach(() => {
    frame?.remove();
    delete (window as unknown as { editor?: unknown }).editor;
    setEditorUser(null);
  });

  it('puts a comment on words that are there once, found letter for letter', () => {
    const editor = textEditor(LEASE);
    mount(editor);
    const done = addComment({
      text: 'Check this matches the books.',
      author: 'Wally',
      quote: 'The rent is $2,400 a month',
    });
    expect(done).toEqual({ quote: 'The rent is $2,400 a month', author: 'Wally' });
    expect(editor.searched()).toEqual({ text: 'The rent is $2,400 a month', matchCase: true, whole: false });
    expect(editor.added).toHaveLength(1);
    expect(editor.added[0]).toMatchObject({
      text: 'Check this matches the books.',
      author: 'Wally',
      userId: 'author-wally',
      solved: false,
      on: 'The rent is $2,400 a month',
    });
    expect(typeof editor.added[0]!.time).toBe('string');
    // The found words are let go: nothing is left picked.
    expect(editor.api.asc_RemoveSelection).toHaveBeenCalledTimes(1);
  });

  it('adds nothing when the words are not there, or are there more than once', () => {
    const editor = textEditor(LEASE);
    mount(editor);
    expect(() => addComment({ text: 'x', quote: 'The rent is $2,600' })).toThrow(
      expect.objectContaining({ code: 'not-found' }),
    );
    expect(() => addComment({ text: 'x', quote: 'The tenant keeps the walk clear' })).toThrow(
      expect.objectContaining({ code: 'not-once' }),
    );
    expect(editor.added).toHaveLength(0);
    expect(editor.api.asc_endFindText).toHaveBeenCalledTimes(2);
  });

  it('with no quote, goes on what is picked, signed by the person the document opened as', () => {
    const editor = textEditor(LEASE);
    mount(editor);
    expect(addComment({ text: 'Looks right.' })).toEqual({ author: 'Sam Rivera' });
    expect(editor.api.asc_findText).not.toHaveBeenCalled();
    // What the person picked stays picked.
    expect(editor.api.asc_RemoveSelection).not.toHaveBeenCalled();
    expect(editor.added[0]).toMatchObject({ text: 'Looks right.', author: 'Sam Rivera', userId: 'owner-1' });
  });

  it('refuses a read-only document, an empty comment, a bad quote, and no document', () => {
    const editor = textEditor(LEASE);
    expect(() => addComment({ text: 'x' })).toThrow(expect.objectContaining({ code: 'no-document' }));
    mount(editor);
    expect(() => addComment({ text: '   ' })).toThrow(expect.objectContaining({ code: 'bad-request' }));
    expect(() => addComment({ text: 'x', quote: '  ' })).toThrow(expect.objectContaining({ code: 'bad-request' }));
    expect(() => addComment({ text: 'x', quote: 'y'.repeat(1001) })).toThrow(
      expect.objectContaining({ code: 'bad-request' }),
    );
    mockGetReadonlyMode.mockReturnValue(true);
    expect(() => addComment({ text: 'x' })).toThrow(expect.objectContaining({ code: 'read-only' }));
    expect(editor.added).toHaveLength(0);
  });

  it('a spreadsheet takes a comment on the picked cell, never a quote', () => {
    const editor = textEditor('');
    const api = { ...editor.api, wb: { getWorksheet: () => ({}) } };
    mount({ ...editor, api, Asc: { asc_CCommentData: editor.Asc.asc_CCommentDataWord } });
    expect(() => addComment({ text: 'x', quote: 'B2' })).toThrow(expect.objectContaining({ code: 'not-supported' }));
    expect(addComment({ text: 'Is this hours or days?' })).toEqual({ author: 'Sam Rivera' });
    expect(editor.added).toHaveLength(1);
  });

  it('says so when the editor cannot look for words', () => {
    mount(textEditor(LEASE, { noFind: true }));
    expect(() => addComment({ text: 'x', quote: 'rent' })).toThrow(CommentError);
  });
});
