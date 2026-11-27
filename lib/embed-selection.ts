/**
 * The picked words, for the page that embeds the editor. Added in this copy
 * (iishaoftedal/document), 2026-11-26.
 *
 * Three things, all reached through the embed API's postMessage protocol:
 * - `readSelection()` answers `document:get-selection`: what is picked, with the
 *   words (or, in a spreadsheet, the range) when there are not too many.
 * - `watchSelection()` pushes `document:selection-changed` when what is picked
 *   changes: only its kind and how many letters, never the words themselves.
 * - `replaceSelection()` answers `document:replace-selection`: it writes new
 *   words over the picked ones through the editor's own add-text action, with
 *   the look of the first picked letter, as one step the editor's Undo takes back.
 *   It changes nothing unless the picked words are still the ones the host read.
 *
 * Why not `pluginMethod_ReplaceTextSmart`: in this build it throws (the native
 * method it wraps is absent) and leaves an empty step in the undo history.
 * `pluginMethod_InputText` / `pluginMethod_PasteText` do replace the words but
 * drop their look (a bold sentence comes back plain), because the new run takes
 * the formatting of the text before the selection.
 */
import { getEditorApi, getEditorContext, type EditorApi } from './agent-plugin/editor-bridge';
import { getReadonlyMode } from './onlyoffice/readonly';

/** Most letters handed over in one answer. Past this the words are not sent. */
export const SELECTION_LETTER_LIMIT = 6000;
/** Pushes of `document:selection-changed` are at least this far apart (four a second). */
export const SELECTION_PUSH_GAP_MS = 250;

export type SelectionKind = 'none' | 'text' | 'cells' | 'other';

export interface PickedSelection {
  kind: SelectionKind;
  /** Letters picked (for cells: 0). */
  letters: number;
  /** The picked words, when kind is "text" and there are at most SELECTION_LETTER_LIMIT letters. */
  text?: string;
  /** True when the words were left out because there are too many. */
  tooLong?: boolean;
  /** Paragraphs the picked words reach into (kind "text"). */
  paragraphs?: number;
  /** For cells: the picked range, like "B2:B14". */
  range?: string;
  /** For cells: how many cells are picked. */
  cells?: number;
}

/** An error the host can tell apart by its `code`. */
export class SelectionError extends Error {
  constructor(
    readonly code:
      'selection-changed' | 'no-selection' | 'not-supported' | 'read-only' | 'bad-request' | 'not-verified',
    message: string,
  ) {
    super(message);
    this.name = 'SelectionError';
  }
}

interface SpreadsheetRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

interface SpreadsheetApi {
  wb?: {
    getCellEditMode?(): boolean;
    getWorksheet?(): { model?: { selectionRange?: { ranges?: SpreadsheetRange[]; getLast?(): SpreadsheetRange } } };
  };
  asc_getActiveRangeStr?(referenceType: number): string;
}

interface TextPr {
  Copy(): TextPr;
}

interface WordLogic {
  Slides?: unknown[];
  GetDirectTextPr?(): TextPr | null | undefined;
  AddTextWithPr?(text: string, settings: unknown): void;
  IsSelectionUse?(): boolean;
  MoveCursorLeft?(addToSelect: boolean, word: boolean): void;
  Document_UpdateSelectionState?(): void;
}

function logicOf(api: EditorApi): WordLogic | undefined {
  return (api.WordControl as { m_oLogicDocument?: WordLogic } | undefined)?.m_oLogicDocument;
}

function isSpreadsheet(api: EditorApi): boolean {
  const wb = (api as unknown as SpreadsheetApi).wb;
  return !!wb && typeof wb === 'object' && typeof wb.getWorksheet === 'function';
}

function columnName(index: number): string {
  let name = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}

function rangeName(range: SpreadsheetRange): string {
  const start = `${columnName(range.c1)}${range.r1 + 1}`;
  const end = `${columnName(range.c2)}${range.r2 + 1}`;
  return start === end ? start : `${start}:${end}`;
}

function readCells(api: EditorApi): PickedSelection {
  const sheet = api as unknown as SpreadsheetApi;
  // While a cell is being typed into, the picked thing is the typing, not cells.
  if (sheet.wb?.getCellEditMode?.()) return { kind: 'other', letters: 0 };
  const selection = sheet.wb?.getWorksheet?.()?.model?.selectionRange;
  const ranges = selection?.ranges;
  if (!ranges?.length) return { kind: 'none', letters: 0 };
  // Several separate blocks (Ctrl+click) are not one range a host can name.
  if (ranges.length > 1) return { kind: 'other', letters: 0 };
  const range = ranges[0];
  const r1 = Math.min(range.r1, range.r2);
  const r2 = Math.max(range.r1, range.r2);
  const c1 = Math.min(range.c1, range.c2);
  const c2 = Math.max(range.c1, range.c2);
  const cells = (r2 - r1 + 1) * (c2 - c1 + 1);
  // A sheet always has one cell picked: that is the cursor, not a pick.
  if (!(cells > 1)) return { kind: 'none', letters: 0 };
  return { kind: 'cells', letters: 0, cells, range: rangeName({ r1, c1, r2, c2 }) };
}

function readWords(api: EditorApi): PickedSelection {
  let text = '';
  let type = 'none';
  try {
    text = api.pluginMethod_GetSelectedText({ TabSymbol: '\t', Numbering: false }) || '';
    type = api.pluginMethod_GetSelectionType() || 'none';
  } catch {
    return { kind: 'none', letters: 0 };
  }
  // The editor separates paragraphs with CRLF; hand over plain line ends.
  text = text.replace(/\r\n?/g, '\n');
  if (text.trim()) {
    const letters = text.length;
    const paragraphs = text.replace(/\n$/, '').split('\n').length;
    if (letters > SELECTION_LETTER_LIMIT) return { kind: 'text', letters, tooLong: true, paragraphs };
    return { kind: 'text', letters, text, paragraphs };
  }
  // Slides report "slide" and Write reports "none" when nothing is picked;
  // a picked picture, shape or table cell without words is something else.
  return { kind: type === 'none' || type === 'slide' || type === 'text' ? 'none' : 'other', letters: 0 };
}

/** What is picked now. Never throws: with no editor, nothing is picked. */
export function readSelection(): PickedSelection {
  const api = getEditorApi();
  if (!api) return { kind: 'none', letters: 0 };
  return isSpreadsheet(api) ? readCells(api) : readWords(api);
}

function asLines(value: unknown): string[] {
  if (!Array.isArray(value) || !value.length || value.length > 200) {
    throw new SelectionError('bad-request', 'document:replace-selection needs "lines", a list of words');
  }
  const lines = value.map((line) => {
    if (typeof line !== 'string') throw new SelectionError('bad-request', 'Every line must be a string');
    return line.replace(/[\r\n]+/g, ' ');
  });
  const total = lines.reduce((sum, line) => sum + line.length, 0);
  if (!lines.join('').trim() || total > SELECTION_LETTER_LIMIT * 3) {
    throw new SelectionError('bad-request', 'The new words are empty or too long');
  }
  return lines;
}

/**
 * Write `lines` over the picked words, if they are still `expect`.
 * Write and Slides only. One paragraph at a time: a pick that reaches into
 * several paragraphs is refused, so no bullet or paragraph is ever merged away.
 */
export function replaceSelection(payload: Record<string, unknown>): { letters: number; text: string } {
  const lines = asLines(payload.lines);
  if (typeof payload.expect !== 'string') {
    throw new SelectionError('bad-request', 'document:replace-selection needs "expect", the picked words as read');
  }
  const expect = payload.expect.replace(/\r\n?/g, '\n');
  const api = getEditorApi();
  if (!api) throw new SelectionError('no-selection', 'No document is open');
  if (isSpreadsheet(api)) throw new SelectionError('not-supported', 'Cells are not changed this way');
  if (getReadonlyMode()) throw new SelectionError('read-only', 'The document is read-only');
  const logic = logicOf(api);
  const context = getEditorContext();
  const Settings = (context?.AscCommon as { CAddTextSettings?: new () => { SetTextPr(pr: TextPr): void } } | undefined)
    ?.CAddTextSettings;
  if (!logic || typeof logic.AddTextWithPr !== 'function' || typeof logic.GetDirectTextPr !== 'function' || !Settings) {
    throw new SelectionError('not-supported', 'This editor cannot change picked words');
  }

  const now = readWords(api);
  if (now.kind !== 'text' || now.tooLong || now.text !== expect) {
    throw new SelectionError('selection-changed', 'The picked words are not the ones that were read');
  }
  let picked = expect;
  // A whole paragraph picked by a triple click carries its paragraph end.
  // Leave that end out, or the next paragraph would be pulled into this one.
  if (picked.endsWith('\n')) {
    logic.MoveCursorLeft?.(true, false);
    logic.Document_UpdateSelectionState?.();
    const shrunk = readWords(api);
    if (shrunk.kind !== 'text' || shrunk.text !== picked.slice(0, -1)) {
      throw new SelectionError('not-verified', 'The picked words could not be narrowed to one paragraph');
    }
    picked = shrunk.text;
  }
  if (picked.includes('\n')) {
    throw new SelectionError('not-supported', 'Pick words inside one paragraph');
  }

  const text = lines.join(' ');
  const look = logic.GetDirectTextPr();
  const settings = new Settings();
  if (look) settings.SetTextPr(look.Copy());
  const view = keepView(api);
  logic.AddTextWithPr(text, settings);
  view();
  return { letters: text.length, text };
}

type Push = (payload: { kind: SelectionKind; letters: number }) => void;

interface ScrollBar {
  getCurScrolledX?(): number;
  getCurScrolledY?(): number;
  scrollToX?(x: number): void;
  scrollToY?(y: number): void;
}

/**
 * Notes where the page is scrolled, and gives back a function that puts it
 * there again. The editor scrolls to keep the cursor in sight after new words
 * go in; for words changed in place the page should not move, or it looks as
 * if it jumped (a narrow window scrolled sideways to the end of the new words).
 */
function keepView(api: EditorApi): () => void {
  const control = api.WordControl as { m_oScrollHorApi?: ScrollBar; m_oScrollVerApi?: ScrollBar } | undefined;
  const across = control?.m_oScrollHorApi;
  const down = control?.m_oScrollVerApi;
  const x = across?.getCurScrolledX?.();
  const y = down?.getCurScrolledY?.();
  const restore = () => {
    try {
      if (typeof x === 'number' && across?.getCurScrolledX?.() !== x) across?.scrollToX?.(x);
      if (typeof y === 'number' && down?.getCurScrolledY?.() !== y) down?.scrollToY?.(y);
    } catch {
      // A view that cannot be put back stays where the editor put it.
    }
  };
  return () => {
    restore();
    // The editor may scroll once more as it draws the new words.
    setTimeout(restore, 0);
    setTimeout(restore, 120);
  };
}

/**
 * Push `{ kind, letters }` whenever what is picked changes, at most four times a
 * second. Listens to the editor's own selection events and to the pointer and
 * keys inside its frame, with a slow check as a safety net. Returns a stop function.
 */
export function watchSelection(push: Push, options: { slowCheckMs?: number } = {}): () => void {
  let last = '';
  let lastPushAt = 0;
  let timer: ReturnType<typeof setTimeout> | undefined;
  const hookedApis = new WeakSet<object>();
  const hookedWindows = new WeakSet<object>();

  const check = () => {
    timer = undefined;
    lastPushAt = Date.now();
    const { kind, letters } = readSelection();
    const key = `${kind}:${letters}`;
    if (key === last) return;
    last = key;
    push({ kind, letters });
  };
  const schedule = () => {
    if (timer !== undefined) return;
    const wait = Math.max(0, SELECTION_PUSH_GAP_MS - (Date.now() - lastPushAt));
    timer = setTimeout(check, wait);
  };
  const hook = () => {
    const api = getEditorApi();
    if (api && !hookedApis.has(api) && typeof api.asc_registerCallback === 'function') {
      hookedApis.add(api);
      for (const event of ['asc_onSelectionChanged', 'asc_onFocusObject', 'asc_onSelectionEnd']) {
        try {
          api.asc_registerCallback(event, schedule);
        } catch {
          // An editor without this event has the others and the frame listeners.
        }
      }
    }
    const frame = document.querySelector('iframe[name="frameEditor"]') as HTMLIFrameElement | null;
    let win: Window | null = null;
    try {
      win = frame?.contentWindow ?? null;
    } catch {
      win = null;
    }
    if (win && !hookedWindows.has(win)) {
      hookedWindows.add(win);
      for (const event of ['mouseup', 'keyup', 'touchend', 'pointerup']) {
        win.addEventListener(event, schedule, true);
      }
    }
  };
  const slow = setInterval(() => {
    hook();
    schedule();
  }, options.slowCheckMs ?? 1000);
  hook();
  return () => {
    clearInterval(slow);
    if (timer !== undefined) clearTimeout(timer);
  };
}
