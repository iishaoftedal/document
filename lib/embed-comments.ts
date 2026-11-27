/**
 * A comment from the page that embeds the editor. Added in this copy (iishaoftedal/document), 2026-11-26.
 *
 * `document:add-comment` `{ text, author, quote? }` answers `document:comment-added`. It goes through the same
 * path as the agent's `add_comment` tool (`asc_addComment` with the editor's own comment data), so the comment
 * is the editor's own: shown on the page, kept in the file, and read by Word, Excel and PowerPoint.
 *
 * - With no `quote`, the comment goes on what is picked (in a spreadsheet, the cell that is picked).
 * - With a `quote`, the editor finds those exact words (letter for letter, capitals included). If they are in
 *   the document once, the comment goes on them. If they are not there, or are there more than once, nothing is
 *   added and the answer is `document:error` (`not-found` or `not-once`). Write and Slides; a spreadsheet takes
 *   no quote (`not-supported`), only the picked cell.
 * - `author` is the name the comment is signed with. Left out, it is the person the document was opened as.
 * - Nothing is added to a document opened read-only (`read-only`).
 * - Words found for a quote are let go once the comment is on them, so nothing is left picked.
 */
import { getEditorContext, type EditorApi } from './agent-plugin/editor-bridge';
import { getEditorUser, USER_NAME_LIMIT } from './onlyoffice/editor-user';
import { getReadonlyMode } from './onlyoffice/readonly';

/** Longest comment taken, and longest quote looked for. */
export const COMMENT_LETTER_LIMIT = 4000;
export const QUOTE_LETTER_LIMIT = 1000;

export class CommentError extends Error {
  constructor(
    readonly code: 'not-found' | 'not-once' | 'not-supported' | 'read-only' | 'no-document' | 'bad-request',
    message: string,
  ) {
    super(message);
    this.name = 'CommentError';
  }
}

interface SearchSettings {
  put_Text(text: string): void;
  put_MatchCase(on: boolean): void;
  put_WholeWords(on: boolean): void;
}

interface FindingApi {
  asc_findText?(settings: SearchSettings, next: boolean): number;
  asc_endFindText?(): void;
  wb?: { getWorksheet?: unknown };
}

const isSpreadsheet = (api: EditorApi): boolean => {
  const wb = (api as unknown as FindingApi).wb;
  return !!wb && typeof wb === 'object' && typeof wb.getWorksheet === 'function';
};

const oneLine = (value: string, most: number) => value.replace(/\s+/g, ' ').trim().slice(0, most).trim();

export interface CommentAdded {
  /** The words the comment is on, when a quote was given. */
  quote?: string;
  author: string;
}

export function addComment(payload: Record<string, unknown>): CommentAdded {
  const text = typeof payload.text === 'string' ? payload.text.trim().slice(0, COMMENT_LETTER_LIMIT) : '';
  if (!text) throw new CommentError('bad-request', 'A comment needs words');
  const author =
    (typeof payload.author === 'string' ? oneLine(payload.author, USER_NAME_LIMIT) : '') || getEditorUser().name;
  const quote = typeof payload.quote === 'string' ? payload.quote : undefined;
  if (quote !== undefined && (!quote.trim() || quote.length > QUOTE_LETTER_LIMIT)) {
    throw new CommentError('bad-request', 'The quote is empty or too long');
  }

  const context = getEditorContext();
  if (!context || !window.editor) throw new CommentError('no-document', 'No document is open');
  if (getReadonlyMode()) throw new CommentError('read-only', 'The document is open read-only');
  const { api, Asc, AscCommon } = context;

  if (quote !== undefined) {
    if (isSpreadsheet(api)) throw new CommentError('not-supported', 'A spreadsheet comment goes on the picked cell');
    const finder = api as unknown as FindingApi;
    const Settings = (AscCommon as unknown as { CSearchSettings?: new () => SearchSettings } | undefined)
      ?.CSearchSettings;
    if (typeof finder.asc_findText !== 'function' || typeof Settings !== 'function') {
      throw new CommentError('not-supported', 'This editor cannot look for words');
    }
    const settings = new Settings();
    settings.put_Text(quote);
    settings.put_MatchCase(true);
    settings.put_WholeWords(false);
    // Picks the first one it finds, and says how many there are.
    const found = finder.asc_findText(settings, true);
    if (found !== 1) {
      finder.asc_endFindText?.();
      throw new CommentError(
        found ? 'not-once' : 'not-found',
        found ? `Those words are in the document ${found} times` : 'Those words are not in the document',
      );
    }
  }

  // The path the agent's add_comment tool takes (lib/agent-plugin/tools.ts).
  const CommentCtor = Asc.asc_CCommentDataWord ?? Asc.asc_CCommentData;
  if (typeof CommentCtor !== 'function') throw new CommentError('not-supported', 'Comments are not supported here');
  const data = new CommentCtor() as InstanceType<typeof CommentCtor> & {
    asc_putTime?(time: string): void;
    asc_putSolved?(solved: boolean): void;
  };
  data.asc_putText(text);
  data.asc_putUserName(author);
  const user = getEditorUser();
  data.asc_putUserId(author === user.name ? user.id : `author-${author.toLowerCase().replace(/\W+/g, '-')}`);
  data.asc_putTime?.(String(Date.now()));
  data.asc_putSolved?.(false);
  api.asc_addComment(data);
  (api as unknown as FindingApi).asc_endFindText?.();
  // Words the editor found for a quote are let go again: the person picked nothing, so nothing stays picked.
  if (quote !== undefined) api.asc_RemoveSelection?.();
  return quote !== undefined ? { quote, author } : { author };
}
