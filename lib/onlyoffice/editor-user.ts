/**
 * Who the editor opens a document as. Added in this copy (iishaoftedal/document), 2026-11-26.
 *
 * The editor mounts with an `editorConfig.user`, and every comment, reply and tracked change made in the
 * document is signed with that person's name, the way Word signs them. Until now that was always "Guest".
 * An embedding page may now name the person on its open message (`user: { id, name }`, see embed-api.ts);
 * with no `user`, it is "Guest" as before.
 *
 * The name goes into the document and nowhere else: this editor has no server, and the name is not stored,
 * logged or sent. The next open without a `user` goes back to "Guest".
 */

export interface EditorUser {
  id: string;
  name: string;
}

export const GUEST: EditorUser = Object.freeze({ id: 'local-user', name: 'Guest' });

/** The longest name kept, so a name cannot crowd the comment cards. */
export const USER_NAME_LIMIT = 80;

let current: EditorUser = GUEST;

/**
 * The person an open message names, or null when it names nobody (or names them in a shape this does not
 * take). A name is one line of plain words; an id is kept short and plain, and made up when missing.
 */
export function editorUserFrom(value: unknown): EditorUser | null {
  if (!value || typeof value !== 'object') return null;
  const raw = value as { id?: unknown; name?: unknown };
  const name =
    typeof raw.name === 'string'
      ? Array.from(raw.name, (ch) => (ch.charCodeAt(0) < 32 || ch.charCodeAt(0) === 127 ? ' ' : ch))
          .join('')
          .replace(/\s+/g, ' ')
          .trim()
          .slice(0, USER_NAME_LIMIT)
          .trim()
      : '';
  if (!name) return null;
  const id =
    typeof raw.id === 'string' && /^[\w.@:-]{1,64}$/.test(raw.id)
      ? raw.id
      : `user-${name.toLowerCase().replace(/\W+/g, '-')}`;
  return { id, name };
}

/** Sets who the next document opens as. `null` is "Guest". */
export function setEditorUser(user: EditorUser | null): void {
  current = user ?? GUEST;
}

/** Who the editor opens documents as now. */
export function getEditorUser(): EditorUser {
  return current;
}
