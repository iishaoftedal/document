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

/**
 * Pictures for comment authors the host names (added in this copy, 2026-11-26). The editor draws a comment's
 * author as a round mark with their initials; when it has a picture for that author it draws the picture instead.
 * An open message may name authors with a picture (`people: [{ name, image }]`), so an author that is not a person
 * (Wally's own comments, signed "Wally (AI)") is not drawn like one. A picture is a small image held in the message
 * itself (a data: address, PNG, JPEG or SVG, at most 64 KB): nothing is fetched from anywhere. The next open
 * replaces the list.
 */
export interface EditorPerson {
  id: string;
  name: string;
  image: string;
}

export const PICTURE_LIMIT = 64 * 1024;

/** The id a comment carries for an author who is not the person the document opened as (embed-comments.ts). */
export const authorIdFor = (name: string): string => `author-${name.toLowerCase().replace(/\W+/g, '-')}`;

let people: EditorPerson[] = [];

/** The authors an open message names with a picture; anything else in the list is left out. */
export function editorPeopleFrom(value: unknown): EditorPerson[] {
  if (!Array.isArray(value)) return [];
  return value
    .slice(0, 20)
    .map((raw) => {
      const one = (raw ?? {}) as { name?: unknown; image?: unknown };
      const user = editorUserFrom({ name: one.name });
      const image = typeof one.image === 'string' ? one.image : '';
      const fine =
        /^data:image\/(png|jpeg|svg\+xml);base64,[A-Za-z0-9+/=]+$/.test(image) && image.length <= PICTURE_LIMIT;
      return user && fine ? { id: authorIdFor(user.name), name: user.name, image } : null;
    })
    .filter((one): one is EditorPerson => one !== null);
}

export function setEditorPeople(list: EditorPerson[]): void {
  people = list;
}

/** The authors with a picture among the ids the editor asks about. */
export function editorPeopleFor(ids: unknown): EditorPerson[] {
  const asked = Array.isArray(ids) ? ids : [ids];
  return people.filter((one) => asked.includes(one.id));
}
