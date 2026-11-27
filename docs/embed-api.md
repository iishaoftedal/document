# iframe Embed API

This project supports embedding into any web application via iframe. The recommended pattern is: **the parent system handles auth, file fetching, and upload; the iframe handles editing only.** Tokens, cookies, and business APIs stay in the parent — the editor never sees them.

A working demo is available at `/embed-demo.html` (includes sha256 logging for debugging).

---

## Embedding the editor

```html
<iframe
  id="documentEditor"
  src="https://your-deployment/editor?embed=1"
  style="width: 100%; height: 720px; border: 0"
></iframe>
```

> The editor lives at `/editor`; the homepage `/` is a static landing page. Older links to `/?embed=1`, `/?src=`, `/?file=` and `/?new=` still work -- `/` redirects them to `/editor` with the same query.

Commands are accepted only from the immediate parent window. Without
`embedOrigin`, the first accepted parent command fixes the origin for the lifetime
of the editor page; later commands from a different origin are ignored.
To restrict messages to a specific origin from startup, add `embedOrigin`:

```html
<iframe
  id="documentEditor"
  src="https://your-deployment/editor?embed=1&embedOrigin=https://your-system.example.com"
></iframe>
```

---

## Sending commands

Include an `id` on each command to match it to the response:

```js
const iframe = document.getElementById('documentEditor');
const editorOrigin = 'https://your-deployment';

function sendEditorCommand(type, payload = {}) {
  const id = `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  iframe.contentWindow.postMessage({ id, type, payload }, editorOrigin);
  return id;
}

window.addEventListener('message', (event) => {
  if (event.origin !== editorOrigin) return;
  const { id, type, payload } = event.data || {};
  if (!type?.startsWith('document:')) return;

  switch (type) {
    case 'document:ready':
      console.log('Editor ready');
      break;
    case 'document:opened':
      console.log('Opened', id, payload);
      break;
    case 'document:saved':
      console.log('Saved', payload.fileName, payload.file);
      break;
    case 'document:error':
      console.error('Error', payload.message);
      break;
  }
});
```

---

## Opening a document

### From URL

> **The URL must allow CORS.** The browser -- not a server -- fetches the file,
> so the response needs an `Access-Control-Allow-Origin` header that covers the
> page the editor runs on. Without it the request is blocked before the editor
> sees a single byte, and the same is true of `?src=` and `?file=` on the URL.
> This is the single most common thing to get wrong when integrating: if a file
> opens when you download it by hand but not through the editor, check the
> response headers first. Cross-origin redirects have to keep the header too.
> When you cannot add it (a third-party host, a signed URL, anything behind
> auth), fetch the file in the parent page and pass the bytes with
> `document:open-buffer` instead.

Remote URL downloads are limited to **100 MiB of decoded response data** and
**60 seconds**, including the response body. Download larger files separately
and open them with `document:open-file` or `document:open-buffer`. These limits
apply to URL downloads, not local file opens.

A failed, timed-out or superseded URL open returns `document:error` for its
request ID. Opening another document cancels the older download; late responses
and pending conversions cannot replace the newer editor. Encoded `filename*`
response headers take precedence over `filename`.

```js
sendEditorCommand('document:open-url', {
  url: 'https://example.com/files/demo.xlsx',
  fileName: 'demo.xlsx',
  readonly: false,
});
```

If the URL requires auth headers, pass `fetchOptions`. For protected files it is preferable to fetch in the parent system and pass the binary:

```js
sendEditorCommand('document:open-url', {
  url: 'https://example.com/api/files/1',
  fileName: 'demo.xlsx',
  fetchOptions: { headers: { Authorization: `Bearer ${token}` } },
});
```

### From a file picker

```js
const input = document.createElement('input');
input.type = 'file';
input.accept = '.xlsx,.xls,.csv,.docx,.doc,.pptx,.ppt';
input.onchange = () => {
  sendEditorCommand('document:open-file', { file: input.files[0], readonly: false });
};
input.click();
```

### From an authenticated fetch (recommended for protected files)

```js
const response = await fetch('/api/files/1', {
  headers: { Authorization: `Bearer ${token}` },
});
const buffer = await response.arrayBuffer();
sendEditorCommand('document:open-buffer', { fileName: 'demo.xlsx', buffer, readonly: false });
```

---

## Read-only mode

Set at open time via the `readonly` field, or toggle at any time:

```js
sendEditorCommand('document:set-readonly', { readonly: true });
```

In read-only mode, editing is disabled. Export still works: `document:save`
returns the current bytes via `document:saved` (read-only blocks edits, not
downloads). The File tab's download control inside an embed iframe also
triggers a local download when the host has not armed a `document:save`.

---

## Saving and uploading

The save command exports the current document and returns a `File` via `document:saved`. Default format is the open document's own format (a `.docx` saves as DOCX, a `.csv` as CSV, ...); pass `targetExt` to change it.

```js
sendEditorCommand('document:save', { targetExt: 'XLSX' }); // XLSX, DOCX, PPTX, CSV
```

By default the command waits for the editor to return the **edited** file. If it times out, `document:error` is returned — this prevents accidentally uploading the original unchanged file. To opt in to returning the original on timeout:

```js
sendEditorCommand('document:save', { targetExt: 'XLSX', returnOriginalOnTimeout: true });
```

Upload from the parent:

```js
window.addEventListener('message', async (event) => {
  if (event.origin !== editorOrigin) return;
  const { type, payload } = event.data || {};
  if (type !== 'document:saved') return;

  await fetch('/api/files/1', {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}` },
    body: payload.file,
  });
});
```

> **Note:** Do not rely on `file.size` alone to detect changes. `.xlsx` is a zip archive — a minor edit can produce the exact same byte count. The built-in `/embed-demo.html` logs a `sha256` hash on every save for easier debugging.

---

## Light or dark

Added in this copy (iishaoftedal/document). The host can make the editor match its own look, at load with a
`theme` parameter and at any time with a message. Only `light` and `dark` are accepted.

```html
<iframe src="https://your-deployment/editor?embed=1&theme=dark"></iframe>
```

```js
sendEditorCommand('document:set-theme', { theme: 'light' });
```

---

## Picked words

Added in this copy (iishaoftedal/document), 2026-11-26. The host can learn what the user has picked, read the
picked words when it needs them, and write new words over them. The words cross only when the host asks.

```js
// Pushed by the editor whenever what is picked changes, at most four times a second.
// Only the kind and a count: never the words themselves.
// { type: 'document:selection-changed', payload: { kind: 'text', letters: 42 } }

sendEditorCommand('document:get-selection');
// Response: { type: 'document:selection', payload: { kind: 'text', letters: 42, text: '...', paragraphs: 1 } }
// In a spreadsheet: { kind: 'cells', letters: 0, cells: 13, range: 'B2:B14' }

sendEditorCommand('document:replace-selection', { lines: ['New words.'], expect: '...the words as read...' });
// Response: { type: 'document:selection-replaced', payload: { letters: 10, text: 'New words.' } }
```

- `kind` is `none`, `text`, `cells` or `other` (a picture, a shape, separate blocks of cells, or typing inside a
  cell). In a spreadsheet the single cell that is always selected counts as `none`; two or more cells are `cells`,
  with the range.
- At most 6,000 letters are handed over. Past that `document:selection` has `tooLong: true` and no `text`.
- `document:replace-selection` changes nothing, and answers `document:error` with `code: 'selection-changed'`,
  when what is picked now is not `expect` (the user typed or picked something else meanwhile). It writes through
  the editor's own add-text action with the look of the first picked letter (bold stays bold), as one step that the
  editor's own Undo takes back. Write and Slides only; cells are refused (`code: 'not-supported'`).
- One paragraph at a time: a whole paragraph picked with its paragraph end is narrowed to its words, and words that
  reach into a second paragraph are refused (`code: 'not-supported'`), so no paragraph or bullet is merged away.
  Several `lines` are joined with spaces.
- Other codes on `document:error`: `read-only`, `no-selection` (no document), `bad-request`, `not-verified`.

---

## Comments and who is writing

Added in this copy (iishaoftedal/document), 2026-11-26. Comments are the editor's own: kept inside the file the way
Word, Excel and PowerPoint keep them, so they show there too. Two things are new for the host.

**Who the document opens as.** Every open message (`document:open-file` and the others) takes a `user`. The editor
opens as that person, and the comments, replies and tracked changes made there are signed with their name. With no
`user`, it opens as "Guest", as before. The name is one line of up to 80 letters. It goes into the document and
nowhere else: the editor has no server and does not store, log or send it.

```js
sendEditorCommand('document:open-file', { file, user: { id: 'owner-42', name: 'Sam Rivera' } });
```

**Pictures for authors who are not people.** An open message may also name comment authors with a picture
(`people: [{ name, image }]`, up to 20). The editor draws that author's comments with the picture in place of a round
mark with their initials, so an assistant that leaves comments is not drawn like a person. A picture is a small image
held in the message itself (a `data:` address, PNG, JPEG or SVG, at most 64 KB); nothing is fetched. The next open
replaces the list. When the editor asks for an author it has no picture for (or for the list of people to mention
with @), it is answered with nobody.

```js
sendEditorCommand('document:open-file', {
  file,
  user,
  people: [{ name: 'Wally (AI)', image: 'data:image/svg+xml;base64,…' }],
});
```

**A comment from the host.**

```js
sendEditorCommand('document:add-comment', {
  text: 'Check this matches the books.',
  author: 'Wally',
  quote: 'The rent is $2,400 a month',
});
// Response: { type: 'document:comment-added', payload: { quote: 'The rent is $2,400 a month', author: 'Wally' } }
```

- With no `quote`, the comment goes on what is picked (in a spreadsheet, the picked cell).
- With a `quote`, the editor finds those exact words, capitals included. If they are in the document once, the
  comment goes on them. If not, nothing is added and the answer is `document:error` with `code: 'not-found'`, or
  `'not-once'` when they are there more than once. Write and Slides only: a spreadsheet takes no quote
  (`'not-supported'`).
- `author` signs the comment. Left out, it is the person the document was opened as.
- It goes through the same path as the agent's `add_comment` tool, as one step the editor's own Undo takes back,
  and it marks the document as changed (`document:dirty-changed`), so the host saves it as it saves any edit.
- Other codes on `document:error`: `read-only`, `no-document`, `bad-request` (no words, or an empty or over-long
  quote).

---

## Query current state

```js
sendEditorCommand('document:get-state');
// Response: { type: 'document:state', payload: { readonly: false, hasDocument: true, dirty: false } }
```

---

## Tracking unsaved changes

The editor tells the parent whether the open document has edits the parent does not hold yet, so the parent can save when it matters (switching documents, an idle timer, closing its own page) without exporting documents nobody touched.

```js
window.addEventListener('message', (event) => {
  if (event.origin !== editorOrigin) return;
  const { type, payload } = event.data || {};
  if (type === 'document:dirty-changed') {
    console.log(payload.dirty ? 'Unsaved changes' : 'Everything saved');
  }
});
```

- `document:dirty-changed` is pushed when the flag flips, not on every keystroke: `{ dirty: true }` after the first edit, `{ dirty: false }` once a `document:save` covered every edit (or undo went back to the saved state).
- A successful `document:save` clears the flag: the parent now holds the bytes. Edits made while the export was running are not in those bytes, so they keep the document dirty -- check `payload.dirty` on `document:saved` and save again when it is `true`.
- If the upload on the parent side fails, the editor cannot know: keep your own "needs saving" state until the upload succeeds.
- `document:save` with `returnOriginalOnTimeout: true` never clears the flag, since the file returned on timeout may be the original.
- Opening another document resets the flag to `false`.

A save-on-idle loop for the parent could look like this:

```js
let idleTimer;
window.addEventListener('message', (event) => {
  if (event.origin !== editorOrigin) return;
  const { type, payload } = event.data || {};
  if (type === 'document:dirty-changed' && payload.dirty) {
    clearTimeout(idleTimer);
    idleTimer = setTimeout(() => sendEditorCommand('document:save'), 30_000);
  }
  if (type === 'document:saved' && payload.dirty) {
    // Edits landed during the export: they still need a save.
    clearTimeout(idleTimer);
    idleTimer = setTimeout(() => sendEditorCommand('document:save'), 30_000);
  }
});
```

---

## Message reference

| Direction       | Type                          | Description                                        |
| --------------- | ----------------------------- | -------------------------------------------------- |
| parent → iframe | `document:open-url`           | Open document from URL                             |
| parent → iframe | `document:open-file`          | Open document from `File` / `Blob` (and `user`)    |
| parent → iframe | `document:open-buffer`        | Open document from `ArrayBuffer` / `Uint8Array`    |
| parent → iframe | `document:set-readonly`       | Set read-only or editable                          |
| parent → iframe | `document:save`               | Save and return `File`                             |
| parent → iframe | `document:set-theme`          | Switch the editor to `light` or `dark`             |
| parent → iframe | `document:get-state`          | Query current state                                |
| parent → iframe | `document:get-selection`      | Read what is picked (added in this copy)           |
| parent → iframe | `document:replace-selection`  | Write new words over the picked ones (this copy)   |
| parent → iframe | `document:add-comment`        | Add a comment, on exact words or what is picked    |
| iframe → parent | `document:ready`              | Editor initialised                                 |
| iframe → parent | `document:opened`             | Document opened                                    |
| iframe → parent | `document:readonly-changed`   | Read-only state changed                            |
| iframe → parent | `document:saved`              | Save complete, file returned (with `dirty`)        |
| iframe → parent | `document:dirty-changed`      | Unsaved-changes flag flipped (`{ dirty }`)         |
| iframe → parent | `document:state`              | Current state response                             |
| iframe → parent | `document:selection`          | What is picked (`kind`, `letters`, `text`/`range`) |
| iframe → parent | `document:selection-changed`  | What is picked changed (`{ kind, letters }`)       |
| iframe → parent | `document:selection-replaced` | New words written over the picked ones             |
| iframe → parent | `document:comment-added`      | The comment is in (`{ author, quote? }`)           |
| iframe → parent | `document:error`              | Operation failed                                   |
