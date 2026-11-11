# Live collaboration and share links

ClerdoGrid has no server. People edit one workbook together by linking their devices directly, and share copies through links that carry the workbook inside them.

Code: `src/features/collab/`. Tests: `tests/collab*.test.ts` (unit) and `tests/collab.e2e.cjs` (two browser contexts over real WebRTC).

## Starting and joining

1. The host chooses **File → Share → Start live session**, enters a name, and gets an **invite code** (text starting `CGO1.`) and a QR code.
2. The person joining pastes the code under **File → Share → Join live session**, or scans the QR code: it is a link to ClerdoGrid with `#join=<code>`, which opens the Join dialog with the code filled in. They get a **reply code** (`CGA1.`) and send it back.
3. The host pastes the reply code and selects **Connect**. The joining device receives the host's workbook and the session is live.

Each invite code is for one person. For more people the host makes another invite (people popover → **Invite someone**). An invite can be **view only**: that person sees everything live but can't change anything (their device blocks edits and the host ignores any that arrive).

The host must keep their window open. When the host ends the session (or closes the window), everyone keeps their own copy of the workbook and can save it.

## How the codes work

A code is a WebRTC session description (SDP) for one data channel, with lines a data channel doesn't need removed, deflated and written in base64url. It holds the network candidates (addresses and ports) the device can be reached at, the ICE username fragment and password for this one connection attempt, and the DTLS certificate fingerprint. It holds no workbook data, no account and nothing that stays valid after the connection is made. Codes are not trickled: each side waits up to 5 seconds for its candidates before showing its code.

If the host chooses to share their TURN server (see below), the invite code starts with `CGO2.` and also carries that server's address, user name and password, so the guest can use the same relay. Reply codes never carry it.

## Network needs

- **Same network** (office, home Wi-Fi): works with no internet access at all. Set the STUN server to empty under **Network** in the Start or Join dialog.
- **Different networks**: each device asks a STUN server for its public address. The default is `stun:stun.l.google.com:19302`; it can be changed (or several given) in the **Network** field, stored per device. STUN only reveals a device's public IP address and port to that server; no workbook data goes through it.
- **Strict NATs and firewalls**: two devices that are both behind strict (symmetric) NATs, or on networks that block UDP, can't reach each other directly. Add your own **TURN server** (a relay, for example coturn on a small server) under **Network** in the Start or Join dialog: its address (`turn:host:3478`, or `turns:host:5349` for TLS, optionally with `?transport=tcp`), user name and password. These are kept only on this device (`localStorage` key `clerdogrid.turn`). The host can tick **Put this TURN server and its login in my invite codes** so guests use it without setting anything; only do that with a login you are happy to give to the people you invite (a time-limited TURN login is best). The relay forwards the encrypted data channel; it can't read the workbook. ClerdoGrid doesn't provide a TURN server. Without one, use a share link or the file.
- The content security policy of the built app (`connect-src 'self'`) doesn't restrict WebRTC, so the data channel and STUN work in the browser, the home-screen app and the desktop app.

## Security

- **Encryption**: the data channel runs over DTLS (WebRTC requires it); the certificate fingerprints in the codes make sure each side talks to the device that made the other code. Anyone who intercepts a code before it is used could try to connect in its place; send codes over a channel you trust, and the host sees each person by name as they join.
- **Every message is checked** (`validate.ts`) before use: message and frame sizes (single frames up to 256 KB, whole messages up to 160 MB, at most 4 partly received messages), types and bounds of every field (rows and columns inside the sheet, sheet names, text and formula lengths, style and JSON nesting depth, peer ids and colors), links limited to `https:`, `mailto:` and in-workbook `#` targets, pictures limited to `image/*` types and 40 MB. Unknown operation types are ignored; anything malformed drops the whole message. `__proto__`, `constructor` and `prototype` keys are refused. Nothing is evaluated: messages are parsed with `JSON.parse`, formulas go through the formula parser.
- Peers can't add unknown file parts (`extraParts`) to someone else's workbook; only the host's full workbook (sent on join) carries them.
- **Share links** are read like a file from disk: the `#wb=` data is inflated with a 64 MB cap and opened with the normal `.xlsx` reader, which treats it as untrusted input. The hash part of a link is never sent to a web server.

## Sync model

The host is the authority. Every change made on any device goes through `GridStore.change()`; an observer (`store.observe`) turns the finished change into **operations** (`ops.ts`):

| Operation | What it carries |
|---|---|
| `cells` | cell values, formulas, styles (as style objects, since style ids differ between copies), rich text, array ranges |
| `struct` | a row or column insert/delete (from `store.markOp` in `changeStructure`) |
| `dim` | row heights, column widths, hidden and outline state, row/column styles |
| `keyed` | notes, links and comment threads per cell |
| `item` | one chart, picture, shape or ink stroke by id |
| `prop` | any other sheet part as a whole value: merges, conditional formats, validations, tables, filters, frozen panes, page setup, protection, sparklines, pivots, and parts added later |
| `sheetAdd`, `sheetDel`, `sheetRen`, `sheetPos` | the sheet list |
| `book` | workbook parts: defined names, calculation settings, structure protection, scripts and parts added later |

Changes are found by comparing each part with the snapshot `change()` took first, so every operation that uses `change()` syncs without knowing about collaboration.

**Sheet identity** (`sheets.ts`). In a session each sheet has a stable id and a sort key (kept on the sheet, never saved to files); operations name sheets by id, so edits follow a sheet through renames, moves and concurrent inserts. The sheet list is always in (key, id) order: moving or adding a sheet only sets its own key, so concurrent moves can't disagree about the order. The host gives ids to every sheet when the session starts (including the copies in its undo history). A deleted sheet is kept, out of sight, on every copy and still takes edits made to it at the same time; undoing the delete brings back that kept copy, so everyone gets the same sheet back. If people delete every sheet at once, the deleted sheet last in order is shown again, which every copy works out the same way.

**Ordering.** Guests apply their own changes at once and send them to the host with the last host sequence number they had seen. One batch is in flight per guest; later changes wait and go as one batch when it is acknowledged. The host transforms an incoming batch past every change it ordered since that sequence number (operational transformation, `ot.ts`), applies it, gives it the next sequence number and sends it to everyone (the sender takes it as an acknowledgement). A guest receiving someone else's change transforms it past its own unacknowledged changes and those past it, then applies it. So:

- Concurrent edits to different cells both land. Edits to the same cell: the one the host ordered later wins, on every device.
- Row and column inserts and deletes move other people's concurrent cell edits, formulas (references are adjusted as for a local insert), merges, rule and validation ranges, chart anchors and cursors. A delete that overlaps an insert keeps the inserted lines. An edit inside deleted rows is dropped.
- Sheet renames and deletes fix sheet names in concurrent formulas. Two sheets given the same name at once (by renames, new sheets or undo): the one the host ordered later gets the name plus part of its id, such as "Data (k3f9)". A delete and an undo of that delete at the same time: the later one wins.
- After applying others' changes the engine recalculates.

**Checksum.** While the session is idle the host sends a hash of the shared content (formula results, selections and zoom are left out). A guest whose copy differs asks for the host's workbook and reloads it. This is a safety net: in the random tests (1,500 runs of 60 concurrent edits by three people, mixing cell edits, row and column inserts and deletes, styles, clears, sheet inserts, renames, deletes, moves, undo and redo) every copy agrees without it. In longer, sheet-heavy runs (6,000 runs of 100 edits) about 1 run in 3,000 still needs it, after repeated undo and redo of sheet deletes. Large workbooks are hashed less often so the check stays cheap.

**Undo.** Others' changes never go on your undo stack. When one arrives, the snapshots in your undo and redo history are patched with it, so undoing your step restores what you had before without reverting their later edits. Positions move through the history: someone else's row or column insert or delete is applied to every stored snapshot (cells, whole sheets and whole books, with formulas adjusted, using the same code as a local insert), and their edits are moved back through your own earlier inserts and deletes (and forward through the ones on the redo stack) before they are patched in, so they land in the right place when you undo past those steps. Undoing or redoing a row or column insert or delete is sent as that insert or delete (plus any restored contents), so others' concurrent edits move around it. Only a step whose snapshot no longer matches any sheet is dropped, with the steps behind it.

**Joining and reconnecting.** On join the host sends the whole workbook as a structured snapshot (the same model, with a style table), checked like any message. If the connection drops, the guest keeps working; their unsent changes are kept. The host makes a new invite, the guest selects **Rejoin**, and the changes are sent with the new hello message, transformed on the host if the history still reaches back far enough (the host keeps the last 5,000 changes), and then the guest receives the current workbook. Pings every 4 seconds detect a silent connection after 75 seconds; normal drops are noticed by WebRTC within seconds, and a brief network change shows "Reconnecting…" while WebRTC recovers by itself.

Long messages are split into 16 KB frames with back-pressure on the data channel's buffer.

## Presence

Each person has a name (remembered per device, the same name comments use) and a color the host assigns. The top bar shows everyone's initials; the popover lists where each person is and what they're doing. Other people's selections and active cells are drawn over the grid with name flags; a dashed outline and a pencil mean they are typing in that cell. **Go to** jumps to a person's sheet and selection once; **Follow** keeps doing that until you move yourself.

## Share links

**File → Share → Copy link** (or **Copy read-only link**) writes the workbook as `.xlsx`, deflates it and puts it in the link: `#wb=<base64url>&ro=1&n=<name>`. Links up to 2 MB of encoded data are offered; above that, **Share file** uses the system share sheet where it can take files (iPad, phones) and otherwise downloads the file. Opening a link loads a copy; a read-only copy shows **Read-only** in the top bar with **Edit a copy** to unlock it. The link is removed from the address bar after opening, and an editable copy is kept as an unsaved draft so a reload doesn't lose edits.

## Testing

```bash
npx vitest run tests/collab            # transforms, op application, sessions over an in-memory network, random convergence, QR, links
npx vite build --outDir /tmp/cg && npx vite preview --outDir /tmp/cg --port 5471 &
node tests/collab.e2e.cjs http://127.0.0.1:5471/ /tmp   # two browser contexts over WebRTC; writes screenshots
```

The QR encoder (`qr.ts`) follows ISO/IEC 18004: byte mode, error correction level M, versions 1 to 40, all eight masks with the standard penalty rules.
