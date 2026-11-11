// Protect sheet, protect workbook structure, and password-encrypted files.
import { useState } from 'react';
import { Lock, LockOpen, ShieldCheck } from 'lucide-react';
import { command, registerCommand } from '../../app/commands';
import type { GridStore } from '../../app/store';
import { checkPassword, hasPassword, makePasswordHash, protectionHash } from '../../app/protection';
import { readXlsx } from '../../formats/xlsx';
import { decryptWorkbook, EncryptionError } from '../../formats/encryption';
import type { Sheet, Workbook } from '../../model/workbook';
import { Dialog, Field, type DialogProps } from '../../ui/dialogs/Dialog';
import { Check } from '../../ui/dialogs/shared';
import { N, T, Tr } from '../../ui/i18n';

// ---------- Protect sheet ----------

/**
 * sheetProtection flags (ECMA-376 Part 1, 18.3.1.85). A flag that is true means the action is blocked;
 * the third value is the file format's default.
 */
export const SHEET_FLAGS: [string, string, boolean][] = [
  ['selectLockedCells', N('Select locked cells'), false], ['selectUnlockedCells', N('Select unlocked cells'), false],
  ['formatCells', N('Format cells'), true], ['formatColumns', N('Format columns'), true], ['formatRows', N('Format rows'), true],
  ['insertColumns', N('Insert columns'), true], ['insertRows', N('Insert rows'), true], ['insertHyperlinks', N('Insert hyperlinks'), true],
  ['deleteColumns', N('Delete columns'), true], ['deleteRows', N('Delete rows'), true],
  ['sort', N('Sort'), true], ['autoFilter', N('Use filters'), true], ['pivotTables', N('Use PivotTables and PivotCharts'), true],
  ['objects', N('Edit objects'), false], ['scenarios', N('Edit scenarios'), false],
];

/** True when the sheet's protection blocks this action. */
export const blocked = (s: Sheet, flag: string): boolean => {
  const p = s.protection;
  if (!p?.sheet) return false;
  const def = SHEET_FLAGS.find(([k]) => k === flag)?.[2] ?? true;
  return p.options?.[flag] ?? def;
};

const password = (pw: string) => (pw ? { password: protectionHash(pw), hash: makePasswordHash(pw) } : {});

export const ProtectSheetDialog = ({ store, close }: DialogProps) => {
  const p = store.sheet.protection;
  const on = !!p?.sheet;
  const [pw, setPw] = useState('');
  const [pw2, setPw2] = useState('');
  const [allowed, setAllowed] = useState<Record<string, boolean>>(() => Object.fromEntries(SHEET_FLAGS.map(([k, , d]) => [k, !(p?.options?.[k] ?? d)])));
  const sheetIdx = store.sheetIndex;
  const submit = (): boolean => {
    if (on) {
      if (p && hasPassword(p) && !checkPassword(pw, p)) { store.notify(T('That password isn\'t right.')); return false; }
      store.change(N('Unprotect sheet'), 'sheet', () => { delete store.wb.sheets[sheetIdx]!.protection; });
      store.notify(T('Sheet unprotected.'));
      return true;
    }
    if (pw !== pw2) { store.notify(T('The passwords don\'t match.')); return false; }
    const options: Record<string, boolean> = Object.fromEntries(SHEET_FLAGS.map(([k]) => [k, !allowed[k]]));
    store.change(N('Protect sheet'), 'sheet', () => { store.wb.sheets[sheetIdx]!.protection = { sheet: true, ...password(pw), options }; });
    store.notify(pw ? T('Sheet protected with a password.') : T('Sheet protected.'));
    return true;
  };
  return (
    <Dialog title={on ? T('Unprotect sheet') : T('Protect sheet')} close={close} width={460} okLabel={on ? T('Unprotect') : T('Protect')} onOk={submit}>
      <div className="b-inline muted small">{on ? <LockOpen size={15} strokeWidth={1.9} aria-hidden="true" /> : <Lock size={15} strokeWidth={1.9} aria-hidden="true" />}{store.sheet.name}</div>
      {on ? <>
        <p>{p && hasPassword(p) ? T('Enter the password to unprotect this sheet.') : T('This sheet is protected without a password.')}</p>
        {p && hasPassword(p) && <Field label={T('Password')}><input type="password" autoFocus value={pw} onChange={(e) => setPw(e.target.value)} /></Field>}
      </> : <>
        <p className="muted small">{T('Locked cells can\'t be changed while the sheet is protected. Unlock cells first in Format cells, Protection.')}</p>
        <div className="grid2">
          <Field label={T('Password (optional)')}><input type="password" autoFocus value={pw} onChange={(e) => setPw(e.target.value)} autoComplete="new-password" /></Field>
          <Field label={T('Repeat password')}><input type="password" value={pw2} disabled={!pw} onChange={(e) => setPw2(e.target.value)} autoComplete="new-password" /></Field>
        </div>
        <Field label={T('Everyone can still')}>
          <div className="b-list b-checklist ws-flags">{SHEET_FLAGS.map(([k, l]) => <Check key={k} label={Tr(l)} checked={!!allowed[k]} onChange={(v) => setAllowed({ ...allowed, [k]: v })} />)}</div>
        </Field>
        <p className="muted small">{T('The password is stored as a salted SHA-512 hash. If you forget it, it can\'t be recovered.')}</p>
      </>}
    </Dialog>
  );
};

/** Makes the per-action flags of a protected sheet count: formatting, inserting, deleting, sorting, filtering, links and objects. */
export const guardCommands = (): void => {
  const rules: [RegExp, string, (s: GridStore) => boolean][] = [
    [/^(format\.|border\.|number\.|cells\.styles$|insert\.conditional$|insert\.clear(All)?Rules$|edit\.clearFormats$|edit\.formatPainter$)/, 'formatCells', (s) => anyLocked(s)],
    [/^cells\.(colWidth|autofitCols|hideCols|unhideCols)$/, 'formatColumns', () => true],
    [/^cells\.(rowHeight|autofitRows|hideRows|unhideRows)$/, 'formatRows', () => true],
    [/^cells\.insertCols$/, 'insertColumns', () => true],
    [/^cells\.insertRows$/, 'insertRows', () => true],
    [/^cells\.deleteCols$/, 'deleteColumns', () => true],
    [/^cells\.deleteRows$/, 'deleteRows', () => true],
    [/^data\.(sortAsc|sortDesc|sort)$/, 'sort', () => true],
    [/^data\.(filter|clearFilter|reapply)$/, 'autoFilter', () => true],
    [/^insert\.link$/, 'insertHyperlinks', () => true],
    [/^data\.pivot$/, 'pivotTables', () => true],
    [/^insert\.(chart|picture|shape|textBox)$/, 'objects', () => true],
  ];
  // Commands are looked up at run time, so wrapping the registered ones is enough.
  const ids = ['format.bold', 'format.italic', 'format.underline', 'format.strike', 'format.doubleUnderline', 'format.superscript', 'format.subscript', 'format.fontName', 'format.fontSize', 'format.growFont', 'format.shrinkFont',
    'format.fontColor', 'format.fillColor', 'format.wrap', 'format.indent', 'format.outdent', 'format.rotate', 'format.mergeCenter', 'format.mergeAcross', 'format.merge', 'format.unmerge',
    'format.align.left', 'format.align.center', 'format.align.right', 'format.valign.top', 'format.valign.center', 'format.valign.bottom',
    'number.format', 'number.moreDecimals', 'number.lessDecimals', 'number.currency', 'number.percent', 'number.comma', 'number.shortDate', 'number.time',
    'border.all', 'border.outside', 'border.thickBox', 'border.inside', 'border.top', 'border.bottom', 'border.left', 'border.right', 'border.doubleBottom', 'border.topBottom', 'border.none',
    'cells.styles', 'insert.conditional', 'insert.clearRules', 'insert.clearAllRules', 'edit.clearFormats', 'edit.formatPainter',
    'cells.colWidth', 'cells.autofitCols', 'cells.hideCols', 'cells.unhideCols', 'cells.rowHeight', 'cells.autofitRows', 'cells.hideRows', 'cells.unhideRows',
    'cells.insertCols', 'cells.insertRows', 'cells.deleteCols', 'cells.deleteRows', 'data.sortAsc', 'data.sortDesc', 'data.sort', 'data.filter', 'data.clearFilter', 'data.reapply',
    'insert.link', 'data.pivot', 'insert.chart', 'insert.picture', 'insert.shape', 'insert.textBox'];
  for (const id of ids) {
    const c = command(id);
    const rule = rules.find(([re]) => re.test(id));
    if (!c || !rule) continue;
    const [, flag, applies] = rule;
    const inner = c.run;
    registerCommand({ ...c, run: (s, arg) => { if (blocked(s.sheet, flag) && applies(s)) { s.notify(T('The sheet is protected. Unprotect it to do this.')); return; } inner(s, arg); } });
  }
};

const anyLocked = (s: GridStore): boolean => {
  for (const r of s.sel.ranges) {
    let n = 0;
    for (let row = r.r1; row <= r.r2 && n < 20000; row++) for (let col = r.c1; col <= r.c2 && n < 20000; col++, n++) if (s.isLocked(row, col)) return true;
    if (n >= 20000) return true;
  }
  return false;
};

// ---------- Protect workbook ----------

export const ProtectWorkbookDialog = ({ store, close }: DialogProps) => {
  const p = store.wb.protection;
  const on = !!(p?.structure || p?.windows);
  const [pw, setPw] = useState('');
  const [pw2, setPw2] = useState('');
  const [windows, setWindows] = useState(false);
  const submit = (): boolean => {
    if (on) {
      if (p && hasPassword(p) && !checkPassword(pw, p)) { store.notify(T('That password isn\'t right.')); return false; }
      store.change(N('Unprotect workbook'), 'book', () => { delete store.wb.protection; });
      store.notify(T('Workbook unprotected.'));
      return true;
    }
    if (pw !== pw2) { store.notify(T('The passwords don\'t match.')); return false; }
    store.change(N('Protect workbook'), 'book', () => { store.wb.protection = { structure: true, ...(windows ? { windows: true } : {}), ...password(pw) }; });
    store.notify(T('Workbook structure protected.'));
    return true;
  };
  return (
    <Dialog title={on ? T('Unprotect workbook') : T('Protect workbook')} close={close} width={440} okLabel={on ? T('Unprotect') : T('Protect')} onOk={submit}>
      {on ? <>
        <p>{p && hasPassword(p) ? T('Enter the password to unprotect the workbook structure.') : T('The workbook structure is protected without a password.')}</p>
        {p && hasPassword(p) && <Field label={T('Password')}><input type="password" autoFocus value={pw} onChange={(e) => setPw(e.target.value)} /></Field>}
      </> : <>
        <p className="muted small">{T('Protecting the structure stops people adding, deleting, renaming, moving, copying, hiding or unhiding sheets, and changing tab colors.')}</p>
        <div className="grid2">
          <Field label={T('Password (optional)')}><input type="password" autoFocus value={pw} onChange={(e) => setPw(e.target.value)} autoComplete="new-password" /></Field>
          <Field label={T('Repeat password')}><input type="password" value={pw2} disabled={!pw} onChange={(e) => setPw2(e.target.value)} autoComplete="new-password" /></Field>
        </div>
        <Check label={T('Also keep the window layout (stored for other apps)')} checked={windows} onChange={setWindows} />
        <p className="muted small">{T('This isn\'t encryption: to keep people from opening the file, use Protect with password.')}</p>
      </>}
    </Dialog>
  );
};

// ---------- Encrypt with a password ----------

export const EncryptDialog = ({ store, close }: DialogProps) => {
  const on = !!store.file.password;
  const [pw, setPw] = useState('');
  const [pw2, setPw2] = useState('');
  const submit = (): boolean => {
    if (!pw && on) { store.file = { ...store.file, password: undefined, dirty: true }; store.notify(T('The password will be removed when you save.')); store.emit(); return true; }
    if (!pw) { store.notify(T('Enter a password.')); return false; }
    if (pw !== pw2) { store.notify(T('The passwords don\'t match.')); return false; }
    store.file = { ...store.file, password: pw, dirty: true };
    store.notify(T('The workbook will be encrypted with this password when you save.'));
    store.emit();
    return true;
  };
  return (
    <Dialog title={T('Protect with password')} close={close} width={440} okLabel={on && !pw ? T('Remove password') : T('Set password')} onOk={submit}>
      <div className="b-inline muted small"><ShieldCheck size={15} strokeWidth={1.9} aria-hidden="true" />{store.file.name}</div>
      <p className="muted small">{T('The file is encrypted with AES-256. Anyone opening it needs this password, in ClerdoGrid or in other apps.')}</p>
      <div className="grid2">
        <Field label={on ? T('New password (empty removes it)') : T('Password')}><input type="password" autoFocus value={pw} onChange={(e) => setPw(e.target.value)} autoComplete="new-password" /></Field>
        <Field label={T('Repeat password')}><input type="password" value={pw2} disabled={!pw} onChange={(e) => setPw2(e.target.value)} autoComplete="new-password" /></Field>
      </div>
      <p className="muted small">{T('If you lose the password, the workbook can\'t be opened. Version history isn\'t kept for encrypted workbooks.')}</p>
    </Dialog>
  );
};

/** Asks for the password of an encrypted file being opened. */
export const OpenPasswordDialog = ({ store, close, data }: DialogProps<{ name: string; data: Uint8Array; path: string | null; handle?: unknown }>) => {
  const [pw, setPw] = useState('');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  if (!data) return null;
  const submit = (): boolean => {
    if (busy) return false;
    setBusy(true);
    // Let the dialog paint "Opening…" before the key derivation runs.
    setTimeout(() => {
      let wb: Workbook;
      try {
        wb = readXlsx(decryptWorkbook(data.data, pw));
      } catch (e) {
        setBusy(false);
        setError(e instanceof EncryptionError ? e.message : T('The file could not be opened.'));
        return;
      }
      store.load(wb, data.name, data.path, data.handle);
      store.file.password = pw;
      close();
      store.notify(T('Opened. It stays encrypted with the same password when you save.'));
    }, 30);
    return false;
  };
  return (
    <Dialog title={T('Password required')} close={close} width={400} okLabel={busy ? T('Opening…') : T('Open')} onOk={submit}>
      <p>{T('{name} is encrypted. Enter its password to open it.', { name: data.name })}</p>
      <Field label={T('Password')}><input type="password" autoFocus value={pw} onChange={(e) => { setPw(e.target.value); setError(''); }} /></Field>
      {error && <p className="ws-error" role="alert">{error}</p>}
    </Dialog>
  );
};
