import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import { Camera, Check, Copy, Crosshair, Eye, LogOut, Pencil, Share2, UserPlus, Users, X } from 'lucide-react';
import type { GridStore } from '../../app/store';
import { openDialog } from '../../app/commands';
import { Dialog, Field, type DialogProps } from '../../ui/dialogs/Dialog';
import { paneFor, panesOf, rectIn, type Viewport } from '../../ui/render';
import { addr, rangeAddr } from '../../model/workbook';
import { encodeQr, qrPath } from './qr';
import { answerInvite, checkTurn, createInvite, DEFAULT_STUN, joinUrl, loadTurn, saveTurn, STUN_KEY, type Invite, type Pairing } from './rtc';
import { LINK_LIMIT, makeShareLink, shareFile } from './share';
import { activeSession, endSession, initials, saveName, savedName, startSession, ui } from './state';
import type { Session } from './session';
import { N, T, Tr } from '../../ui/i18n';

const S = 1.9;

// ---------- Small parts ----------

const copyText = async (text: string, el?: HTMLTextAreaElement | null): Promise<boolean> => {
  try { await navigator.clipboard.writeText(text); return true; } catch {
    if (!el) return false;
    el.select();
    try { return document.execCommand('copy'); } catch { return false; }
  }
};

const QrView = ({ text, label }: { text: string; label: string }) => {
  const q = useMemo(() => { try { return encodeQr(text); } catch { return null; } }, [text]);
  const [big, setBig] = useState(false);
  if (!q) return null;
  const n = q.size + 8;
  return (
    <svg className={`cg-qr${big ? ' big' : ''}`} viewBox={`0 0 ${n} ${n}`} role="img" aria-label={Tr(label)} shapeRendering="crispEdges" onClick={() => setBig(!big)}>
      <title>{big ? T('Click to make smaller') : T('Click to enlarge')}</title>
      <rect width={n} height={n} fill="#fff" />
      <path d={qrPath(q)} fill="#000" />
    </svg>
  );
};

const CodeBox = ({ code, label, store }: { code: string; label: string; store: GridStore }) => {
  const ref = useRef<HTMLTextAreaElement>(null);
  const [copied, setCopied] = useState(false);
  return (
    <div className="cg-code">
      <textarea ref={ref} readOnly value={code} aria-label={Tr(label)} rows={3} spellCheck={false} onFocus={(e) => e.currentTarget.select()} />
      <button className="cg-copy" onClick={async () => { const ok = await copyText(code, ref.current); setCopied(ok); if (!ok) store.notify(T('Select the code and copy it.')); }}>
        {copied ? <Check size={15} strokeWidth={S} /> : <Copy size={15} strokeWidth={S} />} {copied ? T('Copied') : T('Copy')}
      </button>
    </div>
  );
};

type Detector = { detect: (src: CanvasImageSource) => Promise<{ rawValue: string }[]> };

/** Reads a QR code with the camera where the browser has a barcode reader. */
const ScanButton = ({ onCode, label }: { onCode: (text: string) => void; label: string }) => {
  const [on, setOn] = useState(false);
  const video = useRef<HTMLVideoElement>(null);
  const Ctor = (globalThis as { BarcodeDetector?: new (o: { formats: string[] }) => Detector }).BarcodeDetector;
  useEffect(() => {
    if (!on || !Ctor) return;
    let stream: MediaStream | null = null, timer: ReturnType<typeof setInterval> | null = null, alive = true;
    const det = new Ctor({ formats: ['qr_code'] });
    navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' } }).then((s) => {
      if (!alive) { s.getTracks().forEach((t) => t.stop()); return; }
      stream = s;
      if (video.current) { video.current.srcObject = s; void video.current.play(); }
      timer = setInterval(async () => {
        if (!video.current || video.current.readyState < 2) return;
        try { const r = await det.detect(video.current); if (r[0]?.rawValue) { onCode(r[0].rawValue); setOn(false); } } catch { /* keep trying */ }
      }, 300);
    }, () => setOn(false));
    return () => { alive = false; if (timer) clearInterval(timer); stream?.getTracks().forEach((t) => t.stop()); };
  }, [on, Ctor, onCode]);
  if (!Ctor || !navigator.mediaDevices?.getUserMedia) return null;
  return (
    <div className="cg-scan">
      <button onClick={() => setOn(!on)}><Camera size={15} strokeWidth={S} /> {on ? T('Stop camera') : label}</button>
      {on && <video ref={video} muted playsInline className="cg-video" />}
    </div>
  );
};

const NetworkField = ({ host = false }: { host?: boolean }) => {
  const [value, setValue] = useState(() => { try { return localStorage.getItem(STUN_KEY) ?? DEFAULT_STUN; } catch { return DEFAULT_STUN; } });
  const [turn, setTurn] = useState(loadTurn);
  const put = (patch: Partial<typeof turn>) => { const t = { ...turn, ...patch }; setTurn(t); saveTurn(t); };
  const bad = turn.urls.trim() !== '' && !checkTurn(turn);
  return (
    <details className="cg-net" open={!!turn.urls || undefined}>
      <summary>{T('Network')}</summary>
      <Field label={T('STUN server (helps devices on different networks find each other; leave empty for the same network only)')}>
        <input value={value} spellCheck={false} placeholder="stun:host:port" onChange={(e) => { setValue(e.target.value); try { localStorage.setItem(STUN_KEY, e.target.value.trim()); } catch { /* blocked */ } }} />
      </Field>
      <p className="muted cg-small">{T('A TURN server relays the connection when devices can\'t reach each other directly (strict firewalls). Use your own; its login is kept on this device.')}</p>
      <div className="cg-turn">
        <Field label={T('TURN server')}><input value={turn.urls} spellCheck={false} placeholder="turn:host:3478" aria-invalid={bad || undefined} onChange={(e) => put({ urls: e.target.value })} /></Field>
        <Field label={T('User name')}><input value={turn.username} spellCheck={false} autoComplete="off" onChange={(e) => put({ username: e.target.value })} /></Field>
        <Field label={T('Password')}><input type="password" value={turn.credential} autoComplete="off" onChange={(e) => put({ credential: e.target.value })} /></Field>
      </div>
      {bad && <p className="cg-error">{T('TURN addresses start with turn: or turns:, for example turn:relay.example.com:3478.')}</p>}
      {host && <label className="b-check"><input type="checkbox" checked={turn.share} onChange={(e) => put({ share: e.target.checked })} /> <span>{T('Put this TURN server and its login in my invite codes, so guests can use it')}</span></label>}
    </details>
  );
};

const ErrorLine = ({ text }: { text: string }) => (text ? <p className="cg-error" role="alert">{Tr(text)}</p> : null);

// ---------- Dialogs ----------

export const StartDialog = ({ store, close }: DialogProps) => {
  const [name, setName] = useState(savedName());
  return (
    <Dialog title={T('Start a live session')} okLabel={T('Start')} close={close} onOk={() => {
      saveName(name);
      startSession(store, 'host', name);
      setTimeout(() => openDialog(store, 'collab.invite'), 0);
    }}>
      <p>{T('Edit this workbook together with people on other devices. You host the session, so keep this window open while others are here.')}</p>
      <Field label={T('Your name')}><input autoFocus value={name} maxLength={60} placeholder={T('Name others will see')} onChange={(e) => setName(e.target.value)} /></Field>
      <NetworkField host />
    </Dialog>
  );
};

export const InviteDialog = ({ store, close }: DialogProps) => {
  const s = activeSession();
  const [code, setCode] = useState('');
  const [reply, setReply] = useState('');
  const [phase, setPhase] = useState<'making' | 'ready' | 'connecting'>('making');
  const [error, setError] = useState('');
  const [viewOnly, setViewOnly] = useState(false);
  const inv = useRef<Invite | null>(null);
  const ro = useRef(false);
  ro.current = viewOnly;
  const closeRef = useRef(close);
  closeRef.current = close;
  useEffect(() => {
    let alive = true, used = false;
    createInvite().then((i) => {
      if (!alive) { i.cancel(); return; }
      inv.current = i;
      setCode(i.code);
      setPhase('ready');
      i.link.then((link) => {
        used = true;
        const cur = activeSession();
        if (!cur || cur.role !== 'host') { link.close(); return; }
        cur.addGuestLink(link, ro.current);
        if (alive) closeRef.current();
      }, (e: Error) => alive && setError(e.message));
    }, (e: Error) => alive && setError(T('Couldn\'t make an invite code: {message}', { message: e.message })));
    return () => { alive = false; if (!used) inv.current?.cancel(); };
  }, []);
  if (!s || s.role !== 'host') return <Dialog title={T('Invite someone')} close={close}><p>{T('Start a live session first.')}</p></Dialog>;
  const connect = () => {
    if (!inv.current || !reply.trim()) { setError(T('Paste the reply code first.')); return false; }
    setError('');
    inv.current.accept(reply).then(() => setPhase('connecting'), (e: Error) => setError(e.message));
    return false;
  };
  return (
    <Dialog title={T('Invite someone')} okLabel={phase === 'connecting' ? T('Connecting…') : T('Connect')} close={close} onOk={connect} wide>
      <div className="cg-steps">
        <div className="cg-step">
          <h3>{T('1. Send this invite code')}</h3>
          {phase === 'making' ? <p className="muted">{T('Making a code…')}</p> : (<>
            <CodeBox code={code} label={T('Invite code')} store={store} />
            <div className="cg-qrrow">
              <QrView text={joinUrl(code)} label={T('QR code with the invite link')} />
              <p className="muted">{T('They can scan this with a phone or tablet camera to open ClerdoGrid with the code filled in, or paste the code under File, Join live session.')}</p>
            </div>
          </>)}
          <label className="b-check"><input type="checkbox" checked={viewOnly} onChange={(e) => setViewOnly(e.target.checked)} /> <span>{T('View only (they can watch but not edit)')}</span></label>
        </div>
        <div className="cg-step">
          <h3>{T('2. Paste their reply code')}</h3>
          <textarea className="cg-reply" rows={4} value={reply} spellCheck={false} placeholder={T('CGA1.…')} aria-label={T('Reply code')} onChange={(e) => setReply(e.target.value)} />
          <ScanButton label={T('Scan their code')} onCode={setReply} />
          {phase === 'connecting' && <p className="muted">{T('Connecting…')}</p>}
          <ErrorLine text={error} />
        </div>
      </div>
      <p className="muted cg-small">{T('Each code works for one person. Codes hold only what the two devices need to find each other; the connection is encrypted.')}</p>
    </Dialog>
  );
};

export const JoinDialog = ({ store, close }: DialogProps) => {
  const lost = activeSession();
  const rejoin = lost && lost.role === 'guest' && lost.status === 'lost' ? lost : null;
  const [name, setName] = useState(rejoin?.me.name ?? savedName());
  const [offer, setOffer] = useState(ui.joinCode);
  const [reply, setReply] = useState('');
  const [phase, setPhase] = useState<'code' | 'making' | 'reply'>('code');
  const [error, setError] = useState('');
  const pairing = useRef<Pairing | null>(null);
  const done = useRef(false);
  const closeRef = useRef(close);
  closeRef.current = close;
  useEffect(() => { ui.joinCode = ''; return () => { if (!done.current) pairing.current?.cancel(); }; }, []);
  const next = () => {
    if (phase !== 'code') return false;
    if (!offer.trim()) { setError(T('Paste the invite code from the host.')); return false; }
    saveName(name);
    setError('');
    setPhase('making');
    answerInvite(offer).then((p) => {
      pairing.current = p;
      setReply(p.code);
      setPhase('reply');
      p.link.then((link) => {
        done.current = true;
        const cur = activeSession();
        const session = cur && cur.role === 'guest' && cur.status === 'lost' ? cur : startSession(store, 'guest', name);
        session.attachHost(link);
        closeRef.current();
      }, (e: Error) => setError(e.message));
    }, (e: Error) => { setError(e.message); setPhase('code'); });
    return false;
  };
  return (
    <Dialog title={rejoin ? T('Rejoin the live session') : T('Join a live session')} okLabel={phase === 'code' ? T('Next') : T('Waiting…')} close={close} onOk={next} wide>
      {phase === 'code' ? (<>
        {!rejoin && <p>{T('Joining opens the host\'s workbook here. Save this one first if you need it.')}</p>}
        <Field label={T('Your name')}><input value={name} maxLength={60} placeholder={T('Name others will see')} onChange={(e) => setName(e.target.value)} /></Field>
        <Field label={T('Invite code from the host')}><textarea autoFocus rows={4} value={offer} spellCheck={false} placeholder={T('CGO1.…')} onChange={(e) => setOffer(e.target.value)} /></Field>
        <ScanButton label={T('Scan the host\'s code')} onCode={setOffer} />
        <NetworkField />
      </>) : phase === 'making' ? <p className="muted">{T('Making your reply code…')}</p> : (<>
        <p>{T('Send this reply code back to the host. You\'ll join as soon as they paste it.')}</p>
        <CodeBox code={reply} label={T('Reply code')} store={store} />
        <div className="cg-qrrow"><QrView text={reply} label={T('QR code with the reply code')} /><p className="muted">{T('The host can scan this if their device has a camera.')}</p></div>
      </>)}
      <ErrorLine text={error} />
    </Dialog>
  );
};

export const ShareDialog = ({ store, close, data }: DialogProps<{ readOnly?: boolean }>) => {
  const [readOnly, setReadOnly] = useState(!!data?.readOnly);
  const [result, setResult] = useState<ReturnType<typeof makeShareLink> | null>(null);
  const ref = useRef<HTMLTextAreaElement>(null);
  const [copied, setCopied] = useState(false);
  const closeRef = useRef(close);
  closeRef.current = close;
  useEffect(() => {
    setResult(null); setCopied(false);
    const t = setTimeout(() => {
      try {
        const r = makeShareLink(store, readOnly);
        setResult(r);
        if ('url' in r) void copyText(r.url).then(setCopied);
      } catch (e) { store.notify(e instanceof Error ? e.message : T('Couldn\'t make a link.')); closeRef.current(); }
    }, 0);
    return () => clearTimeout(t);
  }, [readOnly, store]);
  const mb = (n: number) => `${(n / 1024 / 1024).toFixed(1)} MB`;
  return (
    <Dialog title={T('Share a copy')} close={close}>
      <label className="b-check"><input type="checkbox" checked={readOnly} onChange={(e) => setReadOnly(e.target.checked)} /> <span>{T('Read-only copy')}</span></label>
      {!result ? <p className="muted">{T('Making the link…')}</p> : 'url' in result ? (<>
        <p>{copied ? T('Link copied.') : T('Copy this link.')}{' '}{readOnly ? T('Anyone with it can open a read-only copy of the workbook as it is now.') : T('Anyone with it can open an editable copy of the workbook as it is now.')}{' '}{T('The workbook travels inside the link; it isn\'t uploaded anywhere.')}</p>
        <div className="cg-code">
          <textarea ref={ref} readOnly rows={3} value={result.url} aria-label={T('Share link')} onFocus={(e) => e.currentTarget.select()} />
          <button className="cg-copy" onClick={async () => setCopied(await copyText(result.url, ref.current))}>{copied ? <Check size={15} strokeWidth={S} /> : <Copy size={15} strokeWidth={S} />} {copied ? T('Copied') : T('Copy')}</button>
        </div>
        <p className="muted cg-small">{T("Link size {size}. Some chat apps cut long links; share the file if it doesn't open.", { size: mb(result.size) })}</p>
      </>) : <p>{T('This workbook is too large for a link ({size}, the limit is {limit}). Share the file instead.', { size: mb(result.tooBig), limit: mb(LINK_LIMIT) })}</p>}
      <div><button onClick={() => { void shareFile(store).then((r) => { if (r === 'downloaded') store.notify(T('Downloaded the workbook to share.')); }); }}><Share2 size={15} strokeWidth={S} />{' '}{T('Share file')}</button></div>
    </Dialog>
  );
};

// ---------- Top bar: who is here ----------

const Avatar = ({ name, color, away, editing, size = 26 }: { name: string; color: string; away?: boolean; editing?: boolean; size?: number }) => (
  <span className={`cg-av${away ? ' away' : ''}`} style={{ background: color, width: size, height: size }} title={name} aria-hidden="true">
    {initials(name)}{editing && <span className="cg-typing" />}
  </span>
);

const statusText = (s: Session): string => Tr(({ connecting: N('Connecting…'), live: N('Live'), reconnecting: N('Reconnecting…'), lost: N('Connection lost'), ended: N('Ended') })[s.status]);

export const SessionChip = ({ store }: { store: GridStore }) => {
  const s = activeSession();
  const box = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!ui.peopleOpen) return;
    const off = (e: PointerEvent) => { if (box.current && !box.current.contains(e.target as Node)) { ui.peopleOpen = false; store.emit(); } };
    document.addEventListener('pointerdown', off);
    return () => document.removeEventListener('pointerdown', off);
  });
  if (!s) {
    if (!store.readOnly) return null;
    return (
      <span className="cg-ro" role="status"><Eye size={14} strokeWidth={S} />{' '}{T('Read-only')}<button onClick={() => { store.readOnly = false; store.notify(T('You can edit this copy now.')); store.emit(); }}>{T('Edit a copy')}</button>
      </span>
    );
  }
  const people = [s.me, ...s.peers];
  const shown = people.slice(0, 5);
  return (
    <div className="cg-chip" ref={box}>
      <button className={`cg-people ${s.status}`} aria-label={T('Live session: {count, plural, one {# person} other {# people}}. {status}', { count: people.length, status: statusText(s) })} aria-expanded={ui.peopleOpen}
        onClick={() => { ui.peopleOpen = !ui.peopleOpen; store.emit(); }}>
        <span className="cg-dot" aria-hidden="true" />
        {shown.map((p) => <Avatar key={p.id} name={p.id === s.me.id ? `${p.name} (you)` : p.name} color={p.color} away={p.away} editing={!!s.presence.get(p.id)?.editing} />)}
        {people.length > shown.length && <span className="cg-more">+{people.length - shown.length}</span>}
        {s.status !== 'live' && <span className="cg-state">{statusText(s)}</span>}
      </button>
      {ui.peopleOpen && <PeoplePopover store={store} s={s} />}
    </div>
  );
};

const where = (s: Session, id: string): string => {
  const p = s.presence.get(id);
  if (!p) return '';
  const cell = addr(p.active[0], p.active[1]);
  const r = p.ranges[p.ranges.length - 1]!;
  const sel = r[0] === r[2] && r[1] === r[3] ? cell : rangeAddr({ r1: r[0], c1: r[1], r2: r[2], c2: r[3] });
  const sheet = s.store.wb.sheets.find((x) => x.cid === p.sheet)?.name ?? '';
  return p.editing ? T('Editing {ref}', { ref: `${sheet}!${cell}` }) : T('At {ref}', { ref: `${sheet}!${sel}` });
};

const PeoplePopover = ({ store, s }: { store: GridStore; s: Session }) => {
  const row = (id: string, name: string, color: string, extra: ReactNode, away?: boolean, actions?: ReactNode) => (
    <li key={id} className="cg-person">
      <Avatar name={name} color={color} away={away} size={30} />
      <div className="cg-who"><b>{name}</b><span className="muted">{extra}</span></div>
      <div className="cg-acts">{actions}</div>
    </li>
  );
  return (
    <div className="cg-pop" role="dialog" aria-label={T('People in this session')}>
      <div className="cg-pophead"><Users size={15} strokeWidth={S} /> <b>{T('Live session')}</b> <span className={`cg-badge ${s.status}`}>{statusText(s)}</span></div>
      <ul>
        {row(s.me.id, T('{name} (you)', { name: s.me.name }), s.me.color, s.role === 'host' ? T('Host') : s.me.ro ? T('View only') : T('Editing together'))}
        {s.peers.map((p) => row(p.id, p.name, p.color, p.away ? T('Away') : [p.host ? T('Host') : '', p.ro ? T('View only') : '', where(s, p.id)].filter(Boolean).join(' · '), p.away, (<>
          {!p.away && s.presence.has(p.id) && <button className="ico sm" title={T('Go to {name}', { name: p.name })} aria-label={T('Go to {name}', { name: p.name })} onClick={() => s.goTo(p.id)}><Crosshair size={15} strokeWidth={S} /></button>}
          {!p.away && s.presence.has(p.id) && <button className={`cg-follow${s.following === p.id ? ' on' : ''}`} aria-pressed={s.following === p.id} onClick={() => s.follow(s.following === p.id ? null : p.id)}>{s.following === p.id ? T('Following') : T('Follow')}</button>}
          {p.away && s.role === 'host' && <button className="ico sm" title={T('Remove {name}', { name: p.name })} aria-label={T('Remove {name}', { name: p.name })} onClick={() => s.forget(p.id)}><X size={15} strokeWidth={S} /></button>}
        </>)))}
      </ul>
      <div className="cg-popfoot">
        {s.role === 'host' && <button onClick={() => { ui.peopleOpen = false; openDialog(store, 'collab.invite'); }}><UserPlus size={15} strokeWidth={S} />{' '}{T('Invite someone')}</button>}
        {s.role === 'guest' && s.status === 'lost' && <button onClick={() => { ui.peopleOpen = false; openDialog(store, 'collab.join'); }}><UserPlus size={15} strokeWidth={S} />{' '}{T('Rejoin')}</button>}
        <div className="grow" />
        <button onClick={() => { if (s.role === 'host' && s.liveGuests && !confirm(T('End the session for everyone?'))) return; endSession(store); store.notify(s.role === 'host' ? T('Session ended.') : T('You left the session.')); }}>
          <LogOut size={15} strokeWidth={S} /> {s.role === 'host' ? T('End session') : T('Leave')}
        </button>
      </div>
    </div>
  );
};

// ---------- Others' cursors over the grid ----------

export const RemoteCursors = ({ store, vp }: { store: GridStore; vp: Viewport }) => {
  const s = activeSession();
  if (!s || !s.presence.size) return null;
  const panes = panesOf(vp);
  const minR = Math.min(...panes.map((p) => p.r1)), maxR = Math.max(...panes.map((p) => p.r2)) + 1;
  const minC = Math.min(...panes.map((p) => p.c1)), maxC = Math.max(...panes.map((p) => p.c2)) + 1;
  const clamp = (r: { r1: number; c1: number; r2: number; c2: number }) => ({ r1: Math.max(r.r1, minR - 1, 0), c1: Math.max(r.c1, minC - 1, 0), r2: Math.min(r.r2, maxR), c2: Math.min(r.c2, maxC) });
  const visible = (r: { r1: number; c1: number; r2: number; c2: number }) => r.r2 >= minR - 1 && r.r1 <= maxR && r.c2 >= minC - 1 && r.c1 <= maxC;
  const out: ReactNode[] = [];
  const sheetId = store.sheet.cid;
  for (const p of s.presence.values()) {
    if (p.sheet !== sheetId) continue;
    const person = s.roster.find((x) => x.id === p.id);
    if (!person || person.away) continue;
    const color = person.color;
    p.ranges.forEach(([r1, c1, r2, c2], i) => {
      if (r1 === r2 && c1 === c2) return;
      const r = { r1, c1, r2, c2 };
      if (!visible(r)) return;
      const rc = clamp(r);
      const box = rectIn(paneFor(panes, vp, rc.r1, rc.c1), vp, rc);
      out.push(<div key={`${p.id}r${i}`} className="cg-sel" style={{ left: box.x, top: box.y, width: box.w, height: box.h, borderColor: color, background: `${color}1f` }} />);
    });
    const [row, col] = p.active;
    const m = store.mergeAt(row, col) ?? { r1: row, c1: col, r2: row, c2: col };
    if (!visible(m)) continue;
    const box = rectIn(paneFor(panes, vp, m.r1, m.c1), vp, m);
    const editing = !!p.editing;
    out.push(<div key={`${p.id}a`} className={`cg-cur${editing ? ' editing' : ''}`} style={{ left: box.x - 1, top: box.y - 1, width: box.w + 2, height: box.h + 2, borderColor: color }} />);
    out.push(
      <div key={`${p.id}f`} className="cg-flag" style={{ left: box.x - 1, top: Math.max(vp.headerH, box.y - 17), background: color }} data-peer={p.id}>
        {editing && <Pencil size={10} strokeWidth={2.4} />}{person.name}
      </div>,
    );
  }
  return <div className="cg-cursors" style={{ pointerEvents: 'none' }}>{out}</div>;
};

// ---------- File panel ----------

export const SharePanel = ({ store }: { store: GridStore }) => {
  const s = activeSession();
  const b = (label: string, icon: ReactNode, onClick: () => void, disabled = false) => (
    <button className="tool" disabled={disabled} onMouseDown={(e) => e.preventDefault()} onClick={onClick}>{icon}<span>{Tr(label)}</span></button>
  );
  return (
    <div className="toollist">
      {!s && b(T('Start live session'), <Users size={16} strokeWidth={S} />, () => openDialog(store, 'collab.start'))}
      {!s && b(T('Join live session'), <UserPlus size={16} strokeWidth={S} />, () => openDialog(store, 'collab.join'))}
      {s?.role === 'host' && b(T('Invite someone'), <UserPlus size={16} strokeWidth={S} />, () => openDialog(store, 'collab.invite'))}
      {s?.role === 'guest' && s.status === 'lost' && b(T('Rejoin live session'), <UserPlus size={16} strokeWidth={S} />, () => openDialog(store, 'collab.join'))}
      {s && b(s.role === 'host' ? T('End live session') : T('Leave live session'), <LogOut size={16} strokeWidth={S} />, () => endSession(store))}
      {b(T('Copy link'), <Copy size={16} strokeWidth={S} />, () => openDialog(store, 'share.link', { readOnly: false }))}
      {b(T('Copy read-only link'), <Eye size={16} strokeWidth={S} />, () => openDialog(store, 'share.link', { readOnly: true }))}
      {b(T('Share file'), <Share2 size={16} strokeWidth={S} />, () => { void shareFile(store); })}
    </div>
  );
};
