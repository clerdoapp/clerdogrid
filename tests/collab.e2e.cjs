// End-to-end check of live collaboration in a real browser: two separate browser contexts (two "devices") connect
// over WebRTC with the copy-paste codes, edit together, resolve a concurrent row insert, and show each other's cursor.
// Run with the dev server up:  npx vite --port 5470   then   node tests/collab.e2e.cjs [http://127.0.0.1:5470/] [screenshot dir]
const { chromium } = require('playwright');
const assert = require('node:assert/strict');

const URL = process.argv[2] || 'http://127.0.0.1:5470/';
const OUT = process.argv[3] || '.';

const until = async (page, fn, arg, what, ms = 15000) => {
  try { await page.waitForFunction(fn, arg, { timeout: ms, polling: 50 }); } catch (e) { throw new Error(`Timed out waiting for ${what}`); }
};
const val = (page, r, c) => page.evaluate(([r, c]) => { const v = window.__store.value(r, c); return v && typeof v === 'object' ? String(v) : v; }, [r, c]);

(async () => {
  const browser = await chromium.launch({ args: ['--disable-features=WebRtcHideLocalIpsWithMdns'] });
  const mk = async (name) => {
    const ctx = await browser.newContext({ viewport: { width: 1280, height: 800 } });
    // No STUN server: both "devices" are on this machine. Names are what Options remembers.
    await ctx.addInitScript((n) => { try { localStorage.setItem('clerdogrid.stun', ''); localStorage.setItem('clerdogrid.user', n); } catch {} }, name);
    const page = await ctx.newPage();
    page.on('pageerror', (e) => console.error(`[${name}] page error:`, e.message));
    return page;
  };
  const host = await mk('Hana Host');
  const guest = await mk('Gus Guest');
  try {
    await host.goto(URL);
    await host.waitForSelector('.topbar');
    // Some data to work on.
    await host.evaluate(() => { const s = window.__store; s.change('Seed', [{ area: { r1: 0, c1: 0, r2: 9, c2: 0 } }], () => { for (let r = 0; r < 10; r++) s.enter(r, 0, `r${r + 1}`); }); });

    // Host: File > Share > Start live session.
    await host.click('nav.rail button:has-text("File")');
    await host.click('button.tool:has-text("Start live session")');
    // The host adds a TURN relay and shares it in invites (it isn't running, so direct candidates still connect).
    await host.click('.dialog summary:has-text("Network")');
    await host.fill('input[placeholder="turn:host:3478"]', 'turn:127.0.0.1:3478');
    await host.fill('.cg-turn label:has-text("User name") input', 'team');
    await host.fill('.cg-turn input[type="password"]', 'relay-pass');
    await host.check('label:has-text("Put this TURN server") input');
    await host.screenshot({ path: `${OUT}/collab-start-network.png` });
    await host.click('.dialog button.primary:has-text("Start")');
    const codeBox = host.locator('textarea[aria-label="Invite code"]');
    await codeBox.waitFor({ timeout: 15000 });
    const offer = await codeBox.inputValue();
    assert.match(offer, /^CGO2\./, 'invite carries the TURN relay');
    console.log('invite code length', offer.length);
    assert.ok(await host.locator('svg.cg-qr').count() > 0, 'invite QR code shown');
    await host.screenshot({ path: `${OUT}/collab-invite.png` });
    await host.click('svg.cg-qr');
    await host.locator('svg.cg-qr.big').screenshot({ path: `${OUT}/collab-invite-qr.png` });
    await host.click('svg.cg-qr');

    // Guest opens the join link (as the QR code would), gets a reply code.
    await guest.goto(`${URL}#join=${offer}`);
    await guest.waitForSelector('.dialog:has-text("Join a live session")');
    await guest.click('.dialog button.primary:has-text("Next")');
    const replyBox = guest.locator('textarea[aria-label="Reply code"]');
    await replyBox.waitFor({ timeout: 15000 });
    const reply = await replyBox.inputValue();
    assert.match(reply, /^CGA1\./);

    // Host pastes the reply and connects.
    await host.fill('textarea[aria-label="Reply code"]', reply);
    await host.click('.dialog button.primary:has-text("Connect")');
    await until(guest, () => window.__store.value(4, 0) === 'r5', null, 'guest to receive the workbook');
    await until(host, () => document.querySelectorAll('.cg-people .cg-av').length === 2, null, 'host to show two people');
    await until(guest, () => document.querySelectorAll('.cg-people .cg-av').length === 2, null, 'guest to show two people');
    console.log('connected');

    // Edits both ways, typed into the grid.
    await host.evaluate(() => window.__store.select({ row: 0, col: 1 }));
    await host.locator('.gridwrap').focus();
    await host.keyboard.type('=A1&"!"');
    await host.keyboard.press('Enter');
    await until(guest, () => window.__store.value(0, 1) === 'r1!', null, 'host edit on guest');
    await guest.evaluate(() => window.__store.select({ row: 1, col: 2 }));
    await guest.locator('.gridwrap').focus();
    await guest.keyboard.type('42');
    await guest.keyboard.press('Enter');
    await until(host, () => window.__store.value(1, 2) === 42, null, 'guest edit on host');
    // Bold from the guest, a comment from the host.
    await guest.evaluate(() => { const s = window.__store; s.select({ row: 1, col: 2 }); s.applyStyle({ font: { bold: true } }); });
    await until(host, () => !!window.__store.style(1, 2).font?.bold, null, 'style on host');
    await host.evaluate(() => { const s = window.__store; s.change('New comment', 'sheet', () => { s.sheet.comments = new Map([[0, { id: 'e2e', author: 'Hana Host', text: 'Check this', time: new Date().toISOString(), replies: [] }]]); }); });
    await until(guest, () => window.__store.sheet.comments?.get(0)?.text === 'Check this', null, 'comment on guest');
    console.log('edits sync both ways');

    // Concurrent: host inserts two rows above row 3 while the guest types next to r5 (row 5). Both apply locally first.
    await host.evaluate(() => window.__store.selectRange({ r1: 2, c1: 0, r2: 3, c2: 0 }));
    await host.locator('.gridwrap').focus();
    await guest.evaluate(() => window.__store.select({ row: 4, col: 1 }));
    await Promise.all([
      (async () => { await host.click('.cmdsearch'); await host.keyboard.type('Insert sheet rows'); await host.keyboard.press('Enter'); })(),
      guest.evaluate(() => { const s = window.__store; s.beginEdit('next to r5'); s.commitEdit(); }),
    ]);
    const dump = (p) => p.evaluate(() => Array.from({ length: 9 }, (_, r) => [window.__store.value(r, 0), window.__store.value(r, 1)]));
    for (const [p, who] of [[host, 'host'], [guest, 'guest']]) {
      try { await until(p, () => window.__store.value(6, 0) === 'r5' && window.__store.value(6, 1) === 'next to r5', null, 'x', 4000); } catch { console.log(who, 'host:', JSON.stringify(await dump(host)), 'guest:', JSON.stringify(await dump(guest))); }
      await until(p, () => window.__store.value(6, 0) === 'r5' && window.__store.value(6, 1) === 'next to r5', null, `${who}: row insert and concurrent edit to resolve`);
      assert.equal(await val(p, 2, 0), null);
      assert.equal(await val(p, 4, 0), 'r3');
      assert.equal(await val(p, 0, 1), 'r1!');
    }
    console.log('concurrent insert row resolved: r5 and the guest edit are on row 7 for both');

    // Presence: the guest selects a range; the host sees it with the guest's name.
    await guest.evaluate(() => window.__store.selectRange({ r1: 2, c1: 3, r2: 5, c2: 4 }, { row: 2, col: 3 }));
    await until(host, () => [...document.querySelectorAll('.cg-flag')].some((f) => f.textContent.includes('Gus Guest')), null, 'guest cursor on the host');
    await until(host, () => document.querySelectorAll('.cg-sel').length > 0, null, 'guest selection on the host');
    await guest.evaluate(() => { window.__store.select({ row: 8, col: 2 }); window.__store.beginEdit('typing'); });
    await until(host, () => document.querySelector('.cg-cur.editing') !== null, null, 'editing indicator on the host');
    await host.click('.cg-people');
    await host.waitForSelector('.cg-pop');
    await host.screenshot({ path: `${OUT}/collab-host.png` });
    await guest.evaluate(() => window.__store.cancelEdit());
    // Follow: the host follows the guest to another cell.
    await host.click('.cg-pop .cg-follow');
    await guest.evaluate(() => window.__store.select({ row: 20, col: 5 }));
    await until(host, () => window.__store.sel.active.row === 20 && window.__store.sel.active.col === 5, null, 'host to follow the guest');
    await guest.screenshot({ path: `${OUT}/collab-guest.png` });
    console.log('presence and follow work');

    // Same content on both sides.
    const hash = (p) => p.evaluate(() => JSON.stringify([...window.__store.sheet.cells].map(([k, c]) => [k, c.f ?? c.v]).sort((a, b) => a[0] - b[0])));
    assert.equal(await hash(host), await hash(guest));

    // Leaving: the guest leaves, the host sees one person.
    await guest.click('.cg-people');
    await guest.click('.cg-pop button:has-text("Leave")');
    await until(host, () => document.querySelectorAll('.cg-people .cg-av').length === 1, null, 'host to see the guest leave');
    // Share link: a read-only copy opens from the link alone.
    await host.evaluate(() => { window.__store.view.panel = 'file'; window.__store.emit(); });
    await host.click('button.tool:has-text("Copy read-only link")');
    const linkBox = host.locator('textarea[aria-label="Share link"]');
    await linkBox.waitFor({ timeout: 15000 });
    const link = await linkBox.inputValue();
    assert.match(link, /#wb=[A-Za-z0-9_-]+&ro=1&n=/);
    await host.keyboard.press('Escape');
    const viewer = await guest.context().newPage();
    await viewer.goto(link);
    await until(viewer, () => window.__store.value(6, 0) === 'r5' && window.__store.readOnly && !location.hash, null, 'shared copy to open');
    await viewer.waitForSelector('.cg-ro');
    await viewer.screenshot({ path: `${OUT}/collab-shared-link.png` });
    console.log('share link opens a read-only copy');
    console.log('E2E OK');
  } finally {
    await browser.close();
  }
})().catch((e) => { console.error(e); process.exit(1); });
