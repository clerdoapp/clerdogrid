// Compound File Binary container (MS-CFB): the storage format of .xls files and encrypted workbooks.
// readCfb returns every stream by path ("Workbook", "\u0005SummaryInformation", "Storage/Stream");
// writeCfb builds a version 3 file (512-byte sectors) from such a map.

const SIG = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;
const FATSECT = 0xfffffffd;
const DIFSECT = 0xfffffffc;
const NOSTREAM = 0xffffffff;
const MAX_STREAM = 1 << 30;

export const isCfb = (b: Uint8Array): boolean => b.length >= 512 && SIG.every((x, i) => b[i] === x);

export class CfbError extends Error {}

interface DirEntry { name: string; type: number; left: number; right: number; child: number; start: number; size: number }

/** Reads every stream of a compound file. Storages become path prefixes joined with "/". */
export const readCfb = (bytes: Uint8Array): Map<string, Uint8Array> => {
  if (!isCfb(bytes)) throw new CfbError('Not a compound file.');
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u32 = (o: number) => (o + 4 <= bytes.length ? dv.getUint32(o, true) : FREESECT);
  const shift = dv.getUint16(30, true);
  if (shift !== 9 && shift !== 12) throw new CfbError('Unsupported sector size.');
  const ssz = 1 << shift;
  const miniShift = dv.getUint16(32, true);
  const msz = 1 << (miniShift > 0 && miniShift < shift ? miniShift : 6);
  const nFat = u32(44);
  const firstDir = u32(48);
  const cutoff = u32(56) || 4096;
  const firstMiniFat = u32(60);
  const nMiniFat = u32(64);
  let difat = u32(68);
  const nDifat = u32(72);
  const nSectors = Math.floor((bytes.length - ssz) / ssz) + 1;
  const sectorOff = (s: number) => (s + 1) * ssz;

  const fatSectors: number[] = [];
  for (let i = 0; i < 109 && fatSectors.length < nFat; i++) {
    const s = u32(76 + i * 4);
    if (s < nSectors) fatSectors.push(s);
  }
  const seenDifat = new Set<number>();
  for (let n = 0; n < nDifat && difat < nSectors && !seenDifat.has(difat) && fatSectors.length < nFat; n++) {
    seenDifat.add(difat);
    const off = sectorOff(difat);
    const per = ssz / 4 - 1;
    for (let i = 0; i < per && fatSectors.length < nFat; i++) {
      const s = u32(off + i * 4);
      if (s < nSectors) fatSectors.push(s);
    }
    difat = u32(off + per * 4);
  }
  const fat = new Uint32Array(fatSectors.length * (ssz / 4));
  fatSectors.forEach((s, i) => {
    const off = sectorOff(s);
    for (let j = 0; j < ssz / 4; j++) fat[i * (ssz / 4) + j] = u32(off + j * 4);
  });

  const chain = (start: number, table: Uint32Array, limit: number): number[] => {
    const out: number[] = [];
    const seen = new Set<number>();
    let s = start;
    while (s !== ENDOFCHAIN && s < limit && s < table.length && !seen.has(s)) {
      seen.add(s);
      out.push(s);
      s = table[s]!;
    }
    // a chain that runs past the FAT still yields its last sector
    if (s < limit && s >= table.length && !seen.has(s)) out.push(s);
    return out;
  };
  const readBig = (start: number, size?: number): Uint8Array => {
    const secs = chain(start, fat, nSectors);
    const total = size === undefined ? secs.length * ssz : Math.min(size, secs.length * ssz);
    if (total > MAX_STREAM) throw new CfbError('A stream in this file is too large.');
    const out = new Uint8Array(total);
    let p = 0;
    for (const s of secs) {
      if (p >= total) break;
      const off = sectorOff(s);
      const n = Math.min(ssz, total - p, Math.max(0, bytes.length - off));
      out.set(bytes.subarray(off, off + n), p);
      p += ssz;
    }
    return out;
  };

  const dirBytes = readBig(firstDir);
  const entries: DirEntry[] = [];
  const ddv = new DataView(dirBytes.buffer, dirBytes.byteOffset, dirBytes.byteLength);
  for (let o = 0; o + 128 <= dirBytes.length; o += 128) {
    const nameLen = Math.min(64, ddv.getUint16(o + 64, true));
    let name = '';
    for (let i = 0; i < nameLen - 2; i += 2) name += String.fromCharCode(ddv.getUint16(o + i, true));
    const sizeLo = ddv.getUint32(o + 120, true);
    const sizeHi = ddv.getUint32(o + 124, true);
    entries.push({
      name, type: dirBytes[o + 66]!, left: ddv.getUint32(o + 68, true), right: ddv.getUint32(o + 72, true), child: ddv.getUint32(o + 76, true),
      start: ddv.getUint32(o + 116, true), size: shift === 9 ? sizeLo : sizeHi ? MAX_STREAM + 1 : sizeLo,
    });
  }
  const root = entries[0];
  if (!root || root.type !== 5) throw new CfbError('The file has no root entry.');

  const miniStream = root.start !== ENDOFCHAIN && root.size ? readBig(root.start, root.size) : new Uint8Array(0);
  const miniFatBytes = nMiniFat && firstMiniFat !== ENDOFCHAIN ? readBig(firstMiniFat) : new Uint8Array(0);
  const miniFat = new Uint32Array(miniFatBytes.length / 4);
  const mdv = new DataView(miniFatBytes.buffer, miniFatBytes.byteOffset, miniFatBytes.byteLength);
  for (let i = 0; i < miniFat.length; i++) miniFat[i] = mdv.getUint32(i * 4, true);
  const nMini = Math.floor(miniStream.length / msz);
  const readMini = (start: number, size: number): Uint8Array => {
    const secs = chain(start, miniFat, nMini);
    const total = Math.min(size, secs.length * msz);
    const out = new Uint8Array(total);
    let p = 0;
    for (const s of secs) {
      if (p >= total) break;
      const n = Math.min(msz, total - p);
      out.set(miniStream.subarray(s * msz, s * msz + n), p);
      p += msz;
    }
    return out;
  };

  const out = new Map<string, Uint8Array>();
  const visited = new Set<number>();
  const walkTree = (id: number, prefix: string, depth: number): void => {
    // iterative over siblings, recursive only into storages (bounded depth)
    const stack = [id];
    while (stack.length) {
      const i = stack.pop()!;
      if (i === NOSTREAM || i >= entries.length || visited.has(i)) continue;
      visited.add(i);
      const e = entries[i]!;
      stack.push(e.left, e.right);
      const path = prefix + e.name;
      if (e.type === 2) {
        if (e.size > MAX_STREAM) throw new CfbError('A stream in this file is too large.');
        out.set(path, e.size < cutoff ? readMini(e.start, e.size) : readBig(e.start, e.size));
      } else if (e.type === 1 && depth < 32) walkTree(e.child, `${path}/`, depth + 1);
    }
  };
  walkTree(root.child, '', 0);
  return out;
};

/** Looks a stream up by path, ignoring case as the format does. */
export const cfbGet = (streams: Map<string, Uint8Array>, path: string): Uint8Array | undefined => {
  const hit = streams.get(path);
  if (hit) return hit;
  const want = path.toUpperCase();
  for (const [k, v] of streams) if (k.toUpperCase() === want) return v;
  return undefined;
};

// ---------- Writing ----------

interface Node { name: string; type: 1 | 2 | 5; data?: Uint8Array; children: Node[]; id: number; left: number; right: number; child: number; start: number }

/** Directory order defined by the format: shorter names first, then by upper-cased code units. */
const cmpName = (a: string, b: string): number => {
  if (a.length !== b.length) return a.length - b.length;
  const x = a.toUpperCase(), y = b.toUpperCase();
  return x < y ? -1 : x > y ? 1 : 0;
};

/** Builds a compound file holding the given streams. Paths may contain "/" to create storages. */
export const writeCfb = (streams: Map<string, Uint8Array>): Uint8Array => {
  const root: Node = { name: 'Root Entry', type: 5, children: [], id: 0, left: NOSTREAM, right: NOSTREAM, child: NOSTREAM, start: ENDOFCHAIN };
  for (const [path, data] of streams) {
    const parts = path.split('/').filter((p) => p !== '');
    if (!parts.length) continue;
    let at = root;
    for (let i = 0; i < parts.length; i++) {
      const name = parts[i]!.slice(0, 31);
      const last = i === parts.length - 1;
      let next = at.children.find((c) => c.name.toUpperCase() === name.toUpperCase());
      if (!next) {
        next = { name, type: last ? 2 : 1, children: [], id: 0, left: NOSTREAM, right: NOSTREAM, child: NOSTREAM, start: ENDOFCHAIN };
        at.children.push(next);
      }
      if (last) { next.type = 2; next.data = data; }
      at = next;
    }
  }
  const all: Node[] = [];
  const number = (n: Node): void => { n.id = all.length; all.push(n); n.children.sort((a, b) => cmpName(a.name, b.name)); n.children.forEach(number); };
  number(root);
  // balanced binary tree of each storage's children; all nodes black
  const balance = (list: Node[]): number => {
    if (!list.length) return NOSTREAM;
    const mid = list.length >> 1;
    const n = list[mid]!;
    n.left = balance(list.slice(0, mid));
    n.right = balance(list.slice(mid + 1));
    return n.id;
  };
  for (const n of all) if (n.type !== 2) n.child = balance(n.children);

  const SSZ = 512, MSZ = 64, CUTOFF = 4096;
  const small = all.filter((n) => n.type === 2 && n.data!.length < CUTOFF && n.data!.length > 0);
  const big = all.filter((n) => n.type === 2 && n.data!.length >= CUTOFF);
  let miniLen = 0;
  const miniFatList: number[] = [];
  for (const n of small) {
    const count = Math.ceil(n.data!.length / MSZ);
    n.start = miniLen;
    for (let i = 0; i < count; i++) miniFatList.push(i === count - 1 ? ENDOFCHAIN : miniLen + i + 1);
    miniLen += count;
  }
  const miniStreamBytes = miniLen * MSZ;
  const secs = (n: number) => Math.ceil(n / SSZ);
  const dirSecs = secs(all.length * 128);
  const miniFatSecs = secs(miniFatList.length * 4);
  const miniStreamSecs = secs(miniStreamBytes);
  const bigSecs = big.reduce((t, n) => t + secs(n.data!.length), 0);
  const content = dirSecs + miniFatSecs + miniStreamSecs + bigSecs;
  let nFat = 1, nDifat = 0;
  for (;;) {
    nDifat = nFat > 109 ? Math.ceil((nFat - 109) / 127) : 0;
    const need = Math.ceil((content + nFat + nDifat) / 128);
    if (need <= nFat) break;
    nFat = need;
  }
  const total = nFat + nDifat + content;
  const fat = new Uint32Array(nFat * 128).fill(FREESECT);
  let s = 0;
  for (let i = 0; i < nFat; i++) fat[s++] = FATSECT;
  const difatStart = s;
  for (let i = 0; i < nDifat; i++) fat[s++] = DIFSECT;
  const run = (count: number): number => {
    if (!count) return ENDOFCHAIN;
    const start = s;
    for (let i = 0; i < count; i++, s++) fat[s] = i === count - 1 ? ENDOFCHAIN : s + 1;
    return start;
  };
  const dirStart = run(dirSecs);
  const miniFatStart = run(miniFatSecs);
  const miniStreamStart = run(miniStreamSecs);
  for (const n of big) n.start = run(secs(n.data!.length));
  root.start = miniStreamBytes ? miniStreamStart : ENDOFCHAIN;

  const out = new Uint8Array(SSZ * (total + 1));
  const dv = new DataView(out.buffer);
  out.set(SIG, 0);
  dv.setUint16(24, 0x3e, true);
  dv.setUint16(26, 3, true);
  dv.setUint16(28, 0xfffe, true);
  dv.setUint16(30, 9, true);
  dv.setUint16(32, 6, true);
  dv.setUint32(44, nFat, true);
  dv.setUint32(48, dirStart, true);
  dv.setUint32(56, CUTOFF, true);
  dv.setUint32(60, miniFatSecs ? miniFatStart : ENDOFCHAIN, true);
  dv.setUint32(64, miniFatSecs, true);
  dv.setUint32(68, nDifat ? difatStart : ENDOFCHAIN, true);
  dv.setUint32(72, nDifat, true);
  for (let i = 0; i < 109; i++) dv.setUint32(76 + i * 4, i < nFat ? i : FREESECT, true);
  const off = (sec: number) => (sec + 1) * SSZ;
  for (let d = 0; d < nDifat; d++) {
    const o = off(difatStart + d);
    for (let i = 0; i < 127; i++) {
      const k = 109 + d * 127 + i;
      dv.setUint32(o + i * 4, k < nFat ? k : FREESECT, true);
    }
    dv.setUint32(o + 127 * 4, d === nDifat - 1 ? ENDOFCHAIN : difatStart + d + 1, true);
  }
  for (let i = 0; i < fat.length; i++) dv.setUint32(off(Math.floor(i / 128)) + (i % 128) * 4, fat[i]!, true);
  // directory
  const dirOff = off(dirStart);
  for (let i = 0; i < dirSecs * 4; i++) {
    const o = dirOff + i * 128;
    const n = all[i];
    if (!n) { dv.setUint32(o + 68, NOSTREAM, true); dv.setUint32(o + 72, NOSTREAM, true); dv.setUint32(o + 76, NOSTREAM, true); continue; }
    for (let j = 0; j < n.name.length; j++) dv.setUint16(o + j * 2, n.name.charCodeAt(j), true);
    dv.setUint16(o + 64, (n.name.length + 1) * 2, true);
    out[o + 66] = n.type;
    out[o + 67] = 1;
    dv.setUint32(o + 68, n.left, true);
    dv.setUint32(o + 72, n.right, true);
    dv.setUint32(o + 76, n.child, true);
    const size = n.type === 5 ? miniStreamBytes : n.type === 2 ? n.data!.length : 0;
    dv.setUint32(o + 116, n.type === 1 ? 0 : size ? n.start : ENDOFCHAIN, true);
    dv.setUint32(o + 120, size, true);
  }
  if (miniFatSecs) {
    const o = off(miniFatStart);
    const words = miniFatSecs * 128;
    for (let i = 0; i < words; i++) dv.setUint32(o + i * 4, i < miniFatList.length ? miniFatList[i]! : FREESECT, true);
  }
  if (miniStreamSecs) {
    const o = off(miniStreamStart);
    for (const n of small) out.set(n.data!, o + n.start * MSZ);
  }
  for (const n of big) out.set(n.data!, off(n.start));
  return out;
};
