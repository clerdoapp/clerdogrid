// Namespace-aware XML parsing for OpenDocument parts. Element and attribute names are normalised to
// the usual ODF prefixes (table:, text:, style:, ...) whatever prefixes the file declares, and mixed
// content keeps its order so text spans and links read correctly.
import { decodeEntities, esc } from './xml';

export const NS: Record<string, string> = {
  office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
  style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
  text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
  table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
  draw: 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
  fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
  xlink: 'http://www.w3.org/1999/xlink',
  dc: 'http://purl.org/dc/elements/1.1/',
  meta: 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
  number: 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0',
  svg: 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
  chart: 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0',
  config: 'urn:oasis:names:tc:opendocument:xmlns:config:1.0',
  of: 'urn:oasis:names:tc:opendocument:xmlns:of:1.2',
  manifest: 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0',
  calcext: 'urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0',
  loext: 'urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0',
  tableooo: 'http://openoffice.org/2009/table',
  clerdo: 'https://clerdo.app/ns/grid/odf/1.0',
};
const PREFIX_OF = new Map(Object.entries(NS).map(([p, u]) => [u, p]));

export interface ONode { name: string; attrs: Record<string, string>; kids: (ONode | string)[] }

const ATTR_RE = /([^\s=/>]+)\s*=\s*("([^"]*)"|'([^']*)')/g;

export const parseOdf = (src: string): ONode => {
  const root: ONode = { name: '#document', attrs: {}, kids: [] };
  const stack: { node: ONode; ns: Map<string, string> }[] = [{ node: root, ns: new Map([['xml', 'http://www.w3.org/XML/1998/namespace']]) }];
  let pos = src.charCodeAt(0) === 0xfeff ? 1 : 0;
  const len = src.length;
  const norm = (q: string, ns: Map<string, string>, isAttr: boolean): string => {
    const i = q.indexOf(':');
    if (i < 0) {
      if (isAttr) return q;
      const def = ns.get('');
      return def && PREFIX_OF.get(def) ? `${PREFIX_OF.get(def)}:${q}` : q;
    }
    const p = q.slice(0, i);
    const uri = ns.get(p);
    const canon = uri ? PREFIX_OF.get(uri) : undefined;
    return canon ? `${canon}:${q.slice(i + 1)}` : q;
  };
  while (pos < len) {
    const lt = src.indexOf('<', pos);
    const top = stack[stack.length - 1]!;
    if (lt < 0) { const t = decodeEntities(src.slice(pos)); if (t) top.node.kids.push(t); break; }
    if (lt > pos) top.node.kids.push(decodeEntities(src.slice(pos, lt)));
    const c = src.charCodeAt(lt + 1);
    if (c === 33) {
      if (src.startsWith('<!--', lt)) { const e = src.indexOf('-->', lt + 4); pos = e < 0 ? len : e + 3; continue; }
      if (src.startsWith('<![CDATA[', lt)) { const e = src.indexOf(']]>', lt + 9); top.node.kids.push(src.slice(lt + 9, e < 0 ? len : e)); pos = e < 0 ? len : e + 3; continue; }
      let depth = 0, i = lt + 2;
      for (; i < len; i++) { const ch = src[i]; if (ch === '[') depth++; else if (ch === ']') depth--; else if (ch === '>' && depth <= 0) break; }
      pos = i + 1;
      continue;
    }
    if (c === 63) { const e = src.indexOf('?>', lt + 2); pos = e < 0 ? len : e + 2; continue; }
    if (c === 47) {
      const e = src.indexOf('>', lt);
      if (stack.length > 1) stack.pop();
      pos = e < 0 ? len : e + 1;
      continue;
    }
    // find tag end outside quotes
    let i = lt + 1, q = 0;
    for (; i < len; i++) { const ch = src.charCodeAt(i); if (q) { if (ch === q) q = 0; } else if (ch === 34 || ch === 39) q = ch; else if (ch === 62) break; }
    const body = src.slice(lt + 1, i);
    const self = body.endsWith('/');
    const inner = self ? body.slice(0, -1) : body;
    const sp = inner.search(/[\s/]/);
    const qname = sp < 0 ? inner : inner.slice(0, sp);
    const raw: [string, string][] = [];
    ATTR_RE.lastIndex = 0;
    let m: RegExpExecArray | null;
    const rest = sp < 0 ? '' : inner.slice(sp);
    while ((m = ATTR_RE.exec(rest))) raw.push([m[1]!, decodeEntities(m[3] ?? m[4] ?? '')]);
    const ns = new Map(top.ns);
    for (const [k, v] of raw) { if (k === 'xmlns') ns.set('', v); else if (k.startsWith('xmlns:')) ns.set(k.slice(6), v); }
    const node: ONode = { name: norm(qname, ns, false), attrs: {}, kids: [] };
    for (const [k, v] of raw) if (k !== 'xmlns' && !k.startsWith('xmlns:')) node.attrs[norm(k, ns, true)] = v;
    top.node.kids.push(node);
    if (!self) stack.push({ node, ns });
    pos = i + 1;
    if (stack.length > 512) throw new Error('The document is nested too deeply.');
  }
  return root;
};

export const elems = (n: ONode | undefined): ONode[] => (n ? (n.kids.filter((k) => typeof k !== 'string') as ONode[]) : []);
export const child = (n: ONode | undefined, name: string): ONode | undefined => elems(n).find((k) => k.name === name);
export const children = (n: ONode | undefined, name: string): ONode[] => elems(n).filter((k) => k.name === name);
export const findDeep = (n: ONode | undefined, name: string): ONode | undefined => {
  if (!n) return undefined;
  for (const k of elems(n)) { if (k.name === name) return k; const f = findDeep(k, name); if (f) return f; }
  return undefined;
};
export const allDeep = (n: ONode | undefined, name: string, out: ONode[] = []): ONode[] => {
  for (const k of elems(n)) { if (k.name === name) out.push(k); allDeep(k, name, out); }
  return out;
};
export const a = (n: ONode | undefined, name: string): string | undefined => n?.attrs[name];

/** Paragraph text with ODF whitespace elements expanded. */
export const odfText = (n: ONode): string => {
  let s = '';
  for (const k of n.kids) {
    if (typeof k === 'string') s += k;
    else if (k.name === 'text:s') s += ' '.repeat(Math.min(1000, Math.max(1, +(k.attrs['text:c'] ?? 1) || 1)));
    else if (k.name === 'text:tab') s += '\t';
    else if (k.name === 'text:line-break') s += '\n';
    else if (k.name === 'office:annotation' || k.name === 'office:annotation-end' || k.name === 'text:note') continue;
    else s += odfText(k);
  }
  return s;
};

export const escText = (s: string): string => esc(s);

/** Text for inside text:p: runs of spaces, tabs and line breaks become ODF elements. */
export const textXml = (s: string): string =>
  s.split('\n').map((line) => esc(line).replace(/\t/g, '<text:tab/>').replace(/ {2,}/g, (m) => ` <text:s text:c="${m.length - 1}"/>`).replace(/^ /, '<text:s/>')).join('<text:line-break/>');

export const attrsXml = (o: Record<string, string | number | boolean | undefined>): string => {
  let s = '';
  for (const k in o) {
    const v = o[k];
    if (v === undefined) continue;
    s += ` ${k}="${esc(typeof v === 'boolean' ? (v ? 'true' : 'false') : String(v))}"`;
  }
  return s;
};
export const tag = (name: string, at: Record<string, string | number | boolean | undefined> = {}, inner?: string): string =>
  inner === undefined || inner === '' ? `<${name}${attrsXml(at)}/>` : `<${name}${attrsXml(at)}>${inner}</${name}>`;

export const NS_DECL = (...prefixes: string[]): string => prefixes.map((p) => ` xmlns:${p}="${NS[p]}"`).join('');
