import { UserError } from './errors';

export interface Parsed {
  positional: string[];
  flags: Map<string, string | true>;
  /** Every value given for a repeatable flag, in order. */
  lists: Map<string, string[]>;
}

/**
 * Splits argv into positionals and flags. `valued` lists the flags that take a value
 * ("--sheet Data", "--sheet=Data", "-o out.xlsx"); other flags are switches. "--" ends flags,
 * and a lone "-" or a negative number is a positional.
 */
export const parseArgs = (argv: string[], valued: readonly string[], aliases: Record<string, string> = {}): Parsed => {
  const positional: string[] = [];
  const flags = new Map<string, string | true>();
  const lists = new Map<string, string[]>();
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i]!;
    if (a === '--') { positional.push(...argv.slice(i + 1)); break; }
    if (a === '-' || /^-\d/.test(a) || !a.startsWith('-')) { positional.push(a); continue; }
    let name: string, value: string | undefined;
    if (a.startsWith('--')) {
      const eq = a.indexOf('=');
      name = eq > 0 ? a.slice(2, eq) : a.slice(2);
      value = eq > 0 ? a.slice(eq + 1) : undefined;
    } else {
      name = aliases[a.slice(1, 2)] ?? a.slice(1, 2);
      value = a.length > 2 ? a.slice(2).replace(/^=/, '') : undefined;
    }
    name = aliases[name] ?? name;
    if (valued.includes(name)) {
      if (value === undefined) {
        if (i + 1 >= argv.length) throw new UserError(`--${name} needs a value.`);
        value = argv[++i]!;
      }
      flags.set(name, value);
      lists.set(name, [...(lists.get(name) ?? []), value]);
    } else {
      if (value !== undefined) throw new UserError(`--${name} doesn't take a value.`);
      flags.set(name, true);
    }
  }
  return { positional, flags, lists };
};

export const str = (p: Parsed, name: string): string | undefined => {
  const v = p.flags.get(name);
  return typeof v === 'string' ? v : undefined;
};

export const bool = (p: Parsed, name: string): boolean => p.flags.has(name);

export const int = (p: Parsed, name: string, min = 0, max = Number.MAX_SAFE_INTEGER): number | undefined => {
  const v = str(p, name);
  if (v === undefined) return undefined;
  if (!/^\d+$/.test(v) || +v < min || +v > max) throw new UserError(`--${name} must be a whole number from ${min} to ${max}.`);
  return +v;
};

/** Rejects flags the command doesn't know, so typos don't pass silently. */
export const allowOnly = (p: Parsed, known: readonly string[]): void => {
  for (const k of p.flags.keys()) if (!known.includes(k)) throw new UserError(`Unknown option --${k}. See "clerdogrid help".`);
};
