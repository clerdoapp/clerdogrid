import { describe, expect, it } from 'vitest';
import { synthesize, type FlashExample } from '../src/app/flashfill';

/** Learns from the first `k` rows and returns the predictions for the rest. */
const fill = (rows: (string | string[])[], outs: string[], opt = {}): (string | null)[] => {
  const ins = rows.map((r) => (Array.isArray(r) ? r : [r]));
  const ex: FlashExample[] = outs.map((o, i) => ({ inputs: ins[i]!, output: o }));
  const rest = ins.slice(outs.length);
  const p = synthesize(ex, rest, opt);
  if (!p) return rest.map(() => null);
  return rest.map((r) => p.run(r));
};

describe('Flash Fill: names', () => {
  it('first and last names', () => {
    const names = ['John Smith', 'Mary Ann Jones', 'Bob Lee', 'Alice Walker', 'Zoë Ödegaard'];
    expect(fill(names, ['John', 'Mary'])).toEqual(['Bob', 'Alice', 'Zoë']);
    expect(fill(names, ['Smith', 'Jones'])).toEqual(['Lee', 'Walker', 'Ödegaard']);
  });
  it('reordering with punctuation and case', () => {
    const names = ['john smith', 'mary jones', 'bob lee', 'alice walker'];
    expect(fill(names, ['Smith, John', 'Jones, Mary'])).toEqual(['Lee, Bob', 'Walker, Alice']);
    expect(fill(names, ['JOHN', 'MARY'])).toEqual(['BOB', 'ALICE']);
    expect(fill(names, ['J. Smith', 'M. Jones'])).toEqual(['B. Lee', 'A. Walker']);
  });
  it('initials across different word counts', () => {
    const names = ['John Ronald Reuel Tolkien', 'Mary Shelley', 'George Bernard Shaw', 'Ada Lovelace'];
    expect(fill(names, ['JRRT', 'MS'])).toEqual(['GBS', 'AL']);
    expect(fill(names, ['J.R.R.T.', 'M.S.'])).toEqual(['G.B.S.', 'A.L.']);
  });
  it('joins two columns', () => {
    const rows = [['John', 'Smith'], ['Mary', 'Jones'], ['Bob', 'Lee']];
    expect(fill(rows, ['Smith John', 'Jones Mary'])).toEqual(['Lee Bob']);
    expect(fill(rows, ['jsmith', 'mjones'])).toEqual(['blee']);
  });
});

describe('Flash Fill: emails and codes', () => {
  it('email parts', () => {
    const mails = ['john.smith@example.com', 'mary.jones@corp.org', 'bob.lee@mail.net', 'ann.walker@school.edu'];
    expect(fill(mails, ['john', 'mary'])).toEqual(['bob', 'ann']);
    expect(fill(mails, ['example.com', 'corp.org'])).toEqual(['mail.net', 'school.edu']);
    expect(fill(mails, ['John Smith', 'Mary Jones'])).toEqual(['Bob Lee', 'Ann Walker']);
  });
  it('builds emails from names', () => {
    const names = ['John Smith', 'Mary Jones', 'Bob Lee'];
    expect(fill(names, ['john.smith@acme.com', 'mary.jones@acme.com'])).toEqual(['bob.lee@acme.com']);
  });
  it('product codes', () => {
    const codes = ['AB-1234-X', 'CD-5678-Y', 'EF-9012-Z', 'GHI-77-Q'];
    expect(fill(codes, ['1234', '5678'])).toEqual(['9012', '77']);
    expect(fill(codes, ['X-AB', 'Y-CD'])).toEqual(['Z-EF', 'Q-GHI']);
  });
  it('numbers inside text', () => {
    const rows = ['Order 17 shipped', 'Order 204 shipped', 'Order 3 pending'];
    expect(fill(rows, ['00017', '00204'])).toEqual(['00003']);
    const prices = ['Price: 1234.5 USD', 'Price: 99 USD', 'Price: 7.125 USD'];
    expect(fill(prices, ['1,234.50', '99.00'])).toEqual(['7.13']);
  });
});

describe('Flash Fill: phone numbers and dates', () => {
  it('phone numbers', () => {
    const phones = ['5551234567', '5559876543', '2125550199'];
    expect(fill(phones, ['(555) 123-4567', '(555) 987-6543'])).toEqual(['(212) 555-0199']);
    const dashed = ['555-123-4567', '555-987-6543', '212-555-0199'];
    expect(fill(dashed, ['5551234567', '5559876543'])).toEqual(['2125550199']);
    expect(fill(dashed, ['4567', '6543'])).toEqual(['0199']);
  });
  it('date parts and reformatting', () => {
    const dates = ['2024-03-05', '2023-12-25', '2022-07-14'];
    expect(fill(dates, ['03/05/2024', '12/25/2023'])).toEqual(['07/14/2022']);
    expect(fill(dates, ['March', 'December'])).toEqual(['July']);
    expect(fill(dates, ['Tue', 'Mon'])).toEqual(['Thu']);
    expect(fill(dates, ['5 Mar 2024', '25 Dec 2023'])).toEqual(['14 Jul 2022']);
    expect(fill(['3/5/2024', '12/25/2023', '7/4/2022'], ['2024', '2023'])).toEqual(['2022']);
    expect(fill(['5/3/2024', '25/12/2023', '4/7/2022'], ['March', 'December'], { dateOrder: 'DMY' })).toEqual(['July']);
  });
});

describe('Flash Fill: more patterns', () => {
  it('handles assorted real-world columns', () => {
    expect(fill(['John Smith', 'Mary Jones', 'Bob Lee'], ['jsmith@acme.com', 'mjones@acme.com'])).toEqual(['blee@acme.com']);
    expect(fill(['ID: 4521 (active)', 'ID: 87 (closed)', 'ID: 123456 (active)'], ['4521', '87'])).toEqual(['123456']);
    expect(fill(['+1 (555) 123-4567', '+44 (20) 7946-0958', '+33 (1) 4220-1234'], ['555', '20'])).toEqual(['1']);
    expect(fill(['smith, john', 'jones, mary', 'lee, bob'], ['John Smith', 'Mary Jones'])).toEqual(['Bob Lee']);
    expect(fill(['2024-03-05', '1999-12-31', '2010-01-09'], ['Q1 2024', 'Q4 1999'])).toEqual(['Q1 2010']);
    expect(fill(['Hello World Foo', 'Alpha Beta Gamma', 'One Two Three'], ['Foo', 'Gamma'])).toEqual(['Three']);
    expect(fill(['a-b-c-d', 'e-f-g-h', 'i-j-k-l'], ['d c b a', 'h g f e'])).toEqual(['l k j i']);
    expect(fill(['Dr. Ann Lee MD', 'Mr. Bob Stone', 'Ms. Cy Twombly PhD'], ['Ann Lee', 'Bob Stone'])).toEqual(['Cy Twombly']);
    expect(fill(['$1,234.50', '$99.00', '$5.25'], ['1234.5', '99'])).toEqual(['5.25']);
  });
});

describe('Flash Fill: disambiguation', () => {
  it('a second example settles an ambiguous first one', () => {
    const rows = ['aa bb', 'cc dd ee', 'ff gg'];
    expect(fill(rows, ['bb', 'dd'])).toEqual(['gg']);
    expect(fill(rows, ['bb', 'ee'])).toEqual(['gg']);
  });
  it('one example generalises by tokens, not offsets', () => {
    expect(fill(['Kim Lee', 'Alexandra Brown', 'Jo Park'], ['Lee'])).toEqual(['Brown', 'Park']);
  });
  it('returns null when examples contradict', () => {
    expect(synthesize([{ inputs: ['abc'], output: 'x' }, { inputs: ['abc'], output: 'y' }])).toBeNull();
  });
});
