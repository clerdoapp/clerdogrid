import { expect, it } from 'vitest';
import { protectionHash } from '../src/app/protection';

it('hashes sheet passwords as the file format defines', () => {
  expect(protectionHash('test')).toBe('CBEB');
  expect(protectionHash('Clerdo2026!')).toBe('C13D');
  expect(protectionHash('a')).toBe('CE88');
});
