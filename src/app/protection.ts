/**
 * The sheet-protection password hash defined by the file format (ECMA-376), so protected sheets
 * unlock the same way in other apps. It is a short hash: protection prevents accidental edits,
 * it is not security.
 */
export const protectionHash = (password: string): string => {
  let h = 0;
  for (let i = password.length - 1; i >= 0; i--) {
    h = ((h >> 14) & 0x01) | ((h << 1) & 0x7fff);
    h ^= password.charCodeAt(i);
  }
  h = ((h >> 14) & 0x01) | ((h << 1) & 0x7fff);
  h ^= password.length;
  h ^= 0xce4b;
  return h.toString(16).toUpperCase().padStart(4, '0');
};
