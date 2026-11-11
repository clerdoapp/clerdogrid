/** A problem caused by the input (bad arguments, missing file, unreadable workbook): shown without a stack trace. */
export class UserError extends Error {
  constructor(message: string, readonly exitCode: number = EXIT.usage) {
    super(message);
  }
}

export const EXIT = {
  ok: 0,
  /** The command ran and found something: differences, formula problems, no matches. */
  found: 1,
  usage: 2,
  file: 3,
  internal: 70,
} as const;
