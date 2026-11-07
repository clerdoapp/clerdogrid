import type { ComponentType } from 'react';
import type { DialogProps } from './Dialog';
import { GotoDialog, InsertCellsDialog, PaletteDialog, RenameSheetDialog, SizeDialog, UnhideSheetDialog } from './core';
import { DIALOGS_A } from './setA';
import { DIALOGS_B } from './setB';

type AnyDialog = ComponentType<DialogProps<never>>;

/** Every dialog by id; commands open them with openDialog(store, id, data). */
export const DIALOGS: Record<string, AnyDialog> = {
  palette: PaletteDialog,
  renameSheet: RenameSheetDialog as AnyDialog,
  unhideSheet: UnhideSheetDialog,
  insertCells: InsertCellsDialog as AnyDialog,
  size: SizeDialog as AnyDialog,
  goto: GotoDialog,
  ...DIALOGS_A,
  ...DIALOGS_B,
};
