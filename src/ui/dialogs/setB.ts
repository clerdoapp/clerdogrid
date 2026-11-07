import type { ComponentType } from 'react';
import type { DialogProps } from './Dialog';
import { FormatCellsDialog } from './format';
import { CellStylesDialog, TableDialog } from './styles';
import { ConditionalDialog } from './conditional';
import { ValidationDialog } from './validation';
import { FindDialog, GotoSpecialDialog } from './find';
import { GoalSeekDialog, PasteSpecialDialog, RemoveDuplicatesDialog, SeriesDialog, SortDialog, TextToColumnsDialog } from './data';
import { DefineNameDialog, ErrorsDialog, EvaluateDialog, InsertFunctionDialog, NamesDialog, TraceDialog } from './formulas';
import { AboutDialog, LinkDialog, NoteDialog, PropertiesDialog, ProtectDialog, ShortcutsDialog, SymbolDialog, ZoomToSelectionDialog } from './misc';
import './dialogs.css';

type AnyDialog = ComponentType<DialogProps<never>>;
const d = (c: unknown): AnyDialog => c as AnyDialog;

export const DIALOGS_B: Record<string, AnyDialog> = {
  formatCells: d(FormatCellsDialog),
  cellStyles: d(CellStylesDialog),
  table: d(TableDialog),
  conditional: d(ConditionalDialog),
  validation: d(ValidationDialog),
  find: d(FindDialog),
  gotoSpecial: d(GotoSpecialDialog),
  pasteSpecial: d(PasteSpecialDialog),
  series: d(SeriesDialog),
  insertFunction: d(InsertFunctionDialog),
  names: d(NamesDialog),
  defineName: d(DefineNameDialog),
  evaluate: d(EvaluateDialog),
  trace: d(TraceDialog),
  errors: d(ErrorsDialog),
  removeDuplicates: d(RemoveDuplicatesDialog),
  textToColumns: d(TextToColumnsDialog),
  goalSeek: d(GoalSeekDialog),
  sort: d(SortDialog),
  link: d(LinkDialog),
  note: d(NoteDialog),
  symbol: d(SymbolDialog),
  protect: d(ProtectDialog),
  zoomToSelection: d(ZoomToSelectionDialog),
  properties: d(PropertiesDialog),
  shortcuts: d(ShortcutsDialog),
  about: d(AboutDialog),
};
