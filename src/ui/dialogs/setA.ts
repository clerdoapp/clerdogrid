import type { ComponentType } from 'react';
import type { DialogProps } from './Dialog';
import { ChartDialog } from './a/ChartDialog';
import { DataTableDialog, StatsDialog, SubtotalsDialog } from './a/DataDialogs';
import { ImportCsvDialog, PictureDialog } from './a/FileDialogs';
import { PivotDialog } from './a/PivotDialog';
import { TemplatesDialog } from './a/TemplatesDialog';

type AnyDialog = ComponentType<DialogProps<never>>;

export const DIALOGS_A: Record<string, AnyDialog> = {
  chart: ChartDialog as AnyDialog,
  picture: PictureDialog,
  pivot: PivotDialog,
  subtotals: SubtotalsDialog,
  dataTable: DataTableDialog,
  stats: StatsDialog,
  templates: TemplatesDialog,
  importCsv: ImportCsvDialog,
};
