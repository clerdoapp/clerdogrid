// Sparklines: insert, format, group and clear small in-cell charts.
import { createElement } from 'react';
import { openDialog, registerCommand } from '../../app/commands';
import type { GridStore } from '../../app/store';
import type { SparklineKind } from '../../model/workbook';
import { registerDialog } from '../../ui/dialogs/extra';
import { registerPanelSection } from '../../ui/panelRegistry';
import { registerPanePainter } from '../../ui/render';
import { paintSparklines } from './draw';
import { groupAt, groupsIn } from './model';
import { SparklineDialog, SparklineSection } from './ui';
import { N } from '../../ui/i18n';

const at = (s: GridStore) => groupAt(s.sheet, s.sel.active.row, s.sel.active.col);
const inSel = (s: GridStore, row: number, col: number) => s.sel.ranges.some((r) => row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2);

export const register = (): void => {
  registerPanePainter(paintSparklines);
  registerDialog('sparklines', SparklineDialog);
  registerPanelSection({ tab: 'insert', title: N('Sparklines'), order: 5, render: (s) => createElement(SparklineSection, { store: s }) });
  registerCommand({ id: 'insert.sparkline', label: N('Sparklines…'), group: N('Insert'), run: (s, kind) => openDialog(s, 'sparklines', { kind: (['line', 'column', 'winLoss'].includes(kind as string) ? kind : 'line') as SparklineKind }) });
  registerCommand({ id: 'sparkline.editData', label: N('Edit sparkline data…'), group: N('Sparklines'), enabled: (s) => !!at(s), run: (s) => openDialog(s, 'sparklines', { edit: true }) });
  registerCommand({
    id: 'sparkline.group', label: N('Group sparklines'), group: N('Sparklines'), enabled: (s) => s.sel.ranges.reduce((n, r) => n + groupsIn(s.sheet, r).length, 0) > 1,
    run: (s) => {
      const main = at(s)?.group;
      s.change(N('Group sparklines'), 'sheet', () => {
        const groups = (s.sheet.sparklines ?? []).filter((g) => g.sparklines.some((sp) => inSel(s, sp.row, sp.col)));
        if (groups.length < 2) return;
        const keep = groups.find((g) => main && g.sparklines[0]?.row === main.sparklines[0]?.row && g.sparklines[0]?.col === main.sparklines[0]?.col) ?? groups[0]!;
        const moved = groups.filter((g) => g !== keep).flatMap((g) => g.sparklines.filter((sp) => inSel(s, sp.row, sp.col)));
        for (const g of groups) if (g !== keep) g.sparklines = g.sparklines.filter((sp) => !inSel(s, sp.row, sp.col));
        keep.sparklines.push(...moved);
        s.sheet.sparklines = s.sheet.sparklines!.filter((g) => g.sparklines.length);
      });
    },
  });
  registerCommand({
    id: 'sparkline.ungroup', label: N('Ungroup sparklines'), group: N('Sparklines'), enabled: (s) => (at(s)?.group.sparklines.length ?? 0) > 1,
    run: (s) => {
      const { row, col } = s.sel.active;
      s.change(N('Ungroup sparklines'), 'sheet', () => {
        const hit = groupAt(s.sheet, row, col);
        if (!hit) return;
        const list = s.sheet.sparklines!;
        const i = list.indexOf(hit.group);
        list.splice(i, 1, ...hit.group.sparklines.map((sp) => ({ ...structuredClone(hit.group), sparklines: [sp] })));
      });
    },
  });
  registerCommand({
    id: 'sparkline.clear', label: N('Clear selected sparklines'), group: N('Sparklines'), enabled: (s) => s.sel.ranges.some((r) => groupsIn(s.sheet, r).length > 0),
    run: (s) => s.change(N('Clear sparklines'), 'sheet', () => {
      s.sheet.sparklines = (s.sheet.sparklines ?? []).map((g) => ({ ...g, sparklines: g.sparklines.filter((sp) => !inSel(s, sp.row, sp.col)) })).filter((g) => g.sparklines.length);
      if (!s.sheet.sparklines.length) delete s.sheet.sparklines;
    }),
  });
  registerCommand({
    id: 'sparkline.clearGroup', label: N('Clear sparkline group'), group: N('Sparklines'), enabled: (s) => !!at(s),
    run: (s) => {
      const { row, col } = s.sel.active;
      s.change(N('Clear sparkline group'), 'sheet', () => {
        const hit = groupAt(s.sheet, row, col);
        if (!hit) return;
        s.sheet.sparklines = s.sheet.sparklines!.filter((g) => g !== hit.group);
        if (!s.sheet.sparklines.length) delete s.sheet.sparklines;
      });
    },
  });
};
