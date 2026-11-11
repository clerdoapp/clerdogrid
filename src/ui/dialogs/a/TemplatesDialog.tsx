import { CalendarCheck, ChartColumn, GraduationCap, Landmark, ListChecks, Receipt, ReceiptText, Wallet, type LucideIcon } from 'lucide-react';
import '../../charts/charts.css';
import { Dialog, type DialogProps } from '../Dialog';
import { TEMPLATES, type Template } from './templates';
import { T, Tr } from '../../i18n';

const ICONS: Record<Template['icon'], LucideIcon> = {
  budget: Wallet, invoice: ReceiptText, loan: Landmark, todo: ListChecks, sales: ChartColumn, attendance: CalendarCheck, expenses: Receipt, grades: GraduationCap,
};

export const TemplatesDialog = ({ store, close }: DialogProps) => {
  const open = (t: Template) => {
    if (store.file.dirty && !confirm(T('Open this template? Unsaved changes in the current workbook will be lost.'))) return;
    try {
      store.load(t.make(), t.name);
      close();
    } catch {
      store.notify(T('This template couldn\'t be opened.'));
    }
  };
  return (
    <Dialog title={T('Templates')} close={close} wide>
      <div className="tplgrid">
        {TEMPLATES.map((t) => {
          const Icon = ICONS[t.icon];
          return (
            <button key={t.id} className="tpl" onClick={() => open(t)}>
              <Icon size={22} strokeWidth={1.9} />
              <b>{Tr(t.name)}</b>
              <span>{Tr(t.about)}</span>
            </button>
          );
        })}
      </div>
    </Dialog>
  );
};
