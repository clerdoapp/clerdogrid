import type { ComponentType } from 'react';
import type { DialogProps } from './Dialog';

/** Dialogs contributed by feature modules (src/features/*). */
export const EXTRA_DIALOGS: Record<string, ComponentType<DialogProps<never>>> = {};
export const registerDialog = <T,>(id: string, c: ComponentType<DialogProps<T>>): void => { EXTRA_DIALOGS[id] = c as ComponentType<DialogProps<never>>; };
