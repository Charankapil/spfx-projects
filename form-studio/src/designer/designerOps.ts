import { createPage, createSection } from '../common/model/formFactory';
import { IControl, IFormDefinition, IFormPage, ISection } from '../common/model/IFormDefinition';

/**
 * Immutable edits on a form definition. Every designer action goes through
 * one of these, so undo / redo is just a stack of previous definitions.
 */

export type Selection =
  | { kind: 'form' }
  | { kind: 'page'; id: string }
  | { kind: 'section'; id: string }
  | { kind: 'control'; id: string };

function clone<T>(value: T): T {
  return JSON.parse(JSON.stringify(value));
}

function move<T>(items: T[], from: number, to: number): T[] {
  const copy = items.slice();
  const [item] = copy.splice(from, 1);
  copy.splice(Math.max(0, Math.min(to, copy.length)), 0, item);
  return copy;
}

export function findSection(form: IFormDefinition, sectionId: string): { page: IFormPage; section: ISection; index: number } | undefined {
  for (const page of form.pages) {
    const index = page.sections.map((s) => s.id).indexOf(sectionId);
    if (index >= 0) {
      return { page, section: page.sections[index], index };
    }
  }
  return undefined;
}

export function findControl(form: IFormDefinition, controlId: string): { page: IFormPage; section: ISection; control: IControl; index: number } | undefined {
  for (const page of form.pages) {
    for (const section of page.sections) {
      const index = section.controls.map((c) => c.id).indexOf(controlId);
      if (index >= 0) {
        return { page, section, control: section.controls[index], index };
      }
    }
  }
  return undefined;
}

/** Internal names of the fields already placed on the form. */
export function placedFields(form: IFormDefinition): string[] {
  const names: string[] = [];
  form.pages.forEach((p) =>
    p.sections.forEach((s) =>
      s.controls.forEach((c) => {
        if (c.type === 'field' && c.field) {
          names.push(c.field);
        }
      })
    )
  );
  return names;
}

export function updateControl(form: IFormDefinition, controlId: string, patch: Partial<IControl>): IFormDefinition {
  const next = clone(form);
  const loc = findControl(next, controlId);
  if (loc) {
    loc.section.controls[loc.index] = { ...loc.control, ...patch };
  }
  return next;
}

export function removeControl(form: IFormDefinition, controlId: string): IFormDefinition {
  const next = clone(form);
  const loc = findControl(next, controlId);
  if (loc) {
    loc.section.controls.splice(loc.index, 1);
  }
  return next;
}

/** Inserts at `index` (end when undefined). */
export function addControl(form: IFormDefinition, sectionId: string, control: IControl, index?: number): IFormDefinition {
  const next = clone(form);
  const loc = findSection(next, sectionId);
  if (loc) {
    const at = index === undefined ? loc.section.controls.length : index;
    loc.section.controls.splice(at, 0, control);
  }
  return next;
}

export function moveControl(form: IFormDefinition, controlId: string, targetSectionId: string, targetIndex?: number): IFormDefinition {
  const source = findControl(form, controlId);
  if (!source) {
    return form;
  }
  let index = targetIndex;
  // Moving down within the same section: removing the control first shifts the target up by one.
  if (index !== undefined && source.section.id === targetSectionId && source.index < index) {
    index--;
  }
  return addControl(removeControl(form, controlId), targetSectionId, source.control, index);
}

export function nudgeControl(form: IFormDefinition, controlId: string, delta: number): IFormDefinition {
  const next = clone(form);
  const loc = findControl(next, controlId);
  if (loc) {
    loc.section.controls = move(loc.section.controls, loc.index, loc.index + delta);
  }
  return next;
}

export function updateSection(form: IFormDefinition, sectionId: string, patch: Partial<ISection>): IFormDefinition {
  const next = clone(form);
  const loc = findSection(next, sectionId);
  if (loc) {
    loc.page.sections[loc.index] = { ...loc.section, ...patch };
  }
  return next;
}

export function addSection(form: IFormDefinition, pageId: string, afterSectionId?: string): { form: IFormDefinition; sectionId: string } {
  const next = clone(form);
  const page = next.pages.filter((p) => p.id === pageId)[0] || next.pages[0];
  const section = createSection(2);
  const after = afterSectionId ? page.sections.map((s) => s.id).indexOf(afterSectionId) : -1;
  page.sections.splice(after >= 0 ? after + 1 : page.sections.length, 0, section);
  return { form: next, sectionId: section.id };
}

export function removeSection(form: IFormDefinition, sectionId: string): IFormDefinition {
  const next = clone(form);
  const loc = findSection(next, sectionId);
  if (loc) {
    loc.page.sections.splice(loc.index, 1);
    if (!loc.page.sections.length) {
      loc.page.sections.push(createSection(2));
    }
  }
  return next;
}

export function nudgeSection(form: IFormDefinition, sectionId: string, delta: number): IFormDefinition {
  const next = clone(form);
  const loc = findSection(next, sectionId);
  if (!loc) {
    return form;
  }
  const target = loc.index + delta;
  if (target >= 0 && target < loc.page.sections.length) {
    loc.page.sections = move(loc.page.sections, loc.index, target);
    return next;
  }
  // Past the first / last section: move to the neighbouring page.
  const pageIndex = next.pages.indexOf(loc.page);
  const otherPage = next.pages[pageIndex + (delta < 0 ? -1 : 1)];
  if (!otherPage) {
    return form;
  }
  loc.page.sections.splice(loc.index, 1);
  if (!loc.page.sections.length) {
    loc.page.sections.push(createSection(2));
  }
  if (delta < 0) {
    otherPage.sections.push(loc.section);
  } else {
    otherPage.sections.unshift(loc.section);
  }
  return next;
}

export function addPage(form: IFormDefinition): { form: IFormDefinition; pageId: string } {
  const next = clone(form);
  const page = createPage(`Page ${next.pages.length + 1}`);
  next.pages.push(page);
  if (next.layout === 'sections') {
    next.layout = 'tabs';
  }
  return { form: next, pageId: page.id };
}

export function updatePage(form: IFormDefinition, pageId: string, patch: Partial<IFormPage>): IFormDefinition {
  const next = clone(form);
  next.pages = next.pages.map((p) => (p.id === pageId ? { ...p, ...patch } : p));
  return next;
}

/** Removing a page keeps its controls by moving its sections to the previous page. */
export function removePage(form: IFormDefinition, pageId: string): IFormDefinition {
  if (form.pages.length <= 1) {
    return form;
  }
  const next = clone(form);
  const index = next.pages.map((p) => p.id).indexOf(pageId);
  if (index < 0) {
    return form;
  }
  const [removed] = next.pages.splice(index, 1);
  const target = next.pages[Math.max(0, index - 1)];
  removed.sections.filter((s) => s.controls.length || s.title).forEach((s) => target.sections.push(s));
  return next;
}

export function nudgePage(form: IFormDefinition, pageId: string, delta: number): IFormDefinition {
  const index = form.pages.map((p) => p.id).indexOf(pageId);
  const target = index + delta;
  if (index < 0 || target < 0 || target >= form.pages.length) {
    return form;
  }
  const next = clone(form);
  next.pages = move(next.pages, index, target);
  return next;
}

export interface IHistory<T> {
  past: T[];
  present: T;
  future: T[];
}

const HISTORY_LIMIT = 100;

export function pushHistory<T>(history: IHistory<T>, next: T): IHistory<T> {
  if (next === history.present) {
    return history;
  }
  return { past: history.past.concat([history.present]).slice(-HISTORY_LIMIT), present: next, future: [] };
}

export function undo<T>(history: IHistory<T>): IHistory<T> {
  if (!history.past.length) {
    return history;
  }
  return {
    past: history.past.slice(0, -1),
    present: history.past[history.past.length - 1],
    future: [history.present].concat(history.future)
  };
}

export function redo<T>(history: IHistory<T>): IHistory<T> {
  if (!history.future.length) {
    return history;
  }
  return { past: history.past.concat([history.present]), present: history.future[0], future: history.future.slice(1) };
}
