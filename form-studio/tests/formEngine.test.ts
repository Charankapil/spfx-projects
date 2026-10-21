import { test } from 'node:test';
import * as assert from 'node:assert/strict';

import { FormEngine, IRuntimeContext, selectForm } from '../src/common/form/FormEngine';
import { buildPayload, buildSelect, readItem } from '../src/common/form/valueCodec';
import { IFieldInfo } from '../src/common/model/IFieldInfo';
import { createControl, createDefaultForm } from '../src/common/model/formFactory';
import { IFormDefinition } from '../src/common/model/IFormDefinition';

function field(internalName: string, kind: IFieldInfo['kind'], extra: Partial<IFieldInfo> = {}): IFieldInfo {
  return {
    internalName,
    entityPropertyName: internalName,
    title: internalName,
    kind,
    typeAsString: kind,
    description: '',
    required: false,
    readOnly: false,
    ...extra
  };
}

const fields: IFieldInfo[] = [
  field('Title', 'Text', { required: true, maxLength: 20 }),
  field('Status', 'Choice', { choices: ['Draft', 'Submitted', 'Approved'], defaultValue: 'Draft' }),
  field('Amount', 'Currency', { min: 0 }),
  field('Quantity', 'Number'),
  field('Total', 'Number'),
  field('Justification', 'Note'),
  field('Approver', 'User'),
  field('Customer', 'Lookup', { lookupField: 'Title' }),
  field('Tags', 'MultiChoice', { choices: ['A', 'B'] }),
  field('Region', 'TaxonomyFieldType'),
  field('Created', 'DateTime', { readOnly: true })
];

const runtime = (mode: IRuntimeContext['mode']): IRuntimeContext => ({
  mode,
  user: { name: 'Grace', email: 'grace@contoso.com', loginName: 'grace' },
  groups: ['Approvers'],
  itemId: mode === 'New' ? undefined : 5
});

function buildForm(): IFormDefinition {
  const form = createDefaultForm('Test', fields);
  const controls = form.pages[0].sections[0].controls;
  const byField = (name: string) => controls.filter((c) => c.field === name)[0];
  byField('Justification').visibleWhen = '[Amount] > 1000';
  byField('Justification').requiredWhen = 'true';
  byField('Total').calculatedValue = '[Amount] * [Quantity]';
  byField('Quantity').defaultValue = '1';
  byField('Approver').defaultValue = 'UserEmail()';
  byField('Approver').readOnlyWhen = 'Not(InGroup("Approvers"))';
  byField('Title').validations = [{ expression: 'Not(Contains([Title], "test"))', message: 'No tests' }];
  form.settings.validations = [{ expression: '[Amount] < 100000', message: 'Too much' }];
  form.settings.onSave = [{ field: 'Status', expression: 'If(FormMode() = "New", "Submitted", [Status])' }];
  return form;
}

test('new-item defaults: SharePoint default, expression default, unresolved user', () => {
  const engine = new FormEngine(buildForm(), fields, runtime('New'));
  const values = engine.initialValues({});
  assert.equal(values.Status, 'Draft');
  assert.equal(values.Quantity, 1);
  assert.deepEqual(values.Approver, { id: 0, title: 'grace@contoso.com', email: 'grace@contoso.com', unresolved: true });
  assert.deepEqual(values.Tags, []);
});

test('calculated values settle and rules drive visibility / required', () => {
  const engine = new FormEngine(buildForm(), fields, runtime('New'));
  let values = engine.initialValues({});
  values = engine.applyCalculated({ ...values, Amount: 1500, Quantity: 2 });
  assert.equal(values.Total, 3000);
  const state = engine.computeState(values);
  const justification = engine.fieldControls().filter((x) => x.field.internalName === 'Justification')[0].control;
  const total = engine.fieldControls().filter((x) => x.field.internalName === 'Total')[0].control;
  assert.deepEqual(state.controls[justification.id], { visible: true, readOnly: false, required: true });
  assert.equal(state.controls[total.id].readOnly, true);
  assert.equal(engine.computeState({ ...values, Amount: 10 }).controls[justification.id].visible, false);
});

test('validation: required, max length, min, custom rule, hidden fields skipped', () => {
  const engine = new FormEngine(buildForm(), fields, runtime('New'));
  const values = engine.initialValues({});
  const id = (name: string) => engine.fieldControls().filter((x) => x.field.internalName === name)[0].control.id;

  let errors = engine.validate({ ...values, Amount: 50 });
  assert.match(errors[id('Title')], /blank/);
  assert.equal(errors[id('Justification')], undefined, 'hidden field is not validated');

  errors = engine.validate({ ...values, Title: 'this is a very long title indeed', Amount: -1 });
  assert.match(errors[id('Title')], /20 characters/);
  assert.match(errors[id('Amount')], /0 or more/);

  errors = engine.validate({ ...values, Title: 'my test', Amount: 5000 });
  assert.equal(errors[id('Title')], 'No tests');
  assert.match(errors[id('Justification')], /blank/);

  assert.deepEqual(engine.formErrors({ ...values, Amount: 200000 }), ['Too much']);
  assert.deepEqual(engine.formErrors({ ...values, Amount: 2 }), []);
});

test('display mode makes everything read-only and skips calculation', () => {
  const engine = new FormEngine(buildForm(), fields, runtime('Display'));
  const values = engine.initialValues({ Amount: 2, Quantity: 3, Total: 1 });
  assert.equal(values.Total, 1);
  const state = engine.computeState(values);
  assert.ok(Object.keys(state.controls).every((k) => state.controls[k].readOnly));
});

test('assignments and writable fields', () => {
  const engine = new FormEngine(buildForm(), fields, runtime('New'));
  const values = engine.applyAssignments(engine.form.settings.onSave, engine.initialValues({}));
  assert.equal(values.Status, 'Submitted');
  const writable = engine.writableFields([{ field: 'Created', expression: 'Now()' }]).map((f) => f.internalName);
  assert.ok(writable.indexOf('Status') >= 0);
  assert.ok(writable.indexOf('Created') < 0, 'read-only fields are never written');
});

test('REST mapping round trip and change-only payload', () => {
  const sel = buildSelect(fields);
  assert.ok(sel.select.indexOf('Approver/EMail') >= 0);
  assert.ok(sel.expand.indexOf('Customer') >= 0);
  assert.ok(sel.expand.indexOf('FieldValuesAsText') >= 0);

  const raw = {
    Id: 5,
    Title: 'Laptop',
    Status: 'Draft',
    Amount: 10,
    Quantity: null,
    Approver: { Id: 3, Title: 'Ada', EMail: 'ada@contoso.com', Name: 'i:0#.f|membership|ada' },
    Customer: { Id: 9, Title: 'Contoso' },
    Tags: ['A'],
    Region: { TermGuid: 'g-1', Label: '12', WssId: 12 },
    FieldValuesAsText: { Region: 'Europe' },
    Created: '2026-10-01T10:00:00Z'
  };
  const original = readItem(fields, raw);
  assert.deepEqual(original.Customer, { id: 9, title: 'Contoso' });
  assert.deepEqual(original.Region, { termGuid: 'g-1', label: 'Europe' });
  assert.ok(original.Created instanceof Date);

  const edited = { ...original, Amount: 20, Customer: { id: 4, title: 'Fabrikam' }, Region: { termGuid: 'g-2', label: 'Asia' } };
  const payload = buildPayload(fields, edited, original);
  assert.deepEqual(payload.body, { Amount: 20, CustomerId: 4 });
  assert.deepEqual(payload.taxonomy, [{ FieldName: 'Region', FieldValue: 'Asia|g-2' }]);

  const fresh = buildPayload(fields, { Title: 'x', Tags: ['A', 'B'], Approver: { id: 0, title: 'a@b.c', email: 'a@b.c', unresolved: true } });
  assert.deepEqual(fresh.body, { Title: 'x', Tags: ['A', 'B'], ApproverId: 0 });
  assert.deepEqual(fresh.unresolvedUsers, ['a@b.c']);
});

test('form selection honours order, mode, audience and content type', () => {
  const a = createDefaultForm('Approvers', fields);
  a.audience = ['Approvers'];
  a.modes = ['Edit'];
  const b = createDefaultForm('Everyone', fields);
  const c = createDefaultForm('Documents', fields);
  c.contentTypes = ['0x0101'];
  assert.equal(selectForm([a, b], 'Edit', ['approvers'])!.name, 'Approvers');
  assert.equal(selectForm([a, b], 'Edit', ['Visitors'])!.name, 'Everyone');
  assert.equal(selectForm([a, b], 'New', ['Approvers'])!.name, 'Everyone');
  assert.equal(selectForm([c], 'Edit', [], '0x0100ABC'), undefined);
  assert.equal(selectForm([c], 'Edit', [], '0x0101009'), c);
});

test('createControl gives sensible spans', () => {
  assert.equal(createControl('field', fields[5]).span, 2);
  assert.equal(createControl('heading').span, 12);
});

test('designer operations: move, nudge, pages and undo history', async () => {
  const ops = await import('../src/designer/designerOps');
  const form = createDefaultForm('Ops', fields);
  const section = form.pages[0].sections[0];
  const [first, second, third] = section.controls.map((c) => c.id);

  // Moving down within a section lands before the control that was at the target index.
  let moved = ops.moveControl(form, first, section.id, 3);
  assert.deepEqual(moved.pages[0].sections[0].controls.slice(0, 3).map((c) => c.id), [second, third, first]);
  assert.equal(form.pages[0].sections[0].controls[0].id, first, 'original is untouched');

  const added = ops.addSection(moved, moved.pages[0].id);
  moved = ops.moveControl(added.form, second, added.sectionId);
  assert.equal(ops.findControl(moved, second)!.section.id, added.sectionId);

  const paged = ops.addPage(moved);
  assert.equal(paged.form.layout, 'tabs');
  const back = ops.removePage(ops.nudgeSection(paged.form, added.sectionId, 1), paged.pageId);
  assert.equal(back.pages.length, 1);
  assert.ok(ops.findControl(back, second), 'controls survive page removal');
  assert.equal(ops.placedFields(back).length, ops.placedFields(form).length);

  let history = { past: [] as number[], present: 1, future: [] as number[] };
  history = ops.pushHistory(history, 2);
  history = ops.pushHistory(history, 3);
  history = ops.undo(history);
  assert.equal(history.present, 2);
  history = ops.redo(history);
  assert.equal(history.present, 3);
  history = ops.pushHistory(ops.undo(history), 9);
  assert.deepEqual(history.future, []);
});

test('text typed into a number field is reported as not a number', () => {
  const engine = new FormEngine(buildForm(), fields, runtime('New'));
  const values = engine.initialValues({});
  const amount = engine.fieldControls().filter((x) => x.field.internalName === 'Amount')[0].control.id;
  assert.equal(engine.validate({ ...values, Title: 'x', Amount: NaN })[amount], 'Enter a number.');
});
