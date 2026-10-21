import { test } from 'node:test';
import * as assert from 'node:assert/strict';

import { checkExpression, evaluate, IExpressionEnv, renderTemplate } from '../src/common/expressions';

const values: { [name: string]: unknown } = {
  Title: 'Laptop request',
  Status: 'Submitted',
  Amount: 1250,
  Quantity: 3,
  Empty: '',
  Tags: ['Hardware', 'Urgent'],
  Customer: { id: 7, title: 'Contoso' },
  Reviewers: [
    { id: 1, title: 'Ada Lovelace', email: 'ada@contoso.com' },
    { id: 2, title: 'Alan Turing', email: 'alan@contoso.com' }
  ],
  DueDate: new Date(2026, 9, 31),
  Approved: false
};

const env: IExpressionEnv = {
  getField: (name) => values[name],
  hasField: (name) => Object.prototype.hasOwnProperty.call(values, name),
  user: { name: 'Grace Hopper', email: 'grace@contoso.com', loginName: 'i:0#.f|membership|grace@contoso.com' },
  groups: ['Finance Approvers', 'Members'],
  mode: 'Edit',
  itemId: 42,
  now: () => new Date(2026, 9, 21, 14, 30)
};

const run = (expr: string): unknown => evaluate(expr, env);

test('literals and arithmetic', () => {
  assert.equal(run('1 + 2 * 3'), 7);
  assert.equal(run('(1 + 2) * 3'), 9);
  assert.equal(run('-[Amount] / 5'), -250);
  assert.equal(run('"it""s"'), 'it"s');
  assert.equal(run("'single'"), 'single');
  assert.equal(run('[Amount] * [Quantity]'), 3750);
});

test('comparisons are blank-aware and case-insensitive for text', () => {
  assert.equal(run('[Status] = "submitted"'), true);
  assert.equal(run('[Status] <> "Draft"'), true);
  assert.equal(run('[Amount] > 1000 && [Amount] <= 1250'), true);
  assert.equal(run('[Empty] = Blank'), true);
  assert.equal(run('IsBlank([Empty])'), true);
  assert.equal(run('[Empty] > 0'), false);
  assert.equal(run('[Approved] = false'), true);
  assert.equal(run('[Amount] = "1250"'), true);
});

test('word operators and Not', () => {
  assert.equal(run('[Amount] > 10 and not [Approved]'), true);
  assert.equal(run('Not([Approved]) Or false'), true);
  assert.equal(run('[Approved] Or ([Amount] > 2000)'), false);
  assert.equal(run('"Urgent" in [Tags]'), true);
  assert.equal(run('"Software" in [Tags]'), false);
});

test('lookup and person values', () => {
  assert.equal(run('[Customer] = "Contoso"'), true);
  assert.equal(run('[Customer].Id'), 7);
  assert.equal(run('Join([Reviewers].Email, ";")'), 'ada@contoso.com;alan@contoso.com');
  assert.equal(run('[Reviewers].Count'), 2);
  assert.equal(run('Contains([Reviewers], "Alan Turing")'), true);
});

test('logic functions', () => {
  assert.equal(run('If([Amount] > 5000, "CFO", [Amount] > 1000, "Manager", "Auto")'), 'Manager');
  assert.equal(run('If(false, 1)'), null);
  assert.equal(run('Switch([Status], "Draft", 1, "Submitted", 2, 0)'), 2);
  assert.equal(run('Coalesce([Empty], [Title])'), 'Laptop request');
  assert.equal(run('IfError(1 / 0, -1)'), -1);
  assert.equal(run('And(true, [Amount] > 1, Or(false, true))'), true);
});

test('text functions', () => {
  assert.equal(run('Len([Title])'), 14);
  assert.equal(run('Upper(Left([Title], 6))'), 'LAPTOP');
  assert.equal(run('Mid([Title], 8, 7)'), 'request');
  assert.equal(run('Replace([Title], "request", "order")'), 'Laptop order');
  assert.equal(run('[Title] & " (" & [Status] & ")"'), 'Laptop request (Submitted)');
  assert.equal(run('IsMatch("AB-1234", "[A-Z]{2}-\\\\d{4}")'), true);
  assert.equal(run('IsMatch("AB-1234x", "[A-Z]{2}-\\\\d{4}")'), false);
  assert.equal(run('IsEmail(UserEmail())'), true);
  assert.equal(run('Text([Amount], "#,##0.00")'), '1,250.00');
  assert.equal(run('Text(0.256, "0.0%")'), '25.6%');
  assert.deepEqual(run('Split("a, b,,c", ",")'), ['a', 'b', 'c']);
});

test('dates', () => {
  assert.equal(run('Text(Today(), "yyyy-MM-dd")'), '2026-10-21');
  assert.equal(run('DateDiff(Today(), [DueDate])'), 10);
  assert.equal(run('Text(DateAdd([DueDate], 1, "months"), "dd MMM yyyy")'), '01 Dec 2026');
  assert.equal(run('[DueDate] > Today()'), true);
  assert.equal(run('Weekday(Date(2026, 10, 25))'), 1);
  assert.equal(run('Text([DueDate] + 7, "d/M/yy")'), '7/11/26');
  assert.equal(run('Text(Now(), "h:mm tt")'), '2:30 PM');
});

test('context functions', () => {
  assert.equal(run('InGroup("finance approvers")'), true);
  assert.equal(run('InGroup("Owners")'), false);
  assert.equal(run('FormMode() = "Edit"'), true);
  assert.equal(run('ItemId()'), 42);
  assert.equal(run('UserName()'), 'Grace Hopper');
});

test('errors are reported with useful messages', () => {
  assert.throws(() => run('[Missing] = 1'), /no field with the internal name \[Missing\]/);
  assert.throws(() => run('Status = 1'), /square brackets/);
  assert.throws(() => run('Foo(1)'), /Unknown function foo/);
  assert.throws(() => run('Len()'), /takes 1 argument/);
  assert.throws(() => run('"open'), /closing quote/);
  assert.throws(() => run('1 +'), /ended unexpectedly/);
});

test('templates and design-time checks', () => {
  assert.equal(renderTemplate('Request #{{ItemId()}}: {{[Title]}}', env), 'Request #42: Laptop request');
  assert.equal(renderTemplate("Title eq '{{\"O'Brien\"}}'", env, (v) => v.replace(/'/g, "''")), "Title eq 'O''Brien'");
  assert.equal(renderTemplate('{{[Missing]}}x', env), 'x');
  assert.deepEqual(checkExpression('[Status] = "A" && [Nope] > 1', ['Status']), { unknownFields: ['Nope'] });
  assert.equal(checkExpression('Frob([Status])', ['Status']).error, 'Unknown function frob()');
  assert.ok(checkExpression('[Status] =', ['Status']).error);
});
