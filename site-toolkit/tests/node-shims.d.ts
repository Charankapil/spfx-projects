// The Node typings that SPFx 1.18 pins predate the "node:" module names.
declare module 'node:test' {
  export function test(name: string, fn: () => void | Promise<void>): void;
}

declare module 'node:assert/strict' {
  interface IAssert {
    (value: unknown, message?: string): void;
    ok(value: unknown, message?: string): void;
    equal(actual: unknown, expected: unknown, message?: string): void;
    deepEqual(actual: unknown, expected: unknown, message?: string): void;
    match(value: string, pattern: RegExp, message?: string): void;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    rejects(promise: Promise<unknown>, expected?: RegExp | ((err: any) => boolean)): Promise<void>;
  }
  const assert: IAssert;
  export default assert;
}
