#!/usr/bin/env node
import { appendFile } from 'node:fs/promises';
import { check } from './check.js';
import { ConfigError, loadConfig } from './config.js';
import { createLogger } from './log.js';
import { createClients, run, RunResult } from './runner.js';

const VERSION = '1.0.0';

/** Exit codes: 0 done / nothing to do, 1 failed, 2 settings problem, 75 not finished yet (the next run carries on). */
export const EXIT_PARTIAL = 75;

const HELP = `SharePoint access report - every site and library a person or group can reach.

Usage:
  access-report run    [options]   Make this month's report (does nothing when it exists).
  access-report check  [options]   Test sign-in, permissions and the report folder.

Options:
  --config <file>     Settings file (JSON). Default: ACCESS_REPORT_CONFIG, then environment variables only.
  --period <yyyy-mm>  Month the report is for. Default: the current month (UTC).
  --force             Make the report again even when this month's exists; ignore saved progress.
  --no-upload         Trial run: write files to --out instead of the SharePoint folder.
  --out <folder>      Folder for --no-upload. Default: ./out
  --limit <n>         Trial run: only the first n sites.
  --verbose           Show names and addresses in the log (masked by default).
  --json              Print the result as JSON on standard output.
  --help, --version

Environment:
  ACCESS_REPORT_TENANT_ID, ACCESS_REPORT_CLIENT_ID, ACCESS_REPORT_FOLDER_URL,
  ACCESS_REPORT_CERTIFICATE (PEM: certificate + private key), ACCESS_REPORT_AUTH (certificate | federated),
  ACCESS_REPORT_PRINCIPALS ("user:a@contoso.com; group:Finance"), ACCESS_REPORT_CONFIG (settings as JSON text)
`;

interface Args {
  command: string;
  config?: string;
  period?: string;
  out?: string;
  limit?: number;
  force: boolean;
  upload: boolean;
  verbose: boolean;
  json: boolean;
}

export function parseArgs(argv: string[]): Args {
  const a: Args = { command: '', force: false, upload: true, verbose: false, json: false };
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    const value = (): string => {
      const eq = arg.indexOf('=');
      const v = eq >= 0 ? arg.substring(eq + 1) : argv[++i];
      if (v === undefined || v === '') {
        throw new ConfigError(`${arg.split('=')[0]} needs a value.`);
      }
      return v;
    };
    const name = arg.split('=')[0];
    switch (name) {
      case '--config':
        a.config = value();
        break;
      case '--period':
        a.period = value();
        break;
      case '--out':
        a.out = value();
        break;
      case '--limit': {
        const n = Number(value());
        if (!Number.isInteger(n) || n < 1) {
          throw new ConfigError('--limit needs a whole number of sites.');
        }
        a.limit = n;
        break;
      }
      case '--force':
        a.force = true;
        break;
      case '--no-upload':
        a.upload = false;
        break;
      case '--verbose':
      case '-v':
        a.verbose = true;
        break;
      case '--json':
        a.json = true;
        break;
      case '--help':
      case '-h':
        a.command = 'help';
        break;
      case '--version':
        a.command = 'version';
        break;
      default:
        if (arg.charAt(0) === '-') {
          throw new ConfigError(`Unknown option ${arg}. Try --help.`);
        }
        if (a.command && a.command !== arg) {
          throw new ConfigError(`Unexpected "${arg}". Try --help.`);
        }
        a.command = a.command || arg;
    }
  }
  return a;
}

/** A summary for the GitHub Actions run page: numbers only, nothing that identifies people or sites. */
export function stepSummary(r: RunResult): string {
  const head = r.status === 'done' ? `Access report ${r.period}: done` : r.status === 'skipped' ? `Access report ${r.period}: already exists, nothing to do` : `Access report ${r.period}: not finished yet`;
  const lines = [`### ${head}`, ''];
  if (r.status !== 'skipped') {
    lines.push('| | |', '| --- | ---: |');
    lines.push(`| Site collections read | ${r.sitesDone} of ${r.sitesTotal} |`);
    lines.push(`| Site collections not read | ${r.sitesFailed} |`);
    lines.push(`| People and groups | ${r.principals}${r.unresolved ? ` (+${r.unresolved} not found)` : ''} |`);
    if (r.status === 'done') {
      lines.push(`| Places with access | ${r.rows} |`);
      lines.push(`| Changes since ${r.comparedWith || 'the previous report'} | ${r.comparedWith ? `${r.added} added, ${r.removed} removed, ${r.changed} changed` : 'first report'} |`);
    }
    lines.push(`| Requests (throttled) | ${r.requests} (${r.throttled}) |`);
    if (r.status === 'partial') {
      lines.push('', 'The time limit was reached. Progress is saved in the report folder and the next scheduled run carries on from there.');
    }
    if (r.warnings.length) {
      lines.push('', `${r.warnings.length} warning${r.warnings.length === 1 ? '' : 's'}: see the Summary sheet of the report.`);
    }
  }
  return lines.join('\n') + '\n';
}

export async function main(argv: string[], env: NodeJS.ProcessEnv): Promise<number> {
  let args: Args;
  try {
    args = parseArgs(argv);
  } catch (e) {
    process.stderr.write(`${(e as Error).message}\n`);
    return 2;
  }
  if (args.command === 'version') {
    process.stdout.write(`${VERSION}\n`);
    return 0;
  }
  if (!args.command || args.command === 'help') {
    process.stdout.write(HELP);
    return args.command ? 0 : 2;
  }
  if (args.command !== 'run' && args.command !== 'check') {
    process.stderr.write(`Unknown command "${args.command}". Try --help.\n`);
    return 2;
  }
  const log = createLogger(args.verbose);
  try {
    const cfg = await loadConfig(args.config, env);
    const clients = createClients(cfg, log, env);
    if (args.command === 'check') {
      return (await check(cfg, clients, log)) ? 0 : 1;
    }
    const result = await run(cfg, clients, log, { period: args.period, force: args.force, upload: args.upload, outDir: args.out, limit: args.limit });
    if (args.json) {
      // File addresses identify the tenant: only with --verbose.
      process.stdout.write(JSON.stringify(args.verbose ? result : { ...result, files: result.files.length, warnings: result.warnings.length }, null, 2) + '\n');
    }
    if (env.GITHUB_STEP_SUMMARY) {
      await appendFile(env.GITHUB_STEP_SUMMARY, stepSummary(result)).catch(() => undefined);
    }
    return result.status === 'partial' ? EXIT_PARTIAL : 0;
  } catch (e) {
    if (e instanceof ConfigError) {
      log.error(e.message);
      return 2;
    }
    log.error((e as Error).message || String(e));
    if (args.verbose && (e as Error).stack) {
      process.stderr.write(String((e as Error).stack) + '\n');
    }
    return 1;
  }
}

const invoked = process.argv[1] ? /cli\.js$/.test(process.argv[1]) || /access-report$/.test(process.argv[1]) : false;
if (invoked) {
  main(process.argv.slice(2), process.env).then(
    (code) => {
      process.exitCode = code;
    },
    (e) => {
      process.stderr.write(`${(e as Error).message || e}\n`);
      process.exitCode = 1;
    }
  );
}
