#!/usr/bin/env node
import { BankLoadJournal } from "../bank-loader/journal.js";
import { readLoaderConfig, readSourcePackage } from "../bank-loader/config.js";
import { connectLoader } from "../bank-loader/mcp.js";
import {
  diagnose,
  ensureAccount,
  reportHasProblems,
  runPackage,
  type ToolClient,
} from "../bank-loader/runner.js";

/**
 * Операторский CLI первого участка загрузчика БП. Работает через stdio MCP-сервер, поэтому
 * подготовка, предпросмотр, operationId и журнал записи — те же, что у MCP-клиента.
 *
 *   bp-loader diagnose --config bank-loader.json
 *   bp-loader ensure-account --config bank-loader.json [--confirm]
 *   bp-loader prepare --config bank-loader.json --source package.json
 *   bp-loader load --config bank-loader.json --source package.json [--confirm]
 *   bp-loader resume --config bank-loader.json --source package.json
 *   bp-loader status --config bank-loader.json
 */

interface Options {
  command: string;
  config: string;
  source?: string;
  confirm: boolean;
  database?: string;
}

const COMMANDS = new Set(["diagnose", "ensure-account", "prepare", "load", "resume", "status"]);
const USAGE =
  "Команды: diagnose, ensure-account, prepare, load, resume, status. " +
  "Параметры: --config <файл> (по умолчанию bank-loader.json), --source <файл>, --confirm, --database <имя>.";

function parseArgs(argv: readonly string[]): Options {
  const command = argv[0];
  if (!command || !COMMANDS.has(command)) throw new Error(`Неизвестная команда. ${USAGE}`);
  const options: Options = { command, config: "bank-loader.json", confirm: false };
  for (let i = 1; i < argv.length; i++) {
    const arg = argv[i]!;
    if (arg === "--confirm") options.confirm = true;
    else if (arg === "--config" || arg === "-c") {
      const value = argv[++i];
      if (!value) throw new Error("--config требует путь к файлу.");
      options.config = value;
    } else if (arg === "--source" || arg === "-s") {
      const value = argv[++i];
      if (!value) throw new Error("--source требует путь к файлу пакета.");
      options.source = value;
    } else if (arg === "--database") {
      const value = argv[++i];
      if (!value) throw new Error("--database требует имя базы.");
      options.database = value;
    } else throw new Error(`Неизвестный параметр ${arg}. ${USAGE}`);
  }
  if (["prepare", "load", "resume"].includes(command) && !options.source)
    throw new Error(`${command} требует --source <файл> с пакетом строк выписки.`);
  return options;
}

const print = (value: unknown): void => {
  process.stdout.write(`${JSON.stringify(value, null, 2)}\n`);
};

async function main(): Promise<void> {
  const options = parseArgs(process.argv.slice(2));
  const config = readLoaderConfig(options.config);
  const journal = new BankLoadJournal(config.journalDir);

  if (options.command === "status") {
    print(await journal.summary());
    return;
  }

  const sources =
    options.command === "diagnose" || options.command === "ensure-account"
      ? []
      : readSourcePackage(options.source!);
  const connected = await connectLoader(config);
  try {
    const client: ToolClient = connected.client;
    if (options.command === "diagnose") {
      const report = await diagnose(client, config.target, options.database);
      print(report);
      if (!report.ok) process.exitCode = 1;
      return;
    }
    if (options.command === "ensure-account") {
      const report = await ensureAccount(client, config.target, {
        confirm: options.confirm,
        ...(options.database ? { database: options.database } : {}),
      });
      print(report);
      if (report.status === "blocked" || report.status === "refused") process.exitCode = 1;
      return;
    }
    const report = await runPackage(client, config.target, sources, journal, {
      confirm: options.confirm,
      mode:
        options.command === "resume"
          ? "resume"
          : options.command === "prepare"
            ? "prepare"
            : options.confirm
              ? "confirm"
              : "prepare",
      ...(options.database ? { database: options.database } : {}),
    });
    print(report);
    if (reportHasProblems(report)) process.exitCode = 1;
  } finally {
    await connected.close().catch(() => undefined);
  }
}

main().catch((error) => {
  process.stderr.write(
    `${JSON.stringify({ error: error instanceof Error ? error.message : String(error) }, null, 2)}\n`,
  );
  process.exitCode = 2;
});
