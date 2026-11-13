import { readFileSync } from "node:fs";
import { homedir } from "node:os";
import { resolve } from "node:path";
import { LoaderConfigSchema, SourcePackageSchema, type LoaderConfig, type SourceDocument } from "./types.js";

/** ~ в локальных путях: конфигурация оператора не должна зависеть от HOME в кавычках. */
export function expandHome(path: string): string {
  if (path === "~") return homedir();
  if (path.startsWith("~/")) return resolve(homedir(), path.slice(2));
  return path;
}

export function readLoaderConfig(path: string): LoaderConfig {
  let raw: unknown;
  try {
    raw = JSON.parse(readFileSync(path, "utf8"));
  } catch (error) {
    throw new Error(
      `Не прочитана локальная конфигурация ${path}: ${error instanceof Error ? error.message : String(error)}`,
      { cause: error },
    );
  }
  const parsed = LoaderConfigSchema.safeParse(raw);
  if (!parsed.success)
    throw new Error(
      `Локальная конфигурация ${path} неверна: ${parsed.error.issues
        .map((i) => `${i.path.join(".") || "config"}: ${i.message}`)
        .join("; ")}`,
    );
  const cfg = parsed.data;
  return {
    ...cfg,
    target: { ...cfg.target },
    odata: {
      ...cfg.odata,
      baseUrl: cfg.odata.baseUrl.endsWith("/") ? cfg.odata.baseUrl : `${cfg.odata.baseUrl}/`,
    },
    journalDir: expandHome(cfg.journalDir),
  };
}

/** Учётные данные: явные в локальном файле или ссылки на переменные окружения. */
export function resolveODataCredentials(cfg: LoaderConfig): { username: string; password: string } {
  const username =
    cfg.odata.username ?? (cfg.odata.usernameEnv ? process.env[cfg.odata.usernameEnv] : undefined);
  const password =
    cfg.odata.password ?? (cfg.odata.passwordEnv ? process.env[cfg.odata.passwordEnv] : undefined);
  if (!username)
    throw new Error(
      `Не задан логин OData: укажите odata.username либо odata.usernameEnv и переменную ${cfg.odata.usernameEnv ?? "..."}.`,
    );
  if (!password)
    throw new Error(
      `Не задан пароль OData: укажите odata.password либо odata.passwordEnv и переменную ${cfg.odata.passwordEnv ?? "..."}.`,
    );
  return { username, password };
}

/** Переменные процесса MCP-сервера; секреты не попадают в аргументы командной строки. */
export function serverEnvironment(cfg: LoaderConfig): Record<string, string> {
  const credentials = resolveODataCredentials(cfg);
  const env: Record<string, string> = {
    ODATA_BASE_URL: cfg.odata.baseUrl,
    ODATA_USERNAME: credentials.username,
    ODATA_PASSWORD: credentials.password,
    // Загрузчик выполняет предпросмотр и подтверждение: сервер должен видеть инструменты записи.
    READ_ONLY: "false",
    ODATA_WRITABLE: "true",
    TOOLS_PROFILE: "all",
    MCP_TRANSPORT: "stdio",
  };
  if (cfg.odata.writeJournalDir) env.ODATA_WRITE_JOURNAL_DIR = expandHome(cfg.odata.writeJournalDir);
  if (process.env.LOG_LEVEL) env.LOG_LEVEL = process.env.LOG_LEVEL;
  return env;
}

export function readSourcePackage(path: string): SourceDocument[] {
  let raw: unknown;
  try {
    raw = JSON.parse(readFileSync(path, "utf8"));
  } catch (error) {
    throw new Error(
      `Не прочитан пакет выписки ${path}: ${error instanceof Error ? error.message : String(error)}`,
      { cause: error },
    );
  }
  const parsed = SourcePackageSchema.safeParse(raw);
  if (!parsed.success)
    throw new Error(
      `Пакет ${path} неверен: ${parsed.error.issues
        .map((i) => `${i.path.join(".") || "document"}: ${i.message}`)
        .join("; ")}`,
    );
  return Array.isArray(parsed.data) ? parsed.data : parsed.data.documents;
}
