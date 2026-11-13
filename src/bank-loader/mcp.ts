import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { Client } from "@modelcontextprotocol/client";
import { getDefaultEnvironment, StdioClientTransport } from "@modelcontextprotocol/client/stdio";
import { serverEnvironment } from "./config.js";
import type { LoaderConfig } from "./types.js";

function loaderVersion(): string {
  try {
    const pkg = JSON.parse(readFileSync(new URL("../../package.json", import.meta.url), "utf8")) as {
      version?: string;
    };
    return pkg.version ?? "0.0.0";
  } catch {
    return "0.0.0";
  }
}

export interface Connected {
  client: Client;
  close: () => Promise<void>;
}

/**
 * Поднимает штатный stdio MCP-сервер загрузчика. Оператор задаёт адрес и реквизиты локально;
 * секреты уходят дочернему процессу через окружение, а не через командную строку.
 */
export async function connectLoader(cfg: LoaderConfig): Promise<Connected> {
  const serverPath = fileURLToPath(new URL("../index.js", import.meta.url));
  const transport = new StdioClientTransport({
    command: process.execPath,
    args: [serverPath],
    env: { ...getDefaultEnvironment(), ...serverEnvironment(cfg) },
    stderr: "inherit",
  });
  const client = new Client({ name: "1c-bp-bank-loader", version: loaderVersion() });
  await client.connect(transport);
  return { client, close: () => client.close() };
}
