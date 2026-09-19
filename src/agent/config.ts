import { homedir } from 'node:os';
import { join } from 'node:path';
import { mkdir, readFile, writeFile, access } from 'node:fs/promises';
import { constants as fsConstants } from 'node:fs';

export interface BitConfig {
  dataDir: string;
}

export function configFilePath(): string {
  return join(homedir(), '.bit', 'config.json');
}

export function defaultDataDir(): string {
  return join(process.cwd(), 'data');
}

export async function loadConfig(): Promise<BitConfig> {
  if (process.env.BIT_DATA_DIR?.trim()) {
    return { dataDir: process.env.BIT_DATA_DIR.trim() };
  }
  try {
    const raw = await readFile(configFilePath(), 'utf8');
    const parsed = JSON.parse(raw) as Partial<BitConfig>;
    if (typeof parsed.dataDir === 'string' && parsed.dataDir.trim()) {
      return { dataDir: parsed.dataDir.trim() };
    }
  } catch {
    /* missing config is fine */
  }
  return { dataDir: defaultDataDir() };
}

export async function saveConfig(cfg: BitConfig): Promise<void> {
  const dir = join(homedir(), '.bit');
  await mkdir(dir, { recursive: true });
  await writeFile(configFilePath(), JSON.stringify(cfg, null, 2), 'utf8');
}

export async function ensureDataDir(path: string): Promise<void> {
  const trimmed = path.trim();
  if (!trimmed) throw new Error('Folder path is required');
  await mkdir(trimmed, { recursive: true });
  await access(trimmed, fsConstants.W_OK);
}
