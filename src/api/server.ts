import { buildApp } from './app.js';
import { loadConfig } from '../agent/config.js';

const PORT = Number(process.env.PORT ?? 3001);
const HOST = process.env.HOST ?? '127.0.0.1';
const cfg = await loadConfig();

const { app } = await buildApp({ dataDir: cfg.dataDir, logger: true });

await app.listen({ port: PORT, host: HOST });
console.log(`Bit API listening on http://${HOST}:${PORT}`);
console.log(`Data dir: ${cfg.dataDir}`);
console.log('Auth: stub — send X-Bit-Author header or form field "author" (defaults to demo-user)');
