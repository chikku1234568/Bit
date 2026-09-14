import { join } from 'node:path';
import { buildApp } from './app.js';

const PORT = Number(process.env.PORT ?? 3001);
const HOST = process.env.HOST ?? '127.0.0.1';
const DATA_DIR = process.env.BIT_DATA_DIR ?? join(process.cwd(), 'data');

const { app } = await buildApp({ dataDir: DATA_DIR, logger: true });

await app.listen({ port: PORT, host: HOST });
console.log(`Bit API listening on http://${HOST}:${PORT}`);
console.log(`Data dir: ${DATA_DIR}`);
console.log('Auth: stub — send X-Bit-Author header or form field "author" (defaults to demo-user)');
