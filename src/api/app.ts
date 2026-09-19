import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import Fastify, { type FastifyInstance } from 'fastify';
import cors from '@fastify/cors';
import multipart from '@fastify/multipart';
import { BitStore } from '../store/store.js';
import { ensureDataDir, saveConfig } from '../agent/config.js';
import {
  ProjectService,
  NotFoundError,
  ValidationError,
  ConflictError,
} from '../service/projects.js';

const execFileAsync = promisify(execFile);

export interface BuildAppOptions {
  dataDir: string;
  logger?: boolean;
}

async function readUpload(
  request: { file: () => Promise<any> },
): Promise<{ fields: Record<string, string>; fileBuffer: Buffer | null; filename: string | null }> {
  const fields: Record<string, string> = {};
  let fileBuffer: Buffer | null = null;
  let filename: string | null = null;

  const parts = (request as any).parts
    ? (request as any).parts()
    : null;

  if (parts) {
    for await (const part of parts) {
      if (part.type === 'file') {
        const chunks: Buffer[] = [];
        for await (const chunk of part.file) {
          chunks.push(chunk as Buffer);
        }
        fileBuffer = Buffer.concat(chunks);
        filename = part.filename ?? null;
      } else {
        fields[part.fieldname] = String(part.value ?? '');
      }
    }
  } else {
    // Fallback: single file() API + body fields if present
    const data = await request.file();
    if (data) {
      const chunks: Buffer[] = [];
      for await (const chunk of data.file) {
        chunks.push(chunk as Buffer);
      }
      fileBuffer = Buffer.concat(chunks);
      filename = data.filename ?? null;
      for (const [key, val] of Object.entries(data.fields ?? {})) {
        const v = val as { value?: unknown };
        fields[key] = String(v?.value ?? '');
      }
    }
  }

  return { fields, fileBuffer, filename };
}

function authorFrom(request: { headers: Record<string, unknown> }, fields: Record<string, string>): string {
  const header = request.headers['x-bit-author'];
  if (typeof header === 'string' && header.trim()) return header.trim();
  if (fields.author?.trim()) return fields.author.trim();
  return 'demo-user';
}

export async function buildApp(opts: BuildAppOptions): Promise<{
  app: FastifyInstance;
  store: BitStore;
  service: ProjectService;
}> {
  const runtime = {
    dataDir: opts.dataDir,
    store: new BitStore(opts.dataDir),
    service: null as unknown as ProjectService,
  };
  await runtime.store.init();
  runtime.service = new ProjectService(runtime.store);

  async function switchDataDir(next: string): Promise<string> {
    await ensureDataDir(next);
    const store = new BitStore(next);
    await store.init();
    runtime.dataDir = next;
    runtime.store = store;
    runtime.service = new ProjectService(store);
    await saveConfig({ dataDir: next });
    return next;
  }

  const app = Fastify({ logger: opts.logger ?? false });
  await app.register(cors, { origin: true });
  await app.register(multipart, {
    limits: { fileSize: 50 * 1024 * 1024 },
  });

  app.setErrorHandler((err, _req, reply) => {
    if (err instanceof NotFoundError) {
      return reply.status(404).send({ error: err.message });
    }
    if (err instanceof ValidationError) {
      return reply.status(400).send({ error: err.message });
    }
    if (err instanceof ConflictError) {
      return reply.status(409).send({ error: err.message, details: err.details });
    }
    app.log.error(err);
    return reply.status(500).send({ error: 'Internal server error' });
  });

  app.get('/health', async () => ({ ok: true }));

  app.get('/agent/status', async () => {
    const projects = await runtime.service.listProjects();
    return {
      ok: true,
      dataDir: runtime.dataDir,
      projectCount: projects.length,
      addin: true,
    };
  });

  app.put<{ Body: { dataDir?: string } }>('/agent/data-dir', async (request) => {
    const body = (request.body ?? {}) as { dataDir?: string };
    const next = body.dataDir?.trim();
    if (!next) throw new ValidationError('dataDir is required');
    const dataDir = await switchDataDir(next);
    return { ok: true, dataDir };
  });

  app.post('/agent/pick-folder', async () => {
    if (process.platform !== 'win32') {
      throw new ValidationError('Folder picker is Windows-only — paste a path instead');
    }
    const script = [
      'Add-Type -AssemblyName System.Windows.Forms',
      '$d = New-Object System.Windows.Forms.FolderBrowserDialog',
      "$d.Description = 'Choose the Bit storage folder'",
      '$d.ShowNewFolderButton = $true',
      '$r = $d.ShowDialog()',
      'if ($r -eq [System.Windows.Forms.DialogResult]::OK) { Write-Output $d.SelectedPath }',
    ].join('; ');
    try {
      const { stdout } = await execFileAsync(
        'powershell.exe',
        ['-NoProfile', '-STA', '-Command', script],
        { timeout: 120000, windowsHide: false },
      );
      const picked = stdout.trim();
      if (!picked) throw new ValidationError('No folder selected');
      const dataDir = await switchDataDir(picked);
      return { ok: true, dataDir };
    } catch (err) {
      if (err instanceof ValidationError) throw err;
      throw new ValidationError(
        `Couldn't open folder picker (${err instanceof Error ? err.message : String(err)})`,
      );
    }
  });

  app.get('/projects', async () => {
    return runtime.service.listProjects();
  });

  app.post('/projects', async (request, reply) => {
    const { fields, fileBuffer, filename } = await readUpload(request as any);
    if (!fileBuffer || !filename) {
      throw new ValidationError('Upload an .xlsx file to create a project');
    }
    if (!filename.toLowerCase().endsWith('.xlsx')) {
      throw new ValidationError('File must be .xlsx');
    }
    const name =
      fields.name?.trim() ||
      filename.replace(/\.xlsx$/i, '') ||
      'Untitled project';
    const detail = await runtime.service.createProject({
      name,
      author: authorFrom(request as any, fields),
      message: fields.message,
      xlsxBuffer: fileBuffer,
    });
    return reply.status(201).send(detail);
  });

  app.get<{ Params: { id: string } }>('/projects/:id', async (request) => {
    return runtime.service.getProject(request.params.id);
  });

  app.get<{ Params: { id: string } }>('/projects/:id/scenarios', async (request) => {
    return runtime.service.listScenarios(request.params.id);
  });

  app.post<{ Params: { id: string }; Body: { name?: string; author?: string } }>(
    '/projects/:id/scenarios',
    async (request, reply) => {
      const body = (request.body ?? {}) as { name?: string; author?: string };
      const name = typeof body.name === 'string' ? body.name : '';
      const scenario = await runtime.service.createScenario({
        projectId: request.params.id,
        name,
        author: authorFrom(request as any, {
          author: typeof body.author === 'string' ? body.author : '',
        }),
      });
      return reply.status(201).send(scenario);
    },
  );

  app.get<{ Params: { id: string } }>('/scenarios/:id', async (request) => {
    return runtime.service.getScenario(request.params.id);
  });

  app.get<{ Params: { id: string } }>('/projects/:id/versions', async (request) => {
    return runtime.service.listVersions(request.params.id);
  });

  app.post<{ Params: { id: string } }>('/projects/:id/versions', async (request, reply) => {
    const { fields, fileBuffer, filename } = await readUpload(request as any);
    if (!fileBuffer || !filename) {
      throw new ValidationError('Upload an .xlsx file to save a version');
    }
    if (!filename.toLowerCase().endsWith('.xlsx')) {
      throw new ValidationError('File must be .xlsx');
    }
    const version = await runtime.service.saveVersion({
      projectId: request.params.id,
      author: authorFrom(request as any, fields),
      message: fields.message || 'Saved version',
      xlsxBuffer: fileBuffer,
    });
    return reply.status(201).send(version);
  });

  app.post<{ Params: { id: string } }>('/scenarios/:id/versions', async (request, reply) => {
    const { fields, fileBuffer, filename } = await readUpload(request as any);
    if (!fileBuffer || !filename) {
      throw new ValidationError('Upload an .xlsx file to save a version');
    }
    if (!filename.toLowerCase().endsWith('.xlsx')) {
      throw new ValidationError('File must be .xlsx');
    }
    // Resolve scenario → project
    const meta = await runtime.store.readMeta();
    const scenario = meta.scenarios.find((s) => s.id === request.params.id);
    if (!scenario) throw new NotFoundError(`Scenario not found: ${request.params.id}`);
    const version = await runtime.service.saveVersion({
      projectId: scenario.projectId,
      scenarioId: scenario.id,
      author: authorFrom(request as any, fields),
      message: fields.message || 'Saved version',
      xlsxBuffer: fileBuffer,
    });
    return reply.status(201).send(version);
  });

  app.get<{ Params: { id: string } }>('/versions/:id', async (request) => {
    return runtime.service.getVersion(request.params.id);
  });

  app.get<{ Params: { id: string } }>('/versions/:id/xlsx', async (request, reply) => {
    const { buffer, filename } = await runtime.service.getVersionXlsx(request.params.id);
    return reply
      .header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
      .header('Content-Disposition', `attachment; filename="${filename}"`)
      .send(buffer);
  });


  app.get<{ Querystring: { base?: string; compare?: string } }>('/diff', async (request) => {
    const base = request.query.base;
    const compare = request.query.compare;
    if (!base || !compare) {
      throw new ValidationError('Query params base and compare are required');
    }
    return runtime.service.diffVersions(base, compare);
  });


  app.get<{ Querystring: { base?: string; ours?: string; theirs?: string } }>('/merge', async (request) => {
    const { base, ours, theirs } = request.query;
    if (!base || !ours || !theirs) {
      throw new ValidationError('Query params base, ours, and theirs are required');
    }
    return runtime.service.previewMerge({ baseId: base, oursId: ours, theirsId: theirs });
  });


  app.post<{ Params: { id: string }; Body: { note?: string; author?: string } }>(
    '/scenarios/:id/reviews',
    async (request, reply) => {
      const body = (request.body ?? {}) as { note?: string; author?: string };
      const review = await runtime.service.createReview({
        scenarioId: request.params.id,
        author: authorFrom(request as any, {
          author: typeof body.author === 'string' ? body.author : '',
        }),
        note: body.note,
      });
      return reply.status(201).send(review);
    },
  );

  app.get<{ Params: { id: string } }>('/projects/:id/reviews', async (request) => {
    return runtime.service.listReviews(request.params.id);
  });

  app.get<{ Params: { id: string } }>('/reviews/:id', async (request) => {
    return runtime.service.getReview(request.params.id);
  });

  app.post<{ Params: { id: string }; Body: Record<string, unknown> }>(
    '/reviews/:id/combine',
    async (request, reply) => {
      const body = (request.body ?? {}) as {
        author?: string;
        message?: string;
        resolutions?: Record<string, { action: string; cell?: unknown }>;
      };
      const result = await runtime.service.combineReview({
        reviewId: request.params.id,
        author: authorFrom(request as any, {
          author: typeof body.author === 'string' ? body.author : '',
        }),
        message: body.message,
        resolutions: body.resolutions as any,
      });
      return reply.status(200).send(result);
    },
  );

  app.post<{ Params: { id: string }; Body: { note?: string } }>(
    '/reviews/:id/request-changes',
    async (request) => {
      const body = (request.body ?? {}) as { note?: string };
      return runtime.service.updateReviewStatus(request.params.id, 'changes-requested', body.note);
    },
  );

  app.post<{ Params: { id: string }; Body: { note?: string } }>(
    '/reviews/:id/close',
    async (request) => {
      const body = (request.body ?? {}) as { note?: string };
      return runtime.service.updateReviewStatus(request.params.id, 'closed', body.note);
    },
  );

  app.get<{ Params: { id: string } }>('/projects/:id/graph', async (request) => {
    return runtime.service.getProjectGraph(request.params.id);
  });

  app.get<{ Params: { id: string } }>('/versions/:id/xlsx/base64', async (request) => {
    const { buffer, filename } = await runtime.service.getVersionXlsx(request.params.id);
    return {
      filename,
      mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      base64: buffer.toString('base64'),
    };
  });

  return { app, store: runtime.store, service: runtime.service };
}
