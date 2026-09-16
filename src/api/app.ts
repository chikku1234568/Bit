import Fastify, { type FastifyInstance } from 'fastify';
import cors from '@fastify/cors';
import multipart from '@fastify/multipart';
import { BitStore } from '../store/store.js';
import {
  ProjectService,
  NotFoundError,
  ValidationError,
  ConflictError,
} from '../service/projects.js';

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
  const store = new BitStore(opts.dataDir);
  await store.init();
  const service = new ProjectService(store);

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

  app.get('/projects', async () => {
    return service.listProjects();
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
    const detail = await service.createProject({
      name,
      author: authorFrom(request as any, fields),
      message: fields.message,
      xlsxBuffer: fileBuffer,
    });
    return reply.status(201).send(detail);
  });

  app.get<{ Params: { id: string } }>('/projects/:id', async (request) => {
    return service.getProject(request.params.id);
  });

  app.get<{ Params: { id: string } }>('/projects/:id/scenarios', async (request) => {
    return service.listScenarios(request.params.id);
  });

  app.post<{ Params: { id: string }; Body: { name?: string; author?: string } }>(
    '/projects/:id/scenarios',
    async (request, reply) => {
      const body = (request.body ?? {}) as { name?: string; author?: string };
      const name = typeof body.name === 'string' ? body.name : '';
      const scenario = await service.createScenario({
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
    return service.getScenario(request.params.id);
  });

  app.get<{ Params: { id: string } }>('/projects/:id/versions', async (request) => {
    return service.listVersions(request.params.id);
  });

  app.post<{ Params: { id: string } }>('/projects/:id/versions', async (request, reply) => {
    const { fields, fileBuffer, filename } = await readUpload(request as any);
    if (!fileBuffer || !filename) {
      throw new ValidationError('Upload an .xlsx file to save a version');
    }
    if (!filename.toLowerCase().endsWith('.xlsx')) {
      throw new ValidationError('File must be .xlsx');
    }
    const version = await service.saveVersion({
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
    const meta = await store.readMeta();
    const scenario = meta.scenarios.find((s) => s.id === request.params.id);
    if (!scenario) throw new NotFoundError(`Scenario not found: ${request.params.id}`);
    const version = await service.saveVersion({
      projectId: scenario.projectId,
      scenarioId: scenario.id,
      author: authorFrom(request as any, fields),
      message: fields.message || 'Saved version',
      xlsxBuffer: fileBuffer,
    });
    return reply.status(201).send(version);
  });

  app.get<{ Params: { id: string } }>('/versions/:id', async (request) => {
    return service.getVersion(request.params.id);
  });

  app.get<{ Params: { id: string } }>('/versions/:id/xlsx', async (request, reply) => {
    const { buffer, filename } = await service.getVersionXlsx(request.params.id);
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
    return service.diffVersions(base, compare);
  });


  app.get<{ Querystring: { base?: string; ours?: string; theirs?: string } }>('/merge', async (request) => {
    const { base, ours, theirs } = request.query;
    if (!base || !ours || !theirs) {
      throw new ValidationError('Query params base, ours, and theirs are required');
    }
    return service.previewMerge({ baseId: base, oursId: ours, theirsId: theirs });
  });


  app.post<{ Params: { id: string }; Body: { note?: string; author?: string } }>(
    '/scenarios/:id/reviews',
    async (request, reply) => {
      const body = (request.body ?? {}) as { note?: string; author?: string };
      const review = await service.createReview({
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
    return service.listReviews(request.params.id);
  });

  app.get<{ Params: { id: string } }>('/reviews/:id', async (request) => {
    return service.getReview(request.params.id);
  });

  app.post<{ Params: { id: string }; Body: Record<string, unknown> }>(
    '/reviews/:id/combine',
    async (request, reply) => {
      const body = (request.body ?? {}) as {
        author?: string;
        message?: string;
        resolutions?: Record<string, { action: string; cell?: unknown }>;
      };
      const result = await service.combineReview({
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
      return service.updateReviewStatus(request.params.id, 'changes-requested', body.note);
    },
  );

  app.post<{ Params: { id: string }; Body: { note?: string } }>(
    '/reviews/:id/close',
    async (request) => {
      const body = (request.body ?? {}) as { note?: string };
      return service.updateReviewStatus(request.params.id, 'closed', body.note);
    },
  );

  return { app, store, service };
}
