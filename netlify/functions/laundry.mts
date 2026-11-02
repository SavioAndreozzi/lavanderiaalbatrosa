import { and, eq, sql } from 'drizzle-orm';
import { db } from '../../db/index.js';
import { laundry } from '../../db/schema.js';

const reply = (body: unknown, status = 200) => Response.json(body, { status, headers: { 'Cache-Control': 'no-store' } });
function validRecords(records: unknown): records is Record<string, unknown>[] {
  if (!Array.isArray(records) || records.length > 500) return false;
  const ids = new Set();
  const codes = new Set();
  return records.every(r => {
    if (!r || typeof r !== 'object') return false;
    if (!['id', 'cod', 'cognome', 'nome'].every(k => typeof r[k] === 'string' && r[k].length > 0 && r[k].length <= 100)) return false;
    if (!/^[a-zA-Z0-9_-]+$/.test(r.id) || ids.has(r.id) || codes.has(r.cod)) return false;
    ids.add(r.id); codes.add(r.cod);
    return ['pezzo', 'peso'].every(k => r[k] && typeof r[k] === 'object' && !Array.isArray(r[k]) && Object.values(r[k]).every(v => Number.isInteger(v) && Number(v) >= 0 && Number(v) <= 1000));
  });
}
export default async (req: Request) => {
  if (!['GET', 'PUT'].includes(req.method)) return reply({ error: 'Metodo non consentito.' }, 405);
  if (req.method === 'PUT' && req.headers.get('origin') && req.headers.get('origin') !== new URL(req.url).origin) return reply({ error: 'Origine non consentita.' }, 403);
  try {
    await db.insert(laundry).values({ id: 1 }).onConflictDoNothing();
    if (req.method === 'GET') {
      const [state] = await db.select().from(laundry).where(eq(laundry.id, 1));
      return reply(state);
    }
    const body = await req.json();
    if (!Number.isInteger(body.version) || !validRecords(body.records)) return reply({ error: 'Dati non validi.' }, 400);
    // Atomic comparison prevents an older browser from overwriting another person's changes.
    const [state] = await db.update(laundry).set({ records: body.records, version: sql`${laundry.version} + 1`, updatedAt: new Date() })
      .where(and(eq(laundry.id, 1), eq(laundry.version, body.version))).returning();
    if (!state) return reply({ error: 'Il grigliato è stato modificato da un’altra persona. Controlla i dati aggiornati e ripeti il salvataggio.' }, 409);
    return reply(state);
  } catch {
    return reply({ error: 'Servizio temporaneamente non disponibile. Riprova.' }, 503);
  }
};
export const config = { path: '/api/laundry' };
