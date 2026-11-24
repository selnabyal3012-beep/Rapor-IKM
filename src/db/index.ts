import { drizzle } from 'drizzle-orm/node-postgres';
import { Pool } from 'pg';
import * as schema from './schema.ts';

declare global {
  var _postgresPool: Pool | undefined;
}

export const isSqlConfigured = Boolean(
  process.env.SQL_HOST &&
    process.env.SQL_USER &&
    process.env.SQL_DB_NAME
);

export const createPool = () => {
  if (!isSqlConfigured) {
    return null;
  }
  if (!global._postgresPool) {
    global._postgresPool = new Pool({
      host: process.env.SQL_HOST,
      user: process.env.SQL_USER,
      password: process.env.SQL_PASSWORD,
      database: process.env.SQL_DB_NAME,
      max: 10,
      connectionTimeoutMillis: 15000,
    });

    global._postgresPool.on('error', (err) => {
      console.error('Unexpected error on idle SQL pool client:', err);
    });
  }
  return global._postgresPool;
};

const pool = createPool();

function createMockDb() {
  const noOp = {
    findMany: async () => [],
    findFirst: async () => null,
    findUnique: async () => null,
    create: async (d: any) => d?.data ?? {},
    update: async (d: any) => d?.data ?? {},
    delete: async () => ({}),
  };
  return new Proxy({} as ReturnType<typeof drizzle<typeof schema>>, {
    get: (_, prop) =>
      prop === 'query'
        ? new Proxy({}, { get: () => noOp })
        : async () => [],
  });
}

let dbInstance: ReturnType<typeof drizzle<typeof schema>>;
try {
  if (!pool) {
    console.warn('[AI Studio] SQL env vars not set — using in-memory database mock');
    dbInstance = createMockDb();
  } else {
    dbInstance = drizzle(pool, { schema });
  }
} catch {
  console.warn('[AI Studio] Database not connected — using mock');
  dbInstance = createMockDb();
}

export const db = dbInstance;
