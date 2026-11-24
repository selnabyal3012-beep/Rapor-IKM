import { desc, eq } from 'drizzle-orm';
import { db, isSqlConfigured } from './index.ts';
import { raporBackups, raporStates, users } from './schema.ts';

export interface SaveRaporStateInput {
  userId: number;
  schoolKey: string;
  npsn: string;
  namaSekolah: string;
  tahunPelajaran: string;
  payloadJson: string;
  updatedBy: string;
}

interface MemoryRaporState {
  id: number;
  userId: number;
  schoolKey: string;
  npsn: string;
  namaSekolah: string;
  tahunPelajaran: string;
  payloadJson: string;
  updatedBy: string;
  updatedAt: Date;
}

interface MemoryBackup {
  id: number;
  userId: number;
  backupKey: string;
  schoolName: string;
  payloadJson: string;
  createdBy: string;
  createdAt: Date;
}

const memoryStates = new Map<string, MemoryRaporState>();
const memoryBackups: MemoryBackup[] = [];
let nextStateId = 1;
let nextBackupId = 1;

export async function saveSchoolStateToSql(input: SaveRaporStateInput) {
  if (!isSqlConfigured) {
    const now = new Date();
    const existing = memoryStates.get(input.schoolKey);
    const saved: MemoryRaporState = {
      id: existing?.id ?? nextStateId++,
      userId: input.userId,
      schoolKey: input.schoolKey,
      npsn: input.npsn,
      namaSekolah: input.namaSekolah,
      tahunPelajaran: input.tahunPelajaran,
      payloadJson: input.payloadJson,
      updatedBy: input.updatedBy,
      updatedAt: now,
    };
    memoryStates.set(input.schoolKey, saved);
    memoryBackups.unshift({
      id: nextBackupId++,
      userId: input.userId,
      backupKey: `backup_${now.toISOString().replace(/[:.]/g, '-')}`,
      schoolName: input.namaSekolah,
      payloadJson: input.payloadJson,
      createdBy: input.updatedBy,
      createdAt: now,
    });
    if (memoryBackups.length > 25) memoryBackups.length = 25;
    return saved;
  }

  try {
    const now = new Date();
    const [saved] = await db
      .insert(raporStates)
      .values({
        userId: input.userId,
        schoolKey: input.schoolKey,
        npsn: input.npsn,
        namaSekolah: input.namaSekolah,
        tahunPelajaran: input.tahunPelajaran,
        payloadJson: input.payloadJson,
        updatedBy: input.updatedBy,
        updatedAt: now,
      })
      .onConflictDoUpdate({
        target: raporStates.schoolKey,
        set: {
          userId: input.userId,
          npsn: input.npsn,
          namaSekolah: input.namaSekolah,
          tahunPelajaran: input.tahunPelajaran,
          payloadJson: input.payloadJson,
          updatedBy: input.updatedBy,
          updatedAt: now,
        },
      })
      .returning();

    const backupKey = `backup_${now.toISOString().replace(/[:.]/g, '-')}`;
    await db.insert(raporBackups).values({
      userId: input.userId,
      backupKey,
      schoolName: input.namaSekolah,
      payloadJson: input.payloadJson,
      createdBy: input.updatedBy,
      createdAt: now,
    });

    return saved;
  } catch (error) {
    console.error('Database save failed:', error);
    throw new Error('Database save failed. Please try again later.', { cause: error });
  }
}

export async function getSchoolStateFromSql(uid: string, schoolKey: string) {
  if (!isSqlConfigured) {
    void uid;
    return memoryStates.get(schoolKey) ?? null;
  }

  try {
    const rows = await db
      .select({
        id: raporStates.id,
        schoolKey: raporStates.schoolKey,
        npsn: raporStates.npsn,
        namaSekolah: raporStates.namaSekolah,
        tahunPelajaran: raporStates.tahunPelajaran,
        payloadJson: raporStates.payloadJson,
        updatedBy: raporStates.updatedBy,
        updatedAt: raporStates.updatedAt,
      })
      .from(raporStates)
      .innerJoin(users, eq(raporStates.userId, users.id))
      .where(eq(raporStates.schoolKey, schoolKey))
      .limit(1);

    if (!rows.length) {
      return null;
    }

    // Touch user UID reference to ensure audit context
    void uid;
    return rows[0];
  } catch (error) {
    console.error('Database query failed:', error);
    throw new Error('Database query failed. Please try again later.', { cause: error });
  }
}

export async function listSqlBackups(uid: string) {
  if (!isSqlConfigured) {
    void uid;
    return memoryBackups.map(({ id, backupKey, schoolName, createdBy, createdAt }) => ({
      id,
      backupKey,
      schoolName,
      createdBy,
      createdAt,
    }));
  }

  try {
    return await db
      .select({
        id: raporBackups.id,
        backupKey: raporBackups.backupKey,
        schoolName: raporBackups.schoolName,
        createdBy: raporBackups.createdBy,
        createdAt: raporBackups.createdAt,
      })
      .from(raporBackups)
      .innerJoin(users, eq(raporBackups.userId, users.id))
      .where(eq(users.uid, uid))
      .orderBy(desc(raporBackups.createdAt))
      .limit(25);
  } catch (error) {
    console.error('Database query failed:', error);
    throw new Error('Database query failed. Please try again later.', { cause: error });
  }
}
