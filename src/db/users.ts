import { db, isSqlConfigured } from './index.ts';
import { users } from './schema.ts';

const memoryUsers = new Map<string, { id: number; uid: string; email: string; createdAt: Date }>();
let nextUserId = 1;

export async function getOrCreateUser(uid: string, email: string) {
  if (!isSqlConfigured) {
    const existing = memoryUsers.get(uid);
    if (existing) {
      existing.email = email;
      return existing;
    }
    const created = { id: nextUserId++, uid, email, createdAt: new Date() };
    memoryUsers.set(uid, created);
    return created;
  }

  try {
    const result = await db
      .insert(users)
      .values({
        uid,
        email,
      })
      .onConflictDoUpdate({
        target: users.uid,
        set: {
          email,
        },
      })
      .returning();

    return result[0];
  } catch (error) {
    console.error('Failed to get or create user:', error);
    throw new Error('Failed to synchronize user account.', { cause: error });
  }
}
