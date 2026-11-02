import { pgTable, integer, jsonb, timestamp } from 'drizzle-orm/pg-core';

export const laundry = pgTable('laundry', {
  id: integer().primaryKey(),
  records: jsonb().$type<Record<string, unknown>[]>().notNull().default([]),
  version: integer().notNull().default(0),
  updatedAt: timestamp('updated_at', { withTimezone: true }).notNull().defaultNow(),
});
