import { createInsertSchema } from "drizzle-zod";
import { jsonb, pgTable, text, timestamp } from "drizzle-orm/pg-core";
import { z } from "zod/v4";

export type StoredQuestion = {
  id: string;
  text: string;
  sourceRow?: string[];
};

export type StoredAnswer = {
  questionId: string;
  questionText: string;
  answer: string;
  sourceRow?: string[];
};

export const formsTable = pgTable("randomized_forms", {
  id: text("id").primaryKey(),
  title: text("title").notNull(),
  headers: jsonb("headers").$type<string[]>().notNull(),
  questions: jsonb("questions").$type<StoredQuestion[]>().notNull(),
  selectionHistory: jsonb("selection_history")
    .$type<Record<string, number>>()
    .notNull()
    .default({}),
  createdAt: timestamp("created_at", { withTimezone: true })
    .notNull()
    .defaultNow(),
  updatedAt: timestamp("updated_at", { withTimezone: true })
    .notNull()
    .defaultNow()
    .$onUpdate(() => new Date()),
});

export const responsesTable = pgTable("randomized_responses", {
  id: text("id").primaryKey(),
  formId: text("form_id")
    .notNull()
    .references(() => formsTable.id, { onDelete: "cascade" }),
  respondentRole: text("respondent_role").notNull(),
  answers: jsonb("answers").$type<StoredAnswer[]>().notNull(),
  createdAt: timestamp("created_at", { withTimezone: true })
    .notNull()
    .defaultNow(),
});

export const insertFormSchema = createInsertSchema(formsTable).omit({
  createdAt: true,
  updatedAt: true,
});
export const insertResponseSchema = createInsertSchema(responsesTable).omit({
  createdAt: true,
});

export type InsertForm = z.infer<typeof insertFormSchema>;
export type Form = typeof formsTable.$inferSelect;
export type InsertResponse = z.infer<typeof insertResponseSchema>;
export type Response = typeof responsesTable.$inferSelect;