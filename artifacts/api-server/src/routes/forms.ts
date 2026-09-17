import { Router, type IRouter } from "express";
import { randomUUID } from "node:crypto";
import { asc, eq } from "drizzle-orm";
import { db, formsTable, responsesTable } from "@workspace/db";
import {
  GetPublicFormParams,
  GetPublicFormResponse,
  SubmitFormResponseBody,
  SubmitFormResponseParams,
  SubmitFormResponseResponse,
} from "@workspace/api-zod";

const router: IRouter = Router();

function formIdFromParams(params: Record<string, string | string[]>) {
  const value = params.formId;
  return Array.isArray(value) ? value[0] : value;
}

function balancedSelection<T extends { id: string }>(
  questions: T[],
  history: Record<string, number>,
) {
  return [...questions]
    .sort(
      (left, right) =>
        (history[left.id] ?? 0) - (history[right.id] ?? 0) ||
        Math.random() - 0.5,
    )
    .slice(0, 5);
}

router.get("/forms/:formId", async (req, res): Promise<void> => {
  const params = GetPublicFormParams.safeParse({
    formId: formIdFromParams(req.params),
  });
  if (!params.success) {
    res.status(400).json({ error: params.error.message });
    return;
  }

  const [form] = await db
    .select()
    .from(formsTable)
    .where(eq(formsTable.id, params.data.formId))
    .limit(1);
  if (!form) {
    res.status(404).json({ error: "Form not found" });
    return;
  }

  const selected = balancedSelection(form.questions, form.selectionHistory);
  const nextHistory = { ...form.selectionHistory };
  for (const question of selected) {
    nextHistory[question.id] = (nextHistory[question.id] ?? 0) + 1;
  }
  await db
    .update(formsTable)
    .set({ selectionHistory: nextHistory, updatedAt: new Date() })
    .where(eq(formsTable.id, form.id));

  res.json(
    GetPublicFormResponse.parse({
      id: form.id,
      title: form.title,
      headers: form.headers,
      questions: selected,
    }),
  );
});

router.post("/forms/:formId", async (req, res): Promise<void> => {
  const params = SubmitFormResponseParams.safeParse({
    formId: formIdFromParams(req.params),
  });
  const body = SubmitFormResponseBody.safeParse(req.body);
  if (!params.success || !body.success) {
    res.status(400).json({
      error: !params.success
        ? params.error.message
        : body.success
          ? "Invalid request"
          : body.error.message,
    });
    return;
  }

  const [form] = await db
    .select({ id: formsTable.id })
    .from(formsTable)
    .where(eq(formsTable.id, params.data.formId))
    .limit(1);
  if (!form) {
    res.status(404).json({ error: "Form not found" });
    return;
  }

  const [response] = await db
    .insert(responsesTable)
    .values({
      id: randomUUID(),
      formId: form.id,
      respondentRole: body.data.respondentRole.trim(),
      answers: body.data.answers.map((answer) => ({
        ...answer,
        answer: answer.answer.trim(),
      })),
    })
    .returning();

  res.status(201).json(
    SubmitFormResponseResponse.parse({
      id: response.id,
      formId: response.formId,
      respondentRole: response.respondentRole,
      answers: response.answers,
      createdAt: response.createdAt,
    }),
  );
});

export default router;