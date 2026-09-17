import { Router, type IRouter } from "express";
import { desc, eq } from "drizzle-orm";
import { db, formsTable, responsesTable } from "@workspace/db";
import {
  AdminLoginBody,
  AdminLoginResponse,
  GetAdminFormParams,
  GetAdminFormResponse,
  GetAdminStatusResponse,
  ListAdminResponsesParams,
  ListAdminResponsesResponse,
  SaveAdminFormBody,
  SaveAdminFormParams,
  SaveAdminFormResponse,
} from "@workspace/api-zod";
import {
  clearAdminSession,
  isAdminAuthenticated,
  requireAdmin,
  setAdminSession,
  verifyAdminPassword,
} from "../lib/admin-auth";

const router: IRouter = Router();

function formIdFromParams(params: Record<string, string | string[]>) {
  const value = params.formId;
  return Array.isArray(value) ? value[0] : value;
}

function adminFormResponse(form: typeof formsTable.$inferSelect) {
  return {
    id: form.id,
    title: form.title,
    headers: form.headers,
    questions: form.questions,
    updatedAt: form.updatedAt,
  };
}

router.post("/admin/login", (req, res): void => {
  const body = AdminLoginBody.safeParse(req.body);
  if (!body.success || !verifyAdminPassword(body.success ? body.data.password : "")) {
    res.status(401).json({ error: "Invalid admin password" });
    return;
  }
  setAdminSession(res);
  res.json(AdminLoginResponse.parse({ authenticated: true }));
});

router.post("/admin/logout", (_req, res): void => {
  clearAdminSession(res);
  res.sendStatus(204);
});

router.get("/admin/status", (req, res): void => {
  res.json(
    GetAdminStatusResponse.parse({
      authenticated: isAdminAuthenticated(req),
    }),
  );
});

router.get("/admin/forms/:formId", async (req, res): Promise<void> => {
  if (!requireAdmin(req, res)) return;
  const params = GetAdminFormParams.safeParse({
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
  res.json(GetAdminFormResponse.parse(adminFormResponse(form)));
});

router.put("/admin/forms/:formId", async (req, res): Promise<void> => {
  if (!requireAdmin(req, res)) return;
  const params = SaveAdminFormParams.safeParse({
    formId: formIdFromParams(req.params),
  });
  const body = SaveAdminFormBody.safeParse(req.body);
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
    .insert(formsTable)
    .values({
      id: params.data.formId,
      title: body.data.title.trim(),
      headers: body.data.headers,
      questions: body.data.questions,
      selectionHistory: {},
    })
    .onConflictDoUpdate({
      target: formsTable.id,
      set: {
        title: body.data.title.trim(),
        headers: body.data.headers,
        questions: body.data.questions,
        selectionHistory: {},
        updatedAt: new Date(),
      },
    })
    .returning();

  res.json(SaveAdminFormResponse.parse(adminFormResponse(form)));
});

router.get("/admin/responses/:formId", async (req, res): Promise<void> => {
  if (!requireAdmin(req, res)) return;
  const params = ListAdminResponsesParams.safeParse({
    formId: formIdFromParams(req.params),
  });
  if (!params.success) {
    res.status(400).json({ error: params.error.message });
    return;
  }
  const rows = await db
    .select()
    .from(responsesTable)
    .where(eq(responsesTable.formId, params.data.formId))
    .orderBy(desc(responsesTable.createdAt));
  res.json(
    ListAdminResponsesResponse.parse(
      rows.map((response) => ({
        id: response.id,
        formId: response.formId,
        respondentRole: response.respondentRole,
        answers: response.answers,
        createdAt: response.createdAt,
      })),
    ),
  );
});

export default router;