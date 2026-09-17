import {
  useEffect,
  useRef,
  useState,
  type ChangeEvent,
  type DragEvent,
  type ReactNode,
} from "react";
import {
  FileSpreadsheet,
  Upload,
  ArrowRight,
  ArrowLeft,
  Check,
  CircleHelp,
  Database,
  FileText,
  RotateCcw,
  Settings2,
  Sparkles,
  Trash2,
  X,
  AlertTriangle,
  CheckCircle2,
  ChevronRight,
  Download,
  UserRound,
  Link2,
  LockKeyhole,
  LogOut,
  RefreshCw,
  Users,
  Clipboard,
  CheckCheck,
} from "lucide-react";
import * as XLSX from "xlsx";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { ErrorBoundary } from "@/components/error-boundary";
import { Toaster } from "@/components/ui/toaster";
import { TooltipProvider } from "@/components/ui/tooltip";
import NotFound from "@/pages/not-found";
import {
  Route,
  Switch,
  useLocation,
  Router as WouterRouter,
} from "wouter";

type Question = { id: string; text: string; sourceRow?: string[] };
type View = "build" | "respond" | "results";
type ParsedBank = { headers: string[]; questions: Question[] };
type StoredBank = {
  title: string;
  headers: string[];
  questions: Question[];
  fileName?: string;
  selectionHistory?: Record<string, number>;
};
type PublicForm = {
  id: string;
  title: string;
  headers: string[];
  questions: Question[];
};
type SubmittedResponse = {
  id: string;
  formId: string;
  respondentRole: string;
  answers: Array<{
    questionId: string;
    questionText: string;
    answer: string;
    sourceRow?: string[];
  }>;
  createdAt: string;
};

const queryClient = new QueryClient();
const STORAGE_KEY = "randomized-forms-bank-v1";
const ATTACHED_WORKBOOK_URL = "/robothaqafa_questions.xlsx";
const DEFAULT_HEADERS = [
  "topic",
  "scenario(AR)",
  "scenario (ENG)",
  "complexity",
  "english summary",
  "messages",
];

const sampleQuestions: Question[] = [
  {
    id: "sample-01",
    text: "What is one idea from this session you would like to test in your own work?",
  },
  {
    id: "sample-02",
    text: "Which part of the process felt clearest to you, and why?",
  },
  {
    id: "sample-03",
    text: "Describe a moment when a different perspective changed your approach.",
  },
  {
    id: "sample-04",
    text: "What would make this experience more useful for your team or learners?",
  },
  {
    id: "sample-05",
    text: "Where do you notice the biggest gap between intention and action?",
  },
  {
    id: "sample-06",
    text: "What is one assumption you are leaving behind today?",
  },
  {
    id: "sample-07",
    text: "How would you explain the most important takeaway to a colleague?",
  },
  {
    id: "sample-08",
    text: "What evidence would help you decide whether this approach is working?",
  },
  {
    id: "sample-09",
    text: "Which question do you wish we had asked, and what would it reveal?",
  },
  {
    id: "sample-10",
    text: "What is a small next step you can take within the next week?",
  },
  {
    id: "sample-11",
    text: "When does this topic feel most relevant to your day-to-day work?",
  },
  {
    id: "sample-12",
    text: "What would you like to investigate more deeply from here?",
  },
];

function parseCSV(csv: string): ParsedBank {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let quoted = false;
  for (let index = 0; index < csv.length; index += 1) {
    const character = csv[index];
    const next = csv[index + 1];
    if (character === '"' && quoted && next === '"') {
      cell += '"';
      index += 1;
      continue;
    }
    if (character === '"') {
      quoted = !quoted;
      continue;
    }
    if (character === "," && !quoted) {
      row.push(cell.trim());
      cell = "";
      continue;
    }
    if ((character === "\n" || character === "\r") && !quoted) {
      if (character === "\r" && next === "\n") index += 1;
      row.push(cell.trim());
      if (row.some(Boolean)) rows.push(row);
      row = [];
      cell = "";
      continue;
    }
    cell += character;
  }
  row.push(cell.trim());
  if (row.some(Boolean)) rows.push(row);
  if (!rows.length) return { headers: [], questions: [] };
  return questionsFromRows(rows);
}

function questionColumnFor(values: string[]) {
  return values.findIndex((value) => value.trim().length > 0);
}

function questionsFromRows(rows: string[][]): ParsedBank {
  if (!rows.length) return { headers: [], questions: [] };
  const scenarioIndex = rows[0].findIndex((value) =>
    /scen(?:ario|ario)|سناريو/i.test(value),
  );
  const genericHeaderIndex = rows[0].findIndex((value) =>
    /question|prompt|text|item/i.test(value),
  );
  const headerIndex = scenarioIndex >= 0 ? scenarioIndex : genericHeaderIndex;
  const hasHeader = headerIndex >= 0;
  const questionColumn = hasHeader
    ? headerIndex
    : rows.reduce(
        (best, values) =>
          values[best]?.length > values[questionColumnFor(values)]?.length
            ? best
            : questionColumnFor(values),
        0,
      );
  const content = hasHeader ? rows.slice(1) : rows;
  const headers = hasHeader ? rows[0] : [];
  const questions = content
    .map((values, index) => ({
      text: values[questionColumn]?.trim(),
      sourceRow: values,
      index,
    }))
    .filter(
      (item): item is { text: string; sourceRow: string[]; index: number } =>
        Boolean(item.text),
    )
    .map(({ text, sourceRow, index }) => ({
      id: `import-${Date.now()}-${index}`,
      text,
      sourceRow,
    }));
  return { headers, questions };
}

async function readUpload(file: File): Promise<ParsedBank> {
  const lowerName = file.name.toLowerCase();
  if (
    lowerName.endsWith(".csv") ||
    file.type.includes("csv") ||
    file.type.startsWith("text/")
  ) {
    return parseCSV(await file.text());
  }
  if (
    lowerName.endsWith(".xlsx") ||
    lowerName.endsWith(".xls") ||
    file.type.includes("spreadsheet")
  ) {
    const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    if (!firstSheet)
      throw new Error("This workbook does not contain a readable sheet.");
    const rows = XLSX.utils.sheet_to_json<unknown[]>(firstSheet, {
      header: 1,
      blankrows: false,
    });
    return questionsFromRows(
      rows.map((row) => row.map((value) => String(value ?? ""))),
    );
  }
  throw new Error("Upload an Excel workbook (.xlsx or .xls) or a CSV file.");
}

async function readAttachedWorkbook(): Promise<ParsedBank> {
  const response = await fetch(ATTACHED_WORKBOOK_URL);
  if (!response.ok)
    throw new Error("The attached question bank could not be loaded.");
  const workbook = XLSX.read(await response.arrayBuffer(), { type: "array" });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!firstSheet)
    throw new Error("The attached workbook does not contain a readable sheet.");
  const rows = XLSX.utils.sheet_to_json<unknown[]>(firstSheet, {
    header: 1,
    blankrows: false,
  });
  return questionsFromRows(
    rows.map((row) => row.map((value) => String(value ?? ""))),
  );
}

function scenarioPrompt(question: Question) {
  return `تخيل نفسك بالسيناريو التالي: ${question.text} أجب حسب المطلوب.`;
}

function responseMessages(question: Question, answer: string) {
  return JSON.stringify(
    [
      {
        role: "assistant",
        content: scenarioPrompt(question),
      },
      { role: "user", content: answer },
    ],
    null,
    2,
  );
}

function downloadResponses(
  headers: string[],
  questions: Question[],
  role: string,
  answers: Record<string, string>,
  title: string,
) {
  const outputHeaders = headers.length ? [...headers] : [...DEFAULT_HEADERS];
  const roleColumn = outputHeaders.findIndex(
    (header) => header.trim().toLowerCase() === "responder role",
  );
  const roleIndex =
    roleColumn >= 0 ? roleColumn : outputHeaders.push("Responder Role") - 1;
  const messageColumn = outputHeaders.findIndex(
    (header) => header.trim().toLowerCase() === "messages",
  );
  const messagesIndex =
    messageColumn >= 0 ? messageColumn : outputHeaders.push("messages") - 1;
  const rows = questions.map((question) => {
    const row = Array.from(
      { length: outputHeaders.length },
      (_, index) => question.sourceRow?.[index] ?? "",
    );
    row[roleIndex] = role;
    row[messagesIndex] = responseMessages(question, answers[question.id] ?? "");
    return row;
  });
  const worksheet = XLSX.utils.aoa_to_sheet([outputHeaders, ...rows]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Responses");
  const filename = `${(title || "randomized-form").replace(/[^a-z0-9\u0600-\u06ff]+/gi, "-").replace(/^-|-$/g, "") || "randomized-form"}-responses.xlsx`;
  XLSX.writeFile(workbook, filename);
}

function shuffle<T>(items: T[]) {
  return [...items].sort(() => Math.random() - 0.5);
}

function BrandMark() {
  return (
    <div className="brand-mark" aria-hidden="true">
      <span />
      <span />
      <span />
    </div>
  );
}

function Sidebar({
  view,
  onNavigate,
  bankCount,
}: {
  view: View;
  onNavigate: (view: View) => void;
  bankCount: number;
}) {
  const navItems: { id: View; label: string; icon: typeof Settings2 }[] = [
    { id: "build", label: "Build a form", icon: Settings2 },
    { id: "respond", label: "Preview & respond", icon: FileText },
  ];
  return (
    <aside className="app-sidebar">
      <div className="sidebar-brand">
        <BrandMark />
        <span>
          Najdi Culture<span className="brand-dot">.</span>
        </span>
      </div>
      <p className="sidebar-kicker">Question fairness, simplified</p>
      <nav className="sidebar-nav" aria-label="Primary navigation">
        {navItems.map(({ id, label, icon: Icon }) => (
          <button
            key={id}
            type="button"
            data-testid={`button-nav-${id}`}
            className={`sidebar-link ${view === id || (id === "respond" && view === "results") ? "active" : ""}`}
            onClick={() => onNavigate(id)}
          >
            <Icon size={17} strokeWidth={1.8} />
            <span>{label}</span>
            {id === "respond" && bankCount >= 5 ? (
              <span className="nav-status">
                <Check size={12} />
              </span>
            ) : null}
          </button>
        ))}
      </nav>
      <div className="sidebar-bottom">
        <div className="sidebar-bank">
          <div className="sidebar-bank-icon">
            <Database size={16} />
          </div>
          <div>
            <span className="eyebrow">Current bank</span>
            <strong data-testid="text-sidebar-bank-count">
              {bankCount} questions
            </strong>
          </div>
        </div>
        <div className="sidebar-note">
          <CircleHelp size={14} />
          <span>Five fresh questions, balanced over time.</span>
        </div>
        <span className="sidebar-version">LOCAL WORKSPACE · v1.0</span>
      </div>
    </aside>
  );
}

function Topbar({
  view,
  onNavigate,
}: {
  view: View;
  onNavigate: (view: View) => void;
}) {
  return (
    <header className="mobile-topbar">
      <button
        className="mobile-brand"
        type="button"
        data-testid="button-mobile-home"
        onClick={() => onNavigate("build")}
      >
        <BrandMark />
        <span>
          Najdi Culture<span className="brand-dot">.</span>
        </span>
      </button>
      <span className="mobile-context">
        {view === "build"
          ? "Builder"
          : view === "respond"
            ? "Respondent preview"
            : "Submitted"}
      </span>
    </header>
  );
}

function UploadDropzone({
  onFile,
  isLoading,
}: {
  onFile: (file: File) => void;
  isLoading: boolean;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);
  const handleDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDragging(false);
    const file = event.dataTransfer.files[0];
    if (file) onFile(file);
  };
  const handleChange = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) onFile(file);
  };
  return (
    <div
      className={`upload-dropzone ${dragging ? "dragging" : ""}`}
      onDragOver={(event) => {
        event.preventDefault();
        setDragging(true);
      }}
      onDragLeave={() => setDragging(false)}
      onDrop={handleDrop}
    >
      <input
        ref={inputRef}
        className="sr-only"
        data-testid="input-workbook"
        type="file"
        accept=".xlsx,.xls,.csv,text/csv"
        onChange={handleChange}
      />
      <div className="upload-icon">
        <Upload size={21} />
      </div>
      <div className="upload-copy">
        <strong>
          {isLoading
            ? "Reading your question bank…"
            : "Drop a question bank here"}
        </strong>
        <span>Excel or CSV · one question per row</span>
      </div>
      <button
        type="button"
        data-testid="button-browse-workbook"
        className="button button-secondary button-small"
        onClick={() => inputRef.current?.click()}
        disabled={isLoading}
      >
        {isLoading ? "Working" : "Browse files"}
      </button>
    </div>
  );
}

function EmptyBank({ onSample }: { onSample: () => void }) {
  return (
    <div className="empty-state" data-testid="status-empty-bank">
      <div className="empty-art">
        <FileSpreadsheet size={27} />
      </div>
      <h3>Your question bank is waiting</h3>
      <p>
        Upload at least five questions to make a fair, randomized form. Not
        ready with a file? Start with the sample bank.
      </p>
      <button
        type="button"
        data-testid="button-empty-sample"
        className="button button-primary"
        onClick={onSample}
      >
        <Sparkles size={16} /> Load sample bank
      </button>
    </div>
  );
}

function BuildView({
  title,
  setTitle,
  questions,
  fileName,
  error,
  isLoading,
  onFile,
  onSample,
  onLoadAttached,
  onClear,
  onNavigate,
  onPublish,
  isPublishing = false,
  published = false,
  shareLink,
}: {
  title: string;
  setTitle: (title: string) => void;
  questions: Question[];
  fileName?: string;
  error: string;
  isLoading: boolean;
  onFile: (file: File) => void;
  onSample: () => void;
  onLoadAttached: () => void;
  onClear: () => void;
  onNavigate: (view: View) => void;
  onPublish?: () => void;
  isPublishing?: boolean;
  published?: boolean;
  shareLink?: string;
}) {
  const ready = questions.length >= 5;
  const currentRate = questions.length ? (5 / questions.length) * 100 : 0;
  const currentRateLabel =
    currentRate === 10
      ? "10% with this bank"
      : `about ${currentRate < 1 ? currentRate.toFixed(1) : Math.round(currentRate)}% with ${questions.length} questions`;
  return (
    <main className="page-content">
      <div className="page-heading reveal-up">
        <div>
          <span className="eyebrow accent-eyebrow">
            FORM BUILDER <span className="eyebrow-line" />
          </span>
          <h1>Make every question count.</h1>
          <p>
            Build one considered prompt set. Randomized delivery keeps every
            response fresh and every voice in the room.
          </p>
        </div>
        <div
          className={`ready-pill ${ready ? "ready" : "not-ready"}`}
          data-testid="status-bank-ready"
        >
          {ready ? <CheckCircle2 size={15} /> : <AlertTriangle size={15} />}
          <span>{ready ? "Ready to collect" : "Needs five questions"}</span>
        </div>
      </div>
      <section className="builder-hero reveal-up" aria-label="Form overview">
        <div className="hero-copy">
          <span className="hero-overline">THE FAIRNESS ENGINE</span>
          <h2>
            Different questions.
            <br />
            <em>Same chance to be heard.</em>
          </h2>
          <p>
            Respondents receive a private, unpredictable set of five from your
            bank. No two paths need to look alike.
          </p>
          <button
            type="button"
            data-testid="button-hero-respond"
            className="hero-link"
            onClick={() => onNavigate("respond")}
            disabled={!ready}
          >
            Preview respondent view <ArrowRight size={16} />
          </button>
        </div>
        <div className="hero-visual" aria-hidden="true">
          <div className="orbit orbit-one" />
          <div className="orbit orbit-two" />
          <div className="hero-number">05</div>
          <div className="hero-caption">
            questions
            <br />
            per response
          </div>
          <div className="spark spark-one" />
          <div className="spark spark-two" />
        </div>
      </section>
      <div className="builder-grid">
        <section className="panel import-panel reveal-up delay-one">
          <div className="panel-heading">
            <div>
              <span className="eyebrow">01 · SOURCE</span>
              <h2>Bring your question bank</h2>
            </div>
            <FileSpreadsheet className="panel-heading-icon" size={21} />
          </div>
          <p className="panel-intro">
            Keep it simple: a single question column is all you need. We ignore
            blank rows and keep your wording intact.
          </p>
          <UploadDropzone onFile={onFile} isLoading={isLoading} />
          <button
            type="button"
            data-testid="button-load-attached-workbook"
            className="text-button attached-workbook-button"
            onClick={onLoadAttached}
            disabled={isLoading}
          >
            <FileSpreadsheet size={14} /> Load attached Arabic question bank
          </button>
          {error ? (
            <div
              className="error-banner"
              role="alert"
              data-testid="status-upload-error"
            >
              <AlertTriangle size={16} />
              <span>{error}</span>
              <button
                type="button"
                aria-label="Dismiss upload error"
                data-testid="button-dismiss-error"
                onClick={() => onFile(new File([], ""))}
              >
                <X size={15} />
              </button>
            </div>
          ) : null}
          <div className="file-meta">
            {fileName ? (
              <>
                <span className="file-badge">
                  <Check size={12} />
                </span>
                <span data-testid="text-uploaded-filename">{fileName}</span>
                <span className="file-meta-separator">·</span>
                <span>{questions.length} imported</span>
              </>
            ) : (
              <>
                <span className="file-badge muted">
                  <FileText size={12} />
                </span>
                <span>Sample bank loaded</span>
                <span className="file-meta-separator">·</span>
                <span>Stored locally</span>
              </>
            )}
          </div>
        </section>
        <section className="panel settings-panel reveal-up delay-two">
          <div className="panel-heading">
            <div>
              <span className="eyebrow">02 · SHAPE</span>
              <h2>Name the experience</h2>
            </div>
            <Settings2 className="panel-heading-icon" size={21} />
          </div>
          <label className="field-label" htmlFor="form-title">
            Form title
          </label>
          <input
            id="form-title"
            data-testid="input-form-title"
            className="text-input"
            value={title}
            onChange={(event) => setTitle(event.target.value)}
            placeholder="e.g. Week 4 reflection"
            maxLength={72}
          />
          <p className="field-help">
            This is the first thing respondents will see.
          </p>
          <div className="setting-rule" />
          <div className="setting-row">
            <div className="setting-icon">
              <Database size={16} />
            </div>
            <div>
              <strong>Five questions per response</strong>
              <span>Selected without replacement</span>
            </div>
            <span className="fixed-value">FIXED</span>
          </div>
          <div className="setting-row">
            <div className="setting-icon orange">
              <Sparkles size={16} />
            </div>
            <div>
              <strong>Balanced question exposure</strong>
              <span>{currentRateLabel}; 10% needs 50 questions</span>
            </div>
            <span className="fixed-value">ON</span>
          </div>
          <p className="field-help setting-help">
            The form always shows five questions and prioritizes prompts that
            have appeared least often, so the distribution stays even over time.
          </p>
        </section>
      </div>
      <section className="panel bank-panel reveal-up delay-three">
        <div className="bank-panel-heading">
          <div>
            <span className="eyebrow">03 · REVIEW</span>
            <h2>
              Question bank{" "}
              <span className="count-badge" data-testid="text-question-count">
                {questions.length}
              </span>
            </h2>
          </div>
          <div className="bank-actions">
            <button
              type="button"
              data-testid="button-load-sample"
              className="text-button"
              onClick={onSample}
            >
              <RotateCcw size={14} /> Use sample
            </button>
            <button
              type="button"
              data-testid="button-clear-bank"
              className="icon-button"
              onClick={onClear}
              aria-label="Clear question bank"
            >
              <Trash2 size={16} />
            </button>
          </div>
        </div>
        {questions.length ? (
          <div className="question-list">
            {questions.slice(0, 7).map((question, index) => (
              <div
                className="question-row"
                key={question.id}
                data-testid={`row-question-${question.id}`}
              >
                <span className="question-number">
                  {String(index + 1).padStart(2, "0")}
                </span>
                <span>{question.text}</span>
                <ChevronRight size={15} className="question-chevron" />
              </div>
            ))}
            {questions.length > 7 ? (
              <div className="more-row">
                + {questions.length - 7} more questions in this bank
              </div>
            ) : null}
          </div>
        ) : (
          <EmptyBank onSample={onSample} />
        )}
        <div className="bank-footer">
          <span>
            <span className="status-dot" />{" "}
            {ready
              ? "Ready to generate balanced respondent sets"
              : "Upload five or more questions to continue"}
          </span>
          <button
            type="button"
            data-testid="button-review-respondent"
            className="button button-primary"
            disabled={!ready}
            onClick={() => onNavigate("respond")}
          >
            Open respondent preview <ArrowRight size={16} />
          </button>
          {onPublish ? (
            <button
              type="button"
              data-testid="button-publish-form"
              className="button button-secondary"
              disabled={!ready || isPublishing}
              onClick={onPublish}
            >
              {isPublishing ? (
                <>
                  <RefreshCw size={16} className="spin" /> Saving…
                </>
              ) : published ? (
                <>
                  <CheckCheck size={16} /> Update shared form
                </>
              ) : (
                <>
                  <Link2 size={16} /> Publish shared form
                </>
              )}
            </button>
          ) : null}
        </div>
        {shareLink ? (
          <div className="share-link-row" data-testid="status-share-link">
            <Link2 size={15} />
            <span>
              <strong>Share this respondent link</strong>
              <small>{shareLink}</small>
            </span>
          </div>
        ) : null}
      </section>
    </main>
  );
}

function RespondView({
  title,
  role,
  setRole,
  questions,
  onBack,
  onSubmit,
}: {
  title: string;
  role: string;
  setRole: (role: string) => void;
  questions: Question[];
  onBack: () => void;
  onSubmit: (answers: Record<string, string>) => void;
}) {
  const [answers, setAnswers] = useState<Record<string, string>>({});
  const complete =
    role.trim().length > 0 &&
    questions.length === 5 &&
    questions.every((question) => answers[question.id]?.trim());
  const setAnswer = (id: string, value: string) =>
    setAnswers((current) => ({ ...current, [id]: value }));
  return (
    <main className="respond-page" dir="rtl">
      <div className="respond-topline">
        <button
          type="button"
          data-testid="button-back-builder"
          className="back-button"
          onClick={onBack}
        >
          <ArrowLeft size={16} /> Back to builder
        </button>
        <span className="respond-privacy">
          <span className="privacy-mark" /> PRIVATE RESPONSE · LOCAL ONLY
        </span>
      </div>
      <div className="respond-intro reveal-up">
        <span className="eyebrow accent-eyebrow">
          استبيان جمع بيانات باللغة النجدية
        </span>
        <h1 data-testid="text-respond-title">{title || "استبيان جديد"}</h1>
        <p>
          أجب عن الدور أولاً، ثم أجب على خمسة سيناريوهات مختارة لك{" "}
          <b>باللهجة النجدية</b>. لا توجد إجابات صحيحة أو خاطئة.
        </p>
        <div className="progress-meta">
          <span>التقدم</span>
          <span data-testid="text-progress">
            {Object.values(answers).filter((answer) => answer.trim()).length} /
            5 أسئلة مكتملة
          </span>
        </div>
        <div className="progress-track">
          <div
            className="progress-fill"
            style={{
              width: `${((Object.values(answers).filter((answer) => answer.trim()).length + (role.trim() ? 1 : 0)) / 6) * 100}%`,
            }}
          />
        </div>
      </div>
      <form
        className="respond-form"
        onSubmit={(event) => {
          event.preventDefault();
          if (complete) onSubmit(answers);
        }}
      >
        <fieldset className="respond-card role-card reveal-up">
          <legend>
            <span className="question-number large">
              <UserRound size={19} />
            </span>
            <span>بيانات المجيب</span>
          </legend>
          <label htmlFor="respondent-role">ما هو دورك أو مسماك الوظيفي؟</label>
          <input
            id="respondent-role"
            data-testid="input-respondent-role"
            className="text-input"
            value={role}
            onChange={(event) => setRole(event.target.value)}
            placeholder="مثال: معلم، طالب، باحث، ولي أمر"
          />
        </fieldset>
        {questions.map((question, index) => (
          <fieldset
            className="respond-card reveal-up"
            style={{ animationDelay: `${index * 55}ms` }}
            key={question.id}
          >
            <legend>
              <span className="question-number large">
                {String(index + 1).padStart(2, "0")}
              </span>
              <span>السؤال {index + 1} من 5</span>
            </legend>
            <label htmlFor={`answer-${question.id}`}>
              {scenarioPrompt(question)}
            </label>
            <textarea
              id={`answer-${question.id}`}
              data-testid={`textarea-answer-${index + 1}`}
              value={answers[question.id] ?? ""}
              onChange={(event) => setAnswer(question.id, event.target.value)}
              placeholder="اكتب إجابتك هنا…"
              rows={4}
            />
          </fieldset>
        ))}
        <div className="submit-row">
          <span className="submit-note">
            <CircleHelp size={15} /> تحفظ الإجابات في هذا المتصفح فقط.
          </span>
          <button
            type="submit"
            data-testid="button-submit-response"
            className="button button-primary button-submit"
            disabled={!complete}
          >
            إرسال الإجابات <ArrowLeft size={17} />
          </button>
        </div>
      </form>
    </main>
  );
}

function ResultsView({
  title,
  role,
  questions,
  answers,
  onAgain,
  onBuild,
  onDownload,
}: {
  title: string;
  role: string;
  questions: Question[];
  answers: Record<string, string>;
  onAgain: () => void;
  onBuild: () => void;
  onDownload: () => void;
}) {
  return (
    <main className="results-page" dir="rtl">
      <div className="results-card reveal-up">
        <div className="success-mark">
          <Check size={28} strokeWidth={2.5} />
        </div>
        <span className="eyebrow accent-eyebrow">تم تسجيل الإجابات</span>
        <h1>
          شكراً لمشاركتك
          <br />
          <em>في هذا الاستبيان.</em>
        </h1>
        <p>
          تم تسجيل إجابة <strong>{role || "المجيب"}</strong> محلياً لنموذج{" "}
          <strong>{title || "هذا الاستبيان"}</strong>. حمّل ملف Excel لحفظ
          النتيجة.
        </p>
        <div className="result-rule" />
        <div className="response-receipt">
          <div className="receipt-heading">
            <span>الأسئلة الخمسة</span>
            <span>تمت الإجابة</span>
          </div>
          {questions.map((question, index) => (
            <div className="receipt-row" key={question.id}>
              <span className="question-number">
                {String(index + 1).padStart(2, "0")}
              </span>
              <span>{scenarioPrompt(question)}</span>
              <Check size={15} className="receipt-check" />
            </div>
          ))}
        </div>
        <div className="results-actions">
          <button
            type="button"
            data-testid="button-download-responses"
            className="button button-primary"
            onClick={onDownload}
          >
            <Download size={16} /> تحميل ملف Excel
          </button>
          <button
            type="button"
            data-testid="button-new-response"
            className="button button-quiet"
            onClick={onAgain}
          >
            <RotateCcw size={16} /> إجابة جديدة
          </button>
          <button
            type="button"
            data-testid="button-return-builder"
            className="button button-quiet"
            onClick={onBuild}
          >
            العودة إلى الإعداد <ArrowRight size={16} />
          </button>
        </div>
      </div>
      <div className="results-aside">
        <Sparkles size={17} />
        <span>كل إجابة تُحفظ في صف مستقل.</span>
        <p>يتم تكرار صف كل سؤال مع وضع دور المجيب في عمود مستقل، والرسائل بصيغة JSON.</p>
      </div>
    </main>
  );
}

function AdminLoginView({
  password,
  setPassword,
  onSubmit,
  error,
  isLoading,
}: {
  password: string;
  setPassword: (value: string) => void;
  onSubmit: () => void;
  error: string;
  isLoading: boolean;
}) {
  return (
    <main className="auth-page">
      <div className="auth-card reveal-up">
        <div className="auth-icon">
          <LockKeyhole size={24} />
        </div>
        <span className="eyebrow accent-eyebrow">ADMIN WORKSPACE</span>
        <h1>Keep every response in view.</h1>
        <p>
          Sign in to publish the question bank, copy the respondent link, and
          download all collected answers.
        </p>
        <form
          onSubmit={(event) => {
            event.preventDefault();
            onSubmit();
          }}
        >
          <label className="field-label" htmlFor="admin-password">
            Admin password
          </label>
          <input
            id="admin-password"
            data-testid="input-admin-password"
            className="text-input"
            type="password"
            value={password}
            onChange={(event) => setPassword(event.target.value)}
            autoComplete="current-password"
            autoFocus
          />
          {error ? (
            <div className="error-banner" role="alert">
              <AlertTriangle size={16} />
              <span>{error}</span>
            </div>
          ) : null}
          <button
            type="submit"
            data-testid="button-admin-login"
            className="button button-primary auth-submit"
            disabled={!password || isLoading}
          >
            {isLoading ? "Checking…" : "Open admin dashboard"}{" "}
            <ArrowRight size={16} />
          </button>
        </form>
      </div>
    </main>
  );
}

function AdminResponsesView({
  responses,
  headers,
  title,
  isLoading,
  error,
  onRefresh,
  onDownload,
}: {
  responses: SubmittedResponse[];
  headers: string[];
  title: string;
  isLoading: boolean;
  error: string;
  onRefresh: () => void;
  onDownload: () => void;
}) {
  return (
    <section className="panel admin-responses-panel reveal-up">
      <div className="bank-panel-heading">
        <div>
          <span className="eyebrow">04 · RESPONSES</span>
          <h2>
            Response inbox{" "}
            <span className="count-badge" data-testid="text-response-count">
              {responses.length}
            </span>
          </h2>
        </div>
        <div className="bank-actions">
          <button
            type="button"
            className="text-button"
            data-testid="button-refresh-responses"
            onClick={onRefresh}
            disabled={isLoading}
          >
            <RefreshCw size={14} className={isLoading ? "spin" : ""} /> Refresh
          </button>
          <button
            type="button"
            className="button button-primary button-small"
            data-testid="button-download-all-responses"
            onClick={onDownload}
            disabled={!responses.length}
          >
            <Download size={15} /> Download Excel
          </button>
        </div>
      </div>
      {error ? (
        <div className="error-banner" role="alert">
          <AlertTriangle size={16} />
          <span>{error}</span>
        </div>
      ) : null}
      {isLoading && !responses.length ? (
        <div className="admin-empty">
          <RefreshCw size={20} className="spin" /> Loading responses…
        </div>
      ) : responses.length ? (
        <div className="response-table-wrap">
          <table className="response-table">
            <thead>
              <tr>
                <th>Received</th>
                <th>Role</th>
                <th>Answered</th>
                <th>Response ID</th>
              </tr>
            </thead>
            <tbody>
              {responses.map((response) => (
                <tr key={response.id}>
                  <td>
                    {new Intl.DateTimeFormat("en", {
                      dateStyle: "medium",
                      timeStyle: "short",
                    }).format(new Date(response.createdAt))}
                  </td>
                  <td className="response-role">{response.respondentRole}</td>
                  <td>{response.answers.length} / 5</td>
                  <td className="response-id">{response.id.slice(0, 8)}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <p className="table-help">
            Each Excel export repeats the original source row once per answered
            scenario, stores the responder role in its own column, and places
            the assistant prompt before the user answer in messages.
          </p>
        </div>
      ) : (
        <div className="admin-empty">
          <Users size={21} />
          <span>No responses yet. Share the respondent link to begin.</span>
        </div>
      )}
    </section>
  );
}

function downloadAllResponses(
  headers: string[],
  responses: SubmittedResponse[],
  title: string,
) {
  const outputHeaders = headers.length ? [...headers] : [...DEFAULT_HEADERS];
  const roleColumn = outputHeaders.findIndex(
    (header) => header.trim().toLowerCase() === "responder role",
  );
  const roleIndex =
    roleColumn >= 0 ? roleColumn : outputHeaders.push("Responder Role") - 1;
  const messageColumn = outputHeaders.findIndex(
    (header) => header.trim().toLowerCase() === "messages",
  );
  const messagesIndex =
    messageColumn >= 0 ? messageColumn : outputHeaders.push("messages") - 1;
  const rows = responses.flatMap((response) =>
    response.answers.map((answer) => {
      const question: Question = {
        id: answer.questionId,
        text: answer.questionText,
        sourceRow: answer.sourceRow,
      };
      const row = Array.from(
        { length: outputHeaders.length },
        (_, index) => question.sourceRow?.[index] ?? "",
      );
      row[roleIndex] = response.respondentRole;
      row[messagesIndex] = responseMessages(question, answer.answer);
      return row;
    }),
  );
  const worksheet = XLSX.utils.aoa_to_sheet([outputHeaders, ...rows]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Responses");
  const filename = `${(title || "randomized-form")
    .replace(/[^a-z0-9\u0600-\u06ff]+/gi, "-")
    .replace(/^-|-$/g, "") || "randomized-form"}-responses.xlsx`;
  XLSX.writeFile(workbook, filename);
}

function PublicThanks({ title }: { title: string }) {
  return (
    <main className="results-page" dir="rtl">
      <div className="results-card reveal-up">
        <div className="success-mark">
          <Check size={28} strokeWidth={2.5} />
        </div>
        <span className="eyebrow accent-eyebrow">تم تسجيل الإجابات</span>
        <h1>
          شكراً لمشاركتك
          <br />
          <em>في هذا الاستبيان.</em>
        </h1>
        <p>
          تم حفظ إجاباتك بنجاح في لوحة الإدارة لنموذج{" "}
          <strong>{title || "هذا الاستبيان"}</strong>.
        </p>
      </div>
    </main>
  );
}

function PublicFormPage() {
  const [, setLocation] = useLocation();
  const [form, setForm] = useState<PublicForm>();
  const [isLoading, setIsLoading] = useState(true);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [submitted, setSubmitted] = useState(false);
  const [role, setRole] = useState("");
  const [answers, setAnswers] = useState<Record<string, string>>({});
  const [error, setError] = useState("");

  useEffect(() => {
    let cancelled = false;
    fetch("/api/forms/default")
      .then(async (response) => {
        if (!response.ok) {
          const payload = (await response.json().catch(() => null)) as
            | { error?: string }
            | null;
          throw new Error(
            payload?.error === "Form not found"
              ? "This form has not been published yet."
              : payload?.error || "The form could not be loaded.",
          );
        }
        return response.json() as Promise<PublicForm>;
      })
      .then((loaded) => {
        if (!cancelled) setForm(loaded);
      })
      .catch((loadError) => {
        if (!cancelled)
          setError(
            loadError instanceof Error
              ? loadError.message
              : "The form could not be loaded.",
          );
      })
      .finally(() => {
        if (!cancelled) setIsLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, []);

  const submit = async () => {
    if (!form || !role.trim() || form.questions.some((question) => !answers[question.id]?.trim())) {
      return;
    }
    setError("");
    setIsSubmitting(true);
    try {
      const response = await fetch(`/api/forms/${form.id}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          respondentRole: role.trim(),
          answers: form.questions.map((question) => ({
            questionId: question.id,
            questionText: question.text,
            sourceRow: question.sourceRow,
            answer: answers[question.id].trim(),
          })),
        }),
      });
      if (!response.ok) {
        const payload = (await response.json().catch(() => null)) as
          | { error?: string }
          | null;
        throw new Error(payload?.error || "Your response could not be saved.");
      }
      setSubmitted(true);
    } catch (submitError) {
      setError(
        submitError instanceof Error
          ? submitError.message
          : "Your response could not be saved.",
      );
    } finally {
      setIsSubmitting(false);
    }
  };

  if (isLoading) {
    return (
      <main className="auth-page">
        <div className="admin-empty">
          <RefreshCw size={20} className="spin" /> Loading form…
        </div>
      </main>
    );
  }
  if (submitted && form) return <PublicThanks title={form.title} />;
  if (!form) {
    return (
      <main className="auth-page">
        <div className="auth-card reveal-up">
          <div className="auth-icon">
            <AlertTriangle size={24} />
          </div>
          <span className="eyebrow accent-eyebrow">FORM UNAVAILABLE</span>
          <h1>There is nothing to answer yet.</h1>
          <p>{error || "The admin needs to publish this form first."}</p>
          <button
            type="button"
            className="button button-primary auth-submit"
            onClick={() => setLocation("/admin")}
          >
            Open admin page <ArrowRight size={16} />
          </button>
        </div>
      </main>
    );
  }

  return (
    <main className="public-form-page" dir="rtl">
      <div className="public-form-topline">
        <div className="public-brand">
          <BrandMark />
          <span>
            Najdi Culture<span className="brand-dot">.</span>
          </span>
        </div>
        <span className="respond-privacy">
          <span className="privacy-mark" /> SECURE RESPONSE
        </span>
      </div>
      <div className="respond-intro reveal-up">
        <span className="eyebrow accent-eyebrow">
          استبيان جمع بيانات باللغة النجدية
        </span>
        <h1 data-testid="text-respond-title">{form.title}</h1>
        <p>
          أجب عن الدور أولاً، ثم أجب على خمسة سيناريوهات مختارة لك{" "}
          <b>باللهجة النجدية</b>. لا توجد إجابات صحيحة أو خاطئة.
        </p>
        <div className="progress-meta">
          <span>التقدم</span>
          <span data-testid="text-progress">
            {Object.values(answers).filter((answer) => answer.trim()).length} / 5
            أسئلة مكتملة
          </span>
        </div>
        <div className="progress-track">
          <div
            className="progress-fill"
            style={{
              width: `${((Object.values(answers).filter((answer) => answer.trim()).length + (role.trim() ? 1 : 0)) / 6) * 100}%`,
            }}
          />
        </div>
      </div>
      <form
        className="respond-form"
        onSubmit={(event) => {
          event.preventDefault();
          void submit();
        }}
      >
        <fieldset className="respond-card role-card reveal-up">
          <legend>
            <span className="question-number large">
              <UserRound size={19} />
            </span>
            <span>بيانات المجيب</span>
          </legend>
          <label htmlFor="respondent-role">ما هو دورك أو مسماك الوظيفي؟</label>
          <input
            id="respondent-role"
            data-testid="input-respondent-role"
            className="text-input"
            value={role}
            onChange={(event) => setRole(event.target.value)}
            placeholder="مثال: معلم، طالب، باحث، ولي أمر"
          />
        </fieldset>
        {form.questions.map((question, index) => (
          <fieldset
            className="respond-card reveal-up"
            style={{ animationDelay: `${index * 55}ms` }}
            key={question.id}
          >
            <legend>
              <span className="question-number large">
                {String(index + 1).padStart(2, "0")}
              </span>
              <span>السؤال {index + 1} من 5</span>
            </legend>
            <label htmlFor={`answer-${question.id}`}>
              {scenarioPrompt(question)}
            </label>
            <textarea
              id={`answer-${question.id}`}
              data-testid={`textarea-answer-${index + 1}`}
              value={answers[question.id] ?? ""}
              onChange={(event) =>
                setAnswers((current) => ({
                  ...current,
                  [question.id]: event.target.value,
                }))
              }
              placeholder="اكتب إجابتك هنا…"
              rows={4}
            />
          </fieldset>
        ))}
        {error ? (
          <div className="error-banner" role="alert">
            <AlertTriangle size={16} />
            <span>{error}</span>
          </div>
        ) : null}
        <div className="submit-row">
          <span className="submit-note">
            <CircleHelp size={15} /> تحفظ الإجابات في لوحة الإدارة.
          </span>
          <button
            type="submit"
            data-testid="button-submit-response"
            className="button button-primary button-submit"
            disabled={
              isSubmitting ||
              !role.trim() ||
              form.questions.some((question) => !answers[question.id]?.trim())
            }
          >
            {isSubmitting ? "جارٍ الحفظ…" : "إرسال الإجابات"}{" "}
            <ArrowLeft size={17} />
          </button>
        </div>
      </form>
    </main>
  );
}

function AdminPage() {
  const [, setLocation] = useLocation();
  const [authenticated, setAuthenticated] = useState<boolean>();
  const [password, setPassword] = useState("");
  const [loginError, setLoginError] = useState("");
  const [loginLoading, setLoginLoading] = useState(false);
  const [title, setTitle] = useState("شاركنا لهجتك النجدية");
  const [headers, setHeaders] = useState<string[]>(DEFAULT_HEADERS);
  const [questions, setQuestions] = useState<Question[]>([]);
  const [fileName, setFileName] = useState<string>();
  const [responses, setResponses] = useState<SubmittedResponse[]>([]);
  const [error, setError] = useState("");
  const [responseError, setResponseError] = useState("");
  const [isLoading, setIsLoading] = useState(true);
  const [isPublishing, setIsPublishing] = useState(false);
  const [published, setPublished] = useState(false);
  const [responsesLoading, setResponsesLoading] = useState(false);

  const loadResponses = async () => {
    setResponsesLoading(true);
    setResponseError("");
    try {
      const response = await fetch("/api/admin/responses/default", {
        credentials: "include",
      });
      if (!response.ok) throw new Error("Responses could not be loaded.");
      setResponses((await response.json()) as SubmittedResponse[]);
    } catch (loadError) {
      setResponseError(
        loadError instanceof Error
          ? loadError.message
          : "Responses could not be loaded.",
      );
    } finally {
      setResponsesLoading(false);
    }
  };

  const loadAdminData = async () => {
    setIsLoading(true);
    setError("");
    try {
      const response = await fetch("/api/admin/forms/default", {
        credentials: "include",
      });
      if (response.ok) {
        const form = (await response.json()) as PublicForm;
        setTitle(form.title);
        setHeaders(form.headers);
        setQuestions(form.questions);
        setFileName("Published question bank");
        setPublished(true);
      } else if (response.status === 404) {
        const imported = await readAttachedWorkbook();
        setTitle("شاركنا لهجتك النجدية");
        setHeaders(imported.headers.length ? imported.headers : DEFAULT_HEADERS);
        setQuestions(imported.questions);
        setFileName("robothaqafa_questions.xlsx");
        setPublished(false);
      } else {
        throw new Error("The admin form could not be loaded.");
      }
      await loadResponses();
    } catch (loadError) {
      setError(
        loadError instanceof Error
          ? loadError.message
          : "The admin form could not be loaded.",
      );
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    fetch("/api/admin/status", { credentials: "include" })
      .then((response) => response.json() as Promise<{ authenticated: boolean }>)
      .then((status) => {
        setAuthenticated(status.authenticated);
        if (status.authenticated) void loadAdminData();
        else setIsLoading(false);
      })
      .catch(() => {
        setAuthenticated(false);
        setIsLoading(false);
      });
  }, []);

  const login = async () => {
    setLoginLoading(true);
    setLoginError("");
    try {
      const response = await fetch("/api/admin/login", {
        method: "POST",
        credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ password }),
      });
      if (!response.ok) throw new Error("That password is not correct.");
      setAuthenticated(true);
      setPassword("");
      await loadAdminData();
    } catch (loginFailure) {
      setLoginError(
        loginFailure instanceof Error
          ? loginFailure.message
          : "That password is not correct.",
      );
    } finally {
      setLoginLoading(false);
    }
  };

  const handleFile = async (file: File) => {
    if (!file.name) return;
    setError("");
    try {
      const imported = await readUpload(file);
      if (imported.questions.length < 5)
        throw new Error("Add at least five usable scenarios.");
      setQuestions(imported.questions);
      setHeaders(imported.headers.length ? imported.headers : DEFAULT_HEADERS);
      setFileName(file.name);
      setPublished(false);
    } catch (fileError) {
      setError(
        fileError instanceof Error
          ? fileError.message
          : "The question bank could not be read.",
      );
    }
  };

  const loadAttached = async () => {
    try {
      const imported = await readAttachedWorkbook();
      setQuestions(imported.questions);
      setHeaders(imported.headers.length ? imported.headers : DEFAULT_HEADERS);
      setFileName("robothaqafa_questions.xlsx");
      setPublished(false);
      setError("");
    } catch (loadError) {
      setError(
        loadError instanceof Error
          ? loadError.message
          : "The attached workbook could not be loaded.",
      );
    }
  };

  const publish = async () => {
    if (questions.length < 5) return;
    setIsPublishing(true);
    setError("");
    try {
      const response = await fetch("/api/admin/forms/default", {
        method: "PUT",
        credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ title, headers, questions }),
      });
      if (!response.ok) {
        const payload = (await response.json().catch(() => null)) as
          | { error?: string }
          | null;
        throw new Error(payload?.error || "The form could not be published.");
      }
      setPublished(true);
      await loadResponses();
    } catch (publishError) {
      setError(
        publishError instanceof Error
          ? publishError.message
          : "The form could not be published.",
      );
    } finally {
      setIsPublishing(false);
    }
  };

  const logout = async () => {
    await fetch("/api/admin/logout", {
      method: "POST",
      credentials: "include",
    });
    setAuthenticated(false);
  };

  if (authenticated === undefined || (authenticated && isLoading)) {
    return (
      <main className="auth-page">
        <div className="admin-empty">
          <RefreshCw size={20} className="spin" /> Loading admin workspace…
        </div>
      </main>
    );
  }
  if (!authenticated) {
    return (
      <AdminLoginView
        password={password}
        setPassword={setPassword}
        onSubmit={() => void login()}
        error={loginError}
        isLoading={loginLoading}
      />
    );
  }

  return (
    <div className="app-shell">
      <Sidebar
        view="build"
        onNavigate={(nextView) => {
          if (nextView === "respond") setLocation("/form");
        }}
        bankCount={questions.length}
      />
      <Topbar view="build" onNavigate={() => setLocation("/form")} />
      <div className="app-main">
        <div className="admin-toolbar">
          <span className="admin-session">
            <LockKeyhole size={14} /> Admin dashboard
          </span>
          <button
            type="button"
            className="text-button"
            data-testid="button-admin-logout"
            onClick={() => void logout()}
          >
            <LogOut size={14} /> Sign out
          </button>
        </div>
        <BuildView
          title={title}
          setTitle={setTitle}
          questions={questions}
          fileName={fileName}
          error={error}
          isLoading={false}
          onFile={(file) => void handleFile(file)}
          onSample={() => {
            setQuestions(sampleQuestions);
            setHeaders(DEFAULT_HEADERS);
            setFileName(undefined);
            setPublished(false);
          }}
          onLoadAttached={() => void loadAttached()}
          onClear={() => {
            setQuestions([]);
            setHeaders(DEFAULT_HEADERS);
            setFileName(undefined);
            setPublished(false);
          }}
          onNavigate={() => setLocation("/form")}
          onPublish={() => void publish()}
          isPublishing={isPublishing}
          published={published}
          shareLink={`${window.location.origin}/form`}
        />
        <AdminResponsesView
          responses={responses}
          headers={headers}
          title={title}
          isLoading={responsesLoading}
          error={responseError}
          onRefresh={() => void loadResponses()}
          onDownload={() => downloadAllResponses(headers, responses, title)}
        />
      </div>
    </div>
  );
}

function Home() {
  const [view, setView] = useState<View>("build");
  const [title, setTitle] = useState("شاركنا لهجتك النجدية");
  const [headers, setHeaders] = useState<string[]>(DEFAULT_HEADERS);
  const [questions, setQuestions] = useState<Question[]>(sampleQuestions);
  const [fileName, setFileName] = useState<string>();
  const [selectionHistory, setSelectionHistory] = useState<
    Record<string, number>
  >({});
  const [error, setError] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [respondentQuestions, setRespondentQuestions] = useState<Question[]>(
    [],
  );
  const [respondentRole, setRespondentRole] = useState("");
  const [answers, setAnswers] = useState<Record<string, string>>({});

  useEffect(() => {
    try {
      const stored = localStorage.getItem(STORAGE_KEY);
      if (stored) {
        const parsed = JSON.parse(stored) as StoredBank;
        if (parsed.questions?.length) {
          setQuestions(parsed.questions);
          setHeaders(parsed.headers?.length ? parsed.headers : DEFAULT_HEADERS);
          setTitle(parsed.title || "شاركنا لهجتك النجدية");
          setFileName(parsed.fileName);
          setSelectionHistory(parsed.selectionHistory ?? {});
        }
      }
    } catch {
      /* local storage can be unavailable in private browsing */
    }
  }, []);
  useEffect(() => {
    try {
      localStorage.setItem(
        STORAGE_KEY,
        JSON.stringify({
          title,
          headers,
          questions,
          fileName,
          selectionHistory,
        } satisfies StoredBank),
      );
    } catch {
      /* keep the experience usable without storage */
    }
  }, [title, headers, questions, fileName, selectionHistory]);

  useEffect(() => {
    if (fileName) return;
    let cancelled = false;
    const loadAttachedBank = async () => {
      try {
        const imported = await readAttachedWorkbook();
        if (!cancelled && imported.questions.length >= 5) {
          setQuestions(imported.questions);
          setHeaders(
            imported.headers.length ? imported.headers : DEFAULT_HEADERS,
          );
          setFileName("robothaqafa_questions.xlsx");
        }
      } catch {
        // The sample bank remains available if the bundled workbook is unavailable.
      }
    };
    void loadAttachedBank();
    return () => {
      cancelled = true;
    };
  }, [fileName]);

  const makeRespondentSet = () => {
    if (questions.length < 5) return;
    const selected = shuffle(questions)
      .sort(
        (left, right) =>
          (selectionHistory[left.id] ?? 0) - (selectionHistory[right.id] ?? 0),
      )
      .slice(0, 5);
    setRespondentQuestions(selected);
    setSelectionHistory((current) =>
      selected.reduce(
        (next, question) => ({
          ...next,
          [question.id]: (next[question.id] ?? 0) + 1,
        }),
        { ...current },
      ),
    );
    setAnswers({});
    setView("respond");
  };
  const handleFile = async (file: File) => {
    if (!file.name) {
      setError("");
      return;
    }
    setError("");
    setIsLoading(true);
    try {
      const imported = await readUpload(file);
      if (imported.questions.length < 5)
        throw new Error(
          `We found ${imported.questions.length} usable question${imported.questions.length === 1 ? "" : "s"}. Add at least five to create randomized sets.`,
        );
      setQuestions(imported.questions);
      setHeaders(imported.headers.length ? imported.headers : DEFAULT_HEADERS);
      setFileName(file.name);
      setSelectionHistory({});
    } catch (uploadError) {
      setError(
        uploadError instanceof Error
          ? uploadError.message
          : "We could not read that file. Try a CSV with one question per row.",
      );
    } finally {
      setIsLoading(false);
    }
  };
  const loadAttached = async () => {
    setError("");
    setIsLoading(true);
    try {
      const imported = await readAttachedWorkbook();
      if (imported.questions.length < 5)
        throw new Error(
          "The attached workbook has fewer than five usable scenarios.",
        );
      setQuestions(imported.questions);
      setHeaders(imported.headers.length ? imported.headers : DEFAULT_HEADERS);
      setFileName("robothaqafa_questions.xlsx");
      setSelectionHistory({});
    } catch (loadError) {
      setError(
        loadError instanceof Error
          ? loadError.message
          : "The attached workbook could not be loaded.",
      );
    } finally {
      setIsLoading(false);
    }
  };
  const loadSample = () => {
    setQuestions(sampleQuestions);
    setHeaders(DEFAULT_HEADERS);
    setFileName(undefined);
    setSelectionHistory({});
    setError("");
  };
  const clearBank = () => {
    setQuestions([]);
    setHeaders(DEFAULT_HEADERS);
    setFileName(undefined);
    setSelectionHistory({});
    setError("");
  };
  const navigate = (nextView: View) => {
    if (nextView === "respond") makeRespondentSet();
    else setView(nextView);
  };
  const handleSubmit = (nextAnswers: Record<string, string>) => {
    setAnswers(nextAnswers);
    setView("results");
  };
  const handleDownload = () =>
    downloadResponses(
      headers,
      respondentQuestions,
      respondentRole,
      answers,
      title,
    );
  const shellView = view === "results" ? "results" : view;
  return (
    <div className="app-shell">
      <Sidebar
        view={shellView}
        onNavigate={navigate}
        bankCount={questions.length}
      />
      <Topbar view={view} onNavigate={navigate} />
      <div className="app-main">
        {view === "build" ? (
          <BuildView
            title={title}
            setTitle={setTitle}
            questions={questions}
            fileName={fileName}
            error={error}
            isLoading={isLoading}
            onFile={handleFile}
            onSample={loadSample}
            onLoadAttached={loadAttached}
            onClear={clearBank}
            onNavigate={navigate}
          />
        ) : view === "respond" ? (
          <RespondView
            title={title}
            role={respondentRole}
            setRole={setRespondentRole}
            questions={respondentQuestions}
            onBack={() => setView("build")}
            onSubmit={handleSubmit}
          />
        ) : (
          <ResultsView
            title={title}
            role={respondentRole}
            questions={respondentQuestions}
            answers={answers}
            onAgain={makeRespondentSet}
            onBuild={() => setView("build")}
            onDownload={handleDownload}
          />
        )}
      </div>
    </div>
  );
}

function Router() {
  return (
    <RoutedErrorBoundary>
      <Switch>
        <Route path="/" component={PublicFormPage} />
        <Route path="/form" component={PublicFormPage} />
        <Route path="/admin" component={AdminPage} />
        <Route component={NotFound} />
      </Switch>
    </RoutedErrorBoundary>
  );
}

function RoutedErrorBoundary({ children }: { children: ReactNode }) {
  const [location] = useLocation();
  return <ErrorBoundary resetKey={location}>{children}</ErrorBoundary>;
}

function App() {
  return (
    <QueryClientProvider client={queryClient}>
      <TooltipProvider>
        <WouterRouter base={import.meta.env.BASE_URL.replace(/\/$/, "")}>
          <Router />
        </WouterRouter>
        <Toaster />
      </TooltipProvider>
    </QueryClientProvider>
  );
}

export default App;
