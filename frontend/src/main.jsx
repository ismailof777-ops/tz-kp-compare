import React, { useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import {
  ArrowRight,
  BarChart3,
  Check,
  CheckCircle2,
  Clock3,
  Database,
  Download,
  FileCheck2,
  FileSpreadsheet,
  FileText,
  LockKeyhole,
  LogIn,
  Mail,
  RotateCcw,
  ShieldCheck,
  Trash2,
  Upload,
  X,
} from "lucide-react";
import "./styles.css";

const fileSize = (size) => {
  if (size >= 1024 * 1024) return `${(size / 1024 / 1024).toFixed(1).replace(".0", "")} МБ`;
  if (size >= 1024) return `${Math.round(size / 1024)} КБ`;
  return `${size} Б`;
};

const fileExt = (name) => {
  const ext = String(name || "").split(".").pop();
  return ext ? ext.toUpperCase() : "ФАЙЛ";
};

const currentPath = () => window.location.pathname;

function App() {
  const path = currentPath();

  if (path === "/login") return <AuthLayout><LoginPage /></AuthLayout>;
  if (path === "/forgot-password") return <AuthLayout><ForgotPasswordPage /></AuthLayout>;
  if (path === "/reset-password") return <AuthLayout><ResetPasswordPage /></AuthLayout>;
  if (path.startsWith("/done/")) return <AppLayout><DonePage runId={decodeURIComponent(path.replace("/done/", ""))} /></AppLayout>;

  return <AppLayout><UploadPage /></AppLayout>;
}

function AppLayout({ children }) {
  return (
    <div className="app">
      <Header />
      <main className="shell">{children}</main>
      <Footer />
      <CookieBanner />
    </div>
  );
}

function AuthLayout({ children }) {
  return (
    <div className="auth-page">
      <Header compact />
      <main className="auth-shell">{children}</main>
      <Footer />
      <CookieBanner />
    </div>
  );
}

function Header({ compact = false }) {
  return (
    <header className={`site-header ${compact ? "is-compact" : ""}`}>
      <a className="brand" href="/">
        <span className="brand-mark"><BarChart3 size={18} /></span>
        <span>Сравнение КП</span>
      </a>
      <nav className="header-links" aria-label="Навигация">
        <a href="/">Сервис</a>
        <a href="/admin">Админка</a>
        <a href="/#how-it-works">Помощь</a>
        <a href="/privacy">Политика</a>
        {!compact && <a href="/logout">Выход</a>}
      </nav>
    </header>
  );
}

function Footer() {
  return (
    <footer className="site-footer">
      <span>© approvemoscow.ru</span>
      <a href="/privacy">Политика обработки персональных данных</a>
    </footer>
  );
}

function CookieBanner() {
  const [visible, setVisible] = useState(() => {
    try {
      return !window.localStorage.getItem("approvemoscow_cookie_choice");
    } catch {
      return true;
    }
  });

  const choose = (value) => {
    try {
      window.localStorage.setItem("approvemoscow_cookie_choice", value);
    } catch {
      // localStorage may be unavailable in strict browser modes.
    }
    setVisible(false);
  };

  if (!visible) return null;
  return (
    <section className="cookie-banner" aria-label="Уведомление о cookies">
      <div className="cookie-text">
        <b>Cookies и технические данные</b>
        <span>Сервис использует только необходимые технические данные для работы сайта. Аналитика и рекламные cookies не подключены.</span>
        <a href="/privacy">Подробнее</a>
      </div>
      <div className="cookie-actions">
        <button className="btn secondary" type="button" onClick={() => choose("necessary")}>Только необходимые</button>
        <button className="btn" type="button" onClick={() => choose("accepted")}>Понятно</button>
      </div>
    </section>
  );
}

function LoginPage() {
  const params = new URLSearchParams(window.location.search);
  const next = params.get("next") || "/";
  return (
    <section className="auth-card">
      <span className="eyebrow"><LockKeyhole size={15} /> Закрытый доступ</span>
      <h1>Вход в сервис</h1>
      <p>Введите логин и пароль администратора, чтобы работать с загрузкой КП, историей и базой закупок.</p>
      <form className="auth-form" action="/login" method="post">
        <input type="hidden" name="next" value={next} />
        <label>Логин<input className="input" name="username" autoComplete="username" required /></label>
        <label>Пароль<input className="input" type="password" name="password" autoComplete="current-password" required /></label>
        <button className="btn primary-wide" type="submit"><LogIn size={18} /> Войти</button>
        <a className="text-link" href="/forgot-password">Забыли пароль?</a>
      </form>
    </section>
  );
}

function ForgotPasswordPage() {
  const [message, setMessage] = useState("");
  const [isSubmitting, setIsSubmitting] = useState(false);

  const submit = async (event) => {
    event.preventDefault();
    setIsSubmitting(true);
    const form = new FormData(event.currentTarget);
    await fetch("/forgot-password", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams(form),
    }).catch(() => null);
    setMessage("Если email совпадает с администраторским, ссылка для сброса пароля будет отправлена.");
    setIsSubmitting(false);
  };

  return (
    <section className="auth-card">
      <span className="eyebrow"><Mail size={15} /> Восстановление</span>
      <h1>Сброс пароля</h1>
      <p>Введите email администратора. Если он совпадает с настройками сервиса, на него будет отправлена ссылка для смены пароля.</p>
      {message && <div className="notice">{message}</div>}
      <form className="auth-form" onSubmit={submit}>
        <label>Email администратора<input className="input" type="email" name="email" autoComplete="email" required /></label>
        <button className="btn primary-wide" type="submit" disabled={isSubmitting}><Mail size={18} /> Отправить ссылку</button>
        <a className="text-link" href="/login">Вернуться ко входу</a>
      </form>
    </section>
  );
}

function ResetPasswordPage() {
  const token = new URLSearchParams(window.location.search).get("token") || "";
  return (
    <section className="auth-card">
      <span className="eyebrow"><RotateCcw size={15} /> Восстановление</span>
      <h1>Новый пароль</h1>
      <p>Ссылка действует 30 минут и может быть использована только один раз.</p>
      {!token && <div className="notice error">Ссылка недействительна или срок действия истек.</div>}
      {token ? (
        <form className="auth-form" action="/reset-password" method="post">
          <input type="hidden" name="token" value={token} />
          <label>Новый пароль<input className="input" type="password" name="new_password" autoComplete="new-password" minLength={8} required /></label>
          <label>Повторите новый пароль<input className="input" type="password" name="confirm_password" autoComplete="new-password" minLength={8} required /></label>
          <button className="btn primary-wide" type="submit"><CheckCircle2 size={18} /> Сохранить новый пароль</button>
        </form>
      ) : (
        <a className="btn secondary" href="/forgot-password">Запросить новую ссылку</a>
      )}
    </section>
  );
}

function UploadPage() {
  const [requestFiles, setRequestFiles] = useState([]);
  const [offerFiles, setOfferFiles] = useState([]);
  const [consent, setConsent] = useState(false);
  const [message, setMessage] = useState("");
  const [progress, setProgress] = useState(null);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const controllerRef = useRef(null);

  const submit = async (event) => {
    event.preventDefault();
    if (!requestFiles.length) return setMessage("Выберите файл заявки / ТЗ.");
    if (!offerFiles.length) return setMessage("Выберите КП или счет поставщика.");
    if (!consent) return setMessage("Подтвердите согласие на обработку персональных данных.");

    setMessage("");
    setIsSubmitting(true);
    setProgress({ percent: 3, message: "Загружаем файлы и готовим обработку." });
    const controller = new AbortController();
    controllerRef.current = controller;

    try {
      const form = new FormData();
      form.append("privacy_consent", "yes");
      form.append("request", requestFiles[0]);
      offerFiles.forEach((file) => form.append("offers", file));
      const response = await fetch("/process", {
        method: "POST",
        headers: { Accept: "application/json" },
        body: form,
        signal: controller.signal,
      });
      const data = await response.json();
      if (!response.ok || data.state === "error") throw new Error(data.message || "Не удалось обработать файлы.");
      pollProgress(data.run_id);
    } catch (error) {
      if (error.name !== "AbortError") {
        setMessage(error.message || "Не удалось обработать файлы. Проверьте подключение и попробуйте еще раз.");
      }
      setIsSubmitting(false);
      setProgress(null);
      controllerRef.current = null;
    }
  };

  const pollProgress = async (runId) => {
    try {
      const response = await fetch(`/progress/${encodeURIComponent(runId)}`, { cache: "no-store" });
      const data = await response.json();
      setProgress({ percent: Math.max(0, Math.min(100, Number(data.percent || 0))), message: data.message || "Обработка файлов" });
      if (data.state === "done") {
        window.location.href = data.redirect || `/review/${encodeURIComponent(runId)}`;
        return;
      }
      if (data.state === "error") throw new Error(data.message || "Не удалось обработать файлы.");
      window.setTimeout(() => pollProgress(runId), 900);
    } catch (error) {
      setMessage(error.message || "Не удалось получить статус обработки. Обновите страницу или попробуйте заново.");
      setIsSubmitting(false);
      setProgress(null);
      controllerRef.current = null;
    }
  };

  const stop = () => {
    controllerRef.current?.abort();
    setIsSubmitting(false);
    setProgress(null);
    setMessage("Обработка остановлена. Можно изменить файлы и запустить заново.");
  };

  return (
    <>
      <section className="page-hero">
        <div>
          <span className="eyebrow"><ShieldCheck size={15} /> XLSX · PDF · Excel-отчет</span>
          <h1>Сравнение ТЗ и КП поставщиков</h1>
          <p>Загрузите заявку и коммерческие предложения. Сервис сопоставит позиции, найдет минимальные цены и сформирует Excel-сводку для проверки закупки.</p>
        </div>
      </section>
      <section className="main-grid">
        <InfoPanel />
        <form className="work-panel" onSubmit={submit}>
          <Stepper active={isSubmitting ? 3 : offerFiles.length ? 2 : requestFiles.length ? 1 : 0} />
          <UploadField
            title="Заявка / ТЗ"
            description="Загрузите Excel-файл со списком позиций, количеством и единицами измерения."
            accept=".xlsx,.xls,.pdf"
            files={requestFiles}
            setFiles={(files) => setRequestFiles(files.slice(0, 1))}
            multiple={false}
            emptyTitle="Перетащите файл заявки сюда"
            emptyText="или выберите .xlsx/.xls/PDF на компьютере"
            support="Поддерживается .xlsx. Для сравнения двух КП можно загрузить сюда первый КП .xlsx/.xls/PDF."
            templateHref="/template/request.xlsx"
            templateText="Скачать шаблон заявки"
          />
          <UploadField
            title="КП / счета поставщиков"
            description="Добавьте один или несколько КП или счетов от поставщиков. Для сравнения минимальных цен загрузите два и более файла."
            accept=".xlsx,.xls,.pdf"
            files={offerFiles}
            setFiles={setOfferFiles}
            multiple
            emptyTitle="Перетащите КП или счета сюда"
            emptyText="или выберите один или несколько файлов"
            support="Поддерживаются .xlsx, .xls и текстовые PDF"
            templateHref="/template/offer.xlsx"
            templateText="Скачать шаблон КП"
          />
          <label className="consent">
            <input type="checkbox" checked={consent} onChange={(event) => setConsent(event.target.checked)} />
            <span>Я согласен с <a href="/privacy" target="_blank" rel="noreferrer">обработкой персональных данных</a></span>
          </label>
          <div className="action-row">
            <button className="btn primary-wide" type="submit" disabled={isSubmitting}><FileSpreadsheet size={19} /> Сформировать Excel-отчет</button>
            {isSubmitting && <button className="btn danger" type="button" onClick={stop}><X size={18} /> Остановить обработку</button>}
          </div>
          <p className="caption">Обычно обработка занимает 1–3 минуты.</p>
          {message && <div className="notice error">{message}</div>}
          {progress && <ProcessingState progress={progress} />}
        </form>
      </section>
    </>
  );
}

function InfoPanel() {
  return (
    <aside className="info-panel" id="how-it-works">
      <h2>Как это работает</h2>
      <div className="steps-list">
        <Step number="1" title="Загрузите заявку" text="Excel-файл со списком позиций, количеством и единицами измерения." />
        <Step number="2" title="Добавьте КП поставщиков" text="Можно загрузить несколько Excel/PDF файлов от разных поставщиков." />
        <Step number="3" title="Получите Excel-отчет" text="Сервис выделит совпадения, минимальные цены и спорные строки." />
      </div>
      <div className="report-box">
        <h3>В отчете будет</h3>
        <div className="pill-grid">
          {["позиции", "поставщики", "цены за единицу", "итоговые суммы", "минимальные предложения", "спорные совпадения"].map((item) => (
            <span key={item}><Check size={14} /> {item}</span>
          ))}
        </div>
      </div>
    </aside>
  );
}

function Step({ number, title, text }) {
  return (
    <div className="info-step">
      <span>{number}</span>
      <div><b>{title}</b><p>{text}</p></div>
    </div>
  );
}

function Stepper({ active }) {
  const steps = ["Заявка", "КП поставщиков", "Excel-отчет"];
  return (
    <div className="stepper">
      {steps.map((step, index) => (
        <div key={step} className={`step ${index <= active ? "is-active" : ""}`}>
          <span>{index + 1}</span>
          <b>{step}</b>
        </div>
      ))}
    </div>
  );
}

function UploadField({ title, description, accept, files, setFiles, multiple, emptyTitle, emptyText, support, templateHref, templateText }) {
  const inputRef = useRef(null);
  const [dragging, setDragging] = useState(false);

  const addFiles = (fileList) => {
    const next = Array.from(fileList || []);
    if (!multiple) return setFiles(next.slice(0, 1));
    const seen = new Set(files.map((file) => `${file.name}|${file.size}|${file.lastModified}`));
    const merged = [...files];
    next.forEach((file) => {
      const key = `${file.name}|${file.size}|${file.lastModified}`;
      if (!seen.has(key)) merged.push(file);
    });
    setFiles(merged);
  };

  const remove = (index) => setFiles(files.filter((_, itemIndex) => itemIndex !== index));

  return (
    <section className="upload-section">
      <div className="section-head">
        <div><h2>{title}</h2><p>{description}</p></div>
      </div>
      <div
        className={`dropzone ${dragging ? "is-dragging" : ""}`}
        onClick={() => inputRef.current?.click()}
        onDragEnter={(event) => { event.preventDefault(); setDragging(true); }}
        onDragOver={(event) => event.preventDefault()}
        onDragLeave={(event) => { event.preventDefault(); setDragging(false); }}
        onDrop={(event) => { event.preventDefault(); setDragging(false); addFiles(event.dataTransfer.files); }}
      >
        <input ref={inputRef} type="file" accept={accept} multiple={multiple} onChange={(event) => addFiles(event.target.files)} />
        <span className="drop-icon"><Upload size={22} /></span>
        <b>{emptyTitle}</b>
        <p>{emptyText}</p>
        <button className="btn secondary" type="button">Выбрать файл{multiple ? "ы" : ""}</button>
        <small>{support}</small>
      </div>
      <a className="template-link" href={templateHref}><Download size={15} /> {templateText}</a>
      <FileList files={files} multiple={multiple} remove={remove} />
    </section>
  );
}

function FileList({ files, multiple, remove }) {
  if (!files.length) return <div className="files-list empty">{multiple ? "Файлы пока не выбраны." : "Файл пока не выбран."}</div>;
  return (
    <div className="files-list is-filled">
      <div className="file-summary">{multiple ? `Загружено файлов: ${files.length}` : "Файл выбран"}</div>
      <ul>
        {files.map((file, index) => (
          <li key={`${file.name}-${file.size}-${file.lastModified}`}>
            <div className="file-main">
              <FileText size={18} />
              <div>
                <b title={file.name}>{file.name}</b>
                <span>{fileExt(file.name)} · {fileSize(file.size)} · <em>Готово</em></span>
              </div>
            </div>
            <button className="icon-btn" type="button" aria-label="Удалить файл" onClick={() => remove(index)}><Trash2 size={16} /></button>
          </li>
        ))}
      </ul>
    </div>
  );
}

function ProcessingState({ progress }) {
  return (
    <section className="processing-card">
      <div className="processing-head">
        <div><Clock3 size={18} /><b>Формирование Excel-отчета</b></div>
        <strong>{progress.percent}%</strong>
      </div>
      <div className="progress"><span style={{ width: `${progress.percent}%` }} /></div>
      <p>{progress.message}</p>
      <small>Сервис извлекает позиции, цены и сроки, затем сопоставляет товары с учетом ИИ. Если файлов много или PDF большой, обработка может занять несколько минут.</small>
    </section>
  );
}

function DonePage({ runId }) {
  const safeRunId = useMemo(() => encodeURIComponent(runId), [runId]);
  return (
    <section className="done-panel">
      <span className="ready-badge"><CheckCircle2 size={16} /> Готово</span>
      <h1>Excel-отчет готов</h1>
      <p>Ручные правки применены. Итоговая сводка сформирована и доступна для скачивания.</p>
      <div className="done-actions">
        <a className="btn primary-wide" href={`/download/${safeRunId}/summary.xlsx`} download><Download size={19} /> Скачать Excel-отчет</a>
        <a className="btn secondary" href={`/download/${safeRunId}/review.xlsx`} download><FileCheck2 size={18} /> Скачать файл проверки</a>
        <a className="btn secondary" href="/"><ArrowRight size={18} /> Новая обработка</a>
      </div>
    </section>
  );
}

createRoot(document.getElementById("root")).render(<App />);
