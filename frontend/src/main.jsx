import React, { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import {
  Activity,
  Archive,
  ArrowRight,
  CheckCircle2,
  Clock3,
  Database,
  Download,
  FileCheck2,
  FileSpreadsheet,
  FileText,
  History,
  Home,
  KeyRound,
  LockKeyhole,
  LogIn,
  Mail,
  RefreshCw,
  RotateCcw,
  Search,
  Settings,
  Trash2,
  Upload,
  X,
} from "lucide-react";
import "./styles.css";

const NAV_ITEMS = [
  { href: "/", label: "Загрузка файлов", icon: Home },
  { href: "/admin", label: "История обработок", icon: History },
  { href: "/admin/purchases", label: "База закупок", icon: Database },
  { href: "/admin/audit", label: "Журнал действий", icon: Activity },
];

const STATUS_LABELS = {
  done: "готово",
  finalized: "готово",
  review: "на проверке",
  auto: "авто",
  manual: "ручное",
  unmatched: "не найдено",
  service: "услуга",
  error: "ошибка",
  running: "в работе",
  pending: "ожидает",
};

const fileSize = (size) => {
  if (size >= 1024 * 1024) return `${(size / 1024 / 1024).toFixed(1).replace(".0", "")} МБ`;
  if (size >= 1024) return `${Math.round(size / 1024)} КБ`;
  return `${size} Б`;
};

const fileExt = (name) => {
  const ext = String(name || "").split(".").pop();
  return ext ? ext.toUpperCase() : "ФАЙЛ";
};

const formatDateTime = (value) => {
  if (!value) return "—";
  const normalized = String(value).endsWith("Z") ? value : `${value}Z`;
  const date = new Date(normalized);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleString("ru-RU", {
    day: "2-digit",
    month: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  });
};

async function fetchJson(url, options = {}) {
  const response = await fetch(url, {
    ...options,
    headers: {
      Accept: "application/json",
      ...(options.headers || {}),
    },
  });
  if (response.status === 401) {
    window.location.href = `/login?next=${encodeURIComponent(window.location.pathname)}`;
    throw new Error("Требуется вход в сервис.");
  }
  const data = await response.json().catch(() => ({}));
  if (!response.ok || data.ok === false || data.state === "error") {
    throw new Error(data.message || "Не удалось выполнить запрос.");
  }
  return data;
}

function App() {
  const path = window.location.pathname;

  if (path === "/privacy") return <PublicLayout><PrivacyPage /></PublicLayout>;
  if (path === "/login") return <AuthLayout><LoginPage /></AuthLayout>;
  if (path === "/forgot-password") return <AuthLayout><ForgotPasswordPage /></AuthLayout>;
  if (path === "/reset-password") return <AuthLayout><ResetPasswordPage /></AuthLayout>;

  if (path.startsWith("/done/")) {
    return <AppShell><DonePage runId={decodeURIComponent(path.replace("/done/", ""))} /></AppShell>;
  }
  if (path.startsWith("/review/")) {
    return <AppShell><ReviewPage runId={decodeURIComponent(path.replace("/review/", ""))} /></AppShell>;
  }
  if (path === "/admin" || path === "/admin/history") return <AppShell><AdminHistoryPage /></AppShell>;
  if (path === "/admin/purchases") return <AppShell><PurchasesPage /></AppShell>;
  if (path === "/admin/audit") return <AppShell><AuditPage /></AppShell>;
  if (path === "/admin/security") return <AppShell><SecurityPage /></AppShell>;

  return <AppShell><UploadPage /></AppShell>;
}

function AppShell({ children }) {
  return (
    <div className="app-shell">
      <Sidebar />
      <main className="app-main">
        {children}
        <Footer />
      </main>
      <CookieBanner />
    </div>
  );
}

function Sidebar() {
  const path = window.location.pathname;
  return (
    <aside className="sidebar">
      <a className="side-brand" href="/">
        <LogoMark />
        <strong>Approve Moscow</strong>
      </a>

      <nav className="side-nav" aria-label="Основная навигация">
        {NAV_ITEMS.map((item) => {
          const Icon = item.icon;
          const active = item.href === "/" ? path === "/" : path === item.href || (item.href === "/admin" && path === "/admin/history");
          return (
            <a key={item.href} className={active ? "is-active" : ""} href={item.href}>
              <Icon size={17} />
              <span>{item.label}</span>
            </a>
          );
        })}
      </nav>

      <div className="side-bottom">
        <a className={path === "/admin/security" ? "is-active" : ""} href="/admin/security">
          <Settings size={18} />
          <span>Настройки</span>
        </a>
        <div className="side-links">
          <a href="/privacy">Политика</a>
          <a href="/logout">Выход</a>
        </div>
      </div>
    </aside>
  );
}

function AuthLayout({ children }) {
  return (
    <div className="auth-page">
      <header className="auth-header">
        <a className="auth-brand" href="/" aria-label="Approve Moscow">
          <img className="auth-brand-logo" src="/brand/logo-full.png" alt="Approve Moscow" />
        </a>
      </header>
      <main className="auth-shell">
        {children}
      </main>
      <Footer />
      <CookieBanner />
    </div>
  );
}

function PublicLayout({ children }) {
  return (
    <div className="auth-page">
      <header className="auth-header">
        <a className="auth-brand" href="/" aria-label="Approve Moscow">
          <img className="auth-brand-logo" src="/brand/logo-full.png" alt="Approve Moscow" />
        </a>
      </header>
      <main className="public-shell">
        {children}
      </main>
      <Footer />
      <CookieBanner />
    </div>
  );
}

function LogoMark() {
  return (
    <span className="brand-mark" aria-hidden="true">
      <img src="/brand/logo-mark.png" alt="" />
    </span>
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

function PrivacyPage() {
  return (
    <>
      <section className="privacy-hero">
        <span className="soft-badge"><FileText size={15} /> Юридическая информация</span>
        <h1>Политика обработки персональных данных</h1>
        <p>Документ описывает, какие данные обрабатывает сервис approvemoscow.ru при сравнении заявки/ТЗ и КП поставщиков.</p>
      </section>

      <section className="panel legal-panel">
        <div className="legal-meta">
          <span>Сервис</span>
          <b>Approve Moscow</b>
        </div>

        <LegalSection title="1. Общие положения">
          <p>Настоящая политика применяется к сайту approvemoscow.ru и сервису сравнения заявки/ТЗ с коммерческими предложениями поставщиков. Оператор персональных данных: Меренов Кирилл Олегович, ИНН 645120364787, email для обращений: <a href="mailto:Merenov.kirill@mail.ru">Merenov.kirill@mail.ru</a>.</p>
          <p>Используя сервис и загружая файлы, пользователь подтверждает согласие с условиями обработки данных.</p>
        </LegalSection>

        <LegalSection title="2. Какие данные могут обрабатываться">
          <ul>
            <li>сведения из загружаемых файлов заявки, ТЗ, КП, счетов и PDF-документов;</li>
            <li>контактные данные, реквизиты организаций, ФИО, телефоны, адреса электронной почты, если они содержатся в файлах;</li>
            <li>данные авторизации администратора: логин, технический идентификатор сессии и время входа;</li>
            <li>технические данные запроса: IP-адрес, дата и время обращения, сведения о браузере, session-cookie и системные журналы сервера.</li>
          </ul>
        </LegalSection>

        <LegalSection title="3. Цели обработки">
          <p>Данные используются для загрузки и распознавания документов, сопоставления позиций заявки и КП, формирования Excel-отчета, диагностики ошибок и обеспечения работоспособности сервиса.</p>
        </LegalSection>

        <LegalSection title="4. Правовые основания и согласие">
          <p>Обработка выполняется на основании согласия пользователя, выраженного отметкой чекбокса перед отправкой файлов, а также для исполнения действия, запрошенного пользователем в сервисе.</p>
        </LegalSection>

        <LegalSection title="5. Передача третьим лицам">
          <p>Для автоматического сопоставления позиций сервис может использовать внешнего поставщика AI-инфраструктуры, включая DeepSeek/API обработки текста при включенном ИИ. В такие API могут передаваться фрагменты данных из загруженных документов только в объеме, необходимом для сравнения позиций и формирования отчета.</p>
          <p>Сервис не продает персональные данные и не передает их третьим лицам для рекламных целей.</p>
        </LegalSection>

        <LegalSection title="6. Хранение и защита">
          <p>Загруженные файлы, история обработок, база прошлых закупок, подтвержденные ручные сопоставления и сформированные отчеты хранятся на сервере сервиса для работы истории, админки, повторной проверки и самообучения сопоставлений. Срок хранения определяется оператором исходя из целей использования сервиса, требований закона и обращений пользователя.</p>
          <p>Доступ к рабочим страницам сервиса ограничивается логином и паролем. Сессия пользователя хранится в технической HttpOnly cookie, необходимой для авторизации и защиты закрытых разделов.</p>
          <p>В сервисе применяются меры разграничения прав доступа, учет действий администратора с данными, журналирование входов, импортов и обработок, а также технические меры защиты серверной инфраструктуры.</p>
        </LegalSection>

        <LegalSection title="7. Cookies и аналитика">
          <p>Сайт использует только необходимые технические cookies: session-cookie для авторизации и локальное сохранение выбора в уведомлении о cookies. Рекламная аналитика и cookies для отслеживания пользователей не подключены.</p>
        </LegalSection>

        <LegalSection title="8. Права пользователя">
          <p>Пользователь может запросить информацию об обработке данных, уточнение или удаление загруженных материалов и результатов обработки, если такие данные сохраняются на сервере.</p>
        </LegalSection>

        <LegalSection title="9. Контакты">
          <p>По вопросам обработки персональных данных, уточнения или удаления данных можно обратиться к оператору: Меренов Кирилл Олегович, email <a href="mailto:Merenov.kirill@mail.ru">Merenov.kirill@mail.ru</a>.</p>
          <p>Форма уведомления оператора персональных данных размещена на официальном портале Роскомнадзора: <a href="https://pd.rkn.gov.ru/operators-registry/notification/" target="_blank" rel="noopener">pd.rkn.gov.ru/operators-registry/notification/</a>.</p>
        </LegalSection>
      </section>
    </>
  );
}

function LegalSection({ title, children }) {
  return (
    <section className="legal-section">
      <h2>{title}</h2>
      {children}
    </section>
  );
}

function PageHeader({ title, text, badge, actions }) {
  return (
    <header className="page-header">
      <div>
        <h1>{title}</h1>
        {text && <p>{text}</p>}
        {badge && <span className="soft-badge">{badge}</span>}
      </div>
      {actions && <div className="page-actions">{actions}</div>}
    </header>
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

  const choose = () => {
    try {
      window.localStorage.setItem("approvemoscow_cookie_choice", "accepted");
    } catch {
      // localStorage may be blocked by browser settings.
    }
    setVisible(false);
  };

  if (!visible) return null;
  return (
    <section className="cookie-banner" aria-label="Уведомление о cookies">
      <div>
        <b>Технические cookies</b>
        <span>Используются только для входа и корректной работы сервиса. Рекламная аналитика не подключена.</span>
      </div>
      <button className="btn btn-primary" type="button" onClick={choose}>Понятно</button>
    </section>
  );
}

function LoginPage() {
  const params = new URLSearchParams(window.location.search);
  const next = params.get("next") || "/";
  const error = params.get("error");
  const message = params.get("message");
  return (
    <section className="auth-card">
      <span className="soft-badge"><LockKeyhole size={15} /> Закрытый доступ</span>
      <h1>Вход в сервис</h1>
      <p>Введите логин и пароль администратора, чтобы работать с загрузкой КП, историей и базой закупок.</p>
      {error === "invalid" && <div className="auth-error" role="alert">Неверный логин или пароль</div>}
      {message === "password_changed" && <Notice>Пароль изменен. Войдите с новым паролем.</Notice>}
      <form className="form-stack" action="/login" method="post">
        <input type="hidden" name="next" value={next} />
        <label>Логин<input className="input" name="username" autoComplete="username" required /></label>
        <label>Пароль<input className="input" type="password" name="password" autoComplete="current-password" required /></label>
        <button className="btn btn-primary btn-wide" type="submit"><LogIn size={18} /> Войти</button>
        <a className="text-link" href="/forgot-password">Забыли пароль?</a>
      </form>
      <div className="auth-service-note">
        <h2>Сервис для проверки закупок</h2>
        <p>Сравнивает заявку/ТЗ с КП поставщиков, помогает найти минимальные цены, спорные совпадения и сформировать Excel-сводку для закупки.</p>
      </div>
    </section>
  );
}

function ForgotPasswordPage() {
  const [notice, setNotice] = useState(null);
  const [isSubmitting, setIsSubmitting] = useState(false);

  const submit = async (event) => {
    event.preventDefault();
    setIsSubmitting(true);
    setNotice(null);
    const form = new FormData(event.currentTarget);
    try {
      const response = await fetch("/forgot-password", {
        method: "POST",
        headers: {
          Accept: "application/json",
          "Content-Type": "application/x-www-form-urlencoded",
        },
        body: new URLSearchParams(form),
      });
      const data = await response.json().catch(() => ({}));
      setNotice({
        tone: data.matched === false ? "error" : "success",
        text: data.message || "Запрос на сброс пароля отправлен.",
      });
    } catch {
      setNotice({ tone: "error", text: "Не удалось отправить запрос. Попробуйте еще раз." });
    } finally {
      setIsSubmitting(false);
    }
  };

  return (
    <section className="auth-card">
      <span className="soft-badge"><Mail size={15} /> Восстановление</span>
      <h1>Сброс пароля</h1>
      <p>Укажите email администратора. Другие адреса не смогут запросить смену пароля.</p>
      {notice && <div className={notice.tone === "error" ? "auth-error" : "auth-success"} role="alert">{notice.text}</div>}
      <form className="form-stack" onSubmit={submit}>
        <label>Email администратора<input className="input" type="email" name="email" autoComplete="email" required /></label>
        <button className="btn btn-primary btn-wide" type="submit" disabled={isSubmitting}><Mail size={18} /> Отправить ссылку</button>
        <a className="text-link" href="/login">Вернуться ко входу</a>
      </form>
    </section>
  );
}

function ResetPasswordPage() {
  const token = new URLSearchParams(window.location.search).get("token") || "";
  return (
    <section className="auth-card">
      <span className="soft-badge"><RotateCcw size={15} /> Восстановление</span>
      <h1>Новый пароль</h1>
      <p>Ссылка действует ограниченное время и может быть использована только один раз.</p>
      {!token && <Notice tone="error">Ссылка недействительна или срок действия истек.</Notice>}
      {token ? (
        <form className="form-stack" action="/reset-password" method="post">
          <input type="hidden" name="token" value={token} />
          <label>Новый пароль<input className="input" type="password" name="new_password" autoComplete="new-password" minLength={8} required /></label>
          <label>Повторите новый пароль<input className="input" type="password" name="confirm_password" autoComplete="new-password" minLength={8} required /></label>
          <button className="btn btn-primary btn-wide" type="submit"><CheckCircle2 size={18} /> Сохранить пароль</button>
        </form>
      ) : (
        <a className="btn btn-secondary" href="/forgot-password">Запросить новую ссылку</a>
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
        const isNetworkError = error instanceof TypeError || error.message === "Failed to fetch";
        setMessage(isNetworkError
          ? "Сервер не ответил на запрос. Обновите страницу и запустите обработку еще раз."
          : error.message || "Не удалось обработать файлы. Проверьте подключение и попробуйте еще раз.");
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
      setProgress({
        percent: Math.max(0, Math.min(100, Number(data.percent || 0))),
        message: data.message || "Обработка файлов",
      });
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
      <PageHeader
        title="Сравнение ТЗ и КП поставщиков"
        text="Загрузите заявку и коммерческие предложения. Сервис сопоставит позиции, цены и подготовит Excel-отчет."
        badge="XLSX · PDF · Excel-отчет"
        actions={<a className="btn btn-primary" href="/">Новая обработка</a>}
      />
      <section className="upload-grid">
        <form className="panel upload-panel" onSubmit={submit}>
          <Stepper active={isSubmitting ? 3 : offerFiles.length ? 2 : requestFiles.length ? 1 : 0} />
          <UploadField
            title="Заявка / ТЗ"
            description="Excel-файл со списком позиций, количеством и единицами измерения."
            accept=".xlsx,.xls,.pdf"
            files={requestFiles}
            setFiles={(files) => setRequestFiles(files.slice(0, 1))}
            multiple={false}
            emptyTitle="Перетащите файл заявки сюда"
            emptyText="или выберите .xlsx на компьютере"
            support="Поддерживается .xlsx. PDF можно использовать как базовый файл при сравнении двух КП."
            templateHref="/template/request.xlsx"
            templateText="Шаблон заявки"
          />
          <UploadField
            title="КП / счета поставщиков"
            description="Добавьте два или больше КП или счетов от поставщиков."
            accept=".xlsx,.xls,.pdf"
            files={offerFiles}
            setFiles={setOfferFiles}
            multiple
            emptyTitle="Перетащите КП или счета сюда"
            emptyText="или выберите несколько файлов"
            support="Поддерживаются .xlsx, .xls и текстовые PDF"
            templateHref="/template/offer.xlsx"
            templateText="Шаблон КП"
          />
          <label className="consent">
            <input type="checkbox" checked={consent} onChange={(event) => setConsent(event.target.checked)} />
            <span>Я согласен с <a href="/privacy" target="_blank" rel="noreferrer">обработкой персональных данных</a></span>
          </label>
          {message && <Notice tone="error">{message}</Notice>}
          {progress && <ProcessingState progress={progress} />}
          <div className="action-row">
            <button className="btn btn-primary btn-wide" type="submit" disabled={isSubmitting}>
              <FileSpreadsheet size={18} /> {isSubmitting ? "Формируем отчет" : "Сформировать Excel-отчет"}
            </button>
            {isSubmitting && <button className="btn btn-secondary" type="button" onClick={stop}><X size={17} /> Остановить</button>}
          </div>
          <p className="caption">Обычно обработка занимает 1–3 минуты. Большие заявки и сканы PDF могут обрабатываться дольше.</p>
        </form>
      </section>
    </>
  );
}

function Stepper({ active }) {
  return (
    <div className="stepper" aria-label="Этапы обработки">
      {["Заявка", "КП поставщиков", "Excel-отчет"].map((label, index) => (
        <div className={`step ${active >= index + 1 ? "is-active" : ""}`} key={label}>
          <span>{index + 1}</span>
          <b>{label}</b>
        </div>
      ))}
    </div>
  );
}

function UploadField({ title, description, accept, files, setFiles, multiple, emptyTitle, emptyText, support, templateHref, templateText }) {
  const inputRef = useRef(null);
  const [dragging, setDragging] = useState(false);

  const addFiles = (list) => {
    const selected = Array.from(list || []);
    if (!selected.length) return;
    setFiles(multiple ? [...files, ...selected] : selected.slice(0, 1));
  };

  const remove = (index) => setFiles(files.filter((_, fileIndex) => fileIndex !== index));

  return (
    <section className="upload-section">
      <div className="section-head">
        <div>
          <h2>{title}</h2>
          <p>{description}</p>
        </div>
        {templateHref && <a className="template-link" href={templateHref}><Download size={15} /> {templateText}</a>}
      </div>

      {!files.length ? (
        <label
          className={`dropzone upload-card ${dragging ? "is-dragging" : ""}`}
          onDragOver={(event) => { event.preventDefault(); setDragging(true); }}
          onDragLeave={() => setDragging(false)}
          onDrop={(event) => {
            event.preventDefault();
            setDragging(false);
            addFiles(event.dataTransfer.files);
          }}
        >
          <input ref={inputRef} type="file" accept={accept} multiple={multiple} onChange={(event) => addFiles(event.target.files)} />
          <span className="drop-icon"><Upload size={24} /></span>
          <div className="drop-copy">
            <b>{emptyTitle}</b>
            <p>{emptyText}</p>
          </div>
          <div className="drop-actions">
            <button className="btn btn-secondary" type="button" onClick={() => inputRef.current?.click()}>Выбрать файл{multiple ? "ы" : ""}</button>
            <span>{support}</span>
          </div>
        </label>
      ) : (
        <div className="selected-drop upload-card">
          <span className="selected-icon"><FileText size={22} /></span>
          <div>
            <b title={files[0].name}>{multiple ? `Загружено ${files.length} файл${files.length === 1 ? "" : files.length < 5 ? "а" : "ов"}` : files[0].name}</b>
            <span>{multiple ? "Файлы готовы к обработке" : `${fileSize(files[0].size)} · ${fileExt(files[0].name)}`}</span>
          </div>
          <StatusBadge status="done" />
          <button className="btn btn-secondary" type="button" onClick={() => inputRef.current?.click()}>Заменить</button>
          <button className="btn btn-secondary" type="button" onClick={() => setFiles([])}>Удалить</button>
          <input ref={inputRef} type="file" accept={accept} multiple={multiple} hidden onChange={(event) => addFiles(event.target.files)} />
        </div>
      )}

      <FilesList files={files} multiple={multiple} remove={remove} />
    </section>
  );
}

function FilesList({ files, multiple, remove }) {
  if (!files.length) {
    return (
      <div className="files-list empty">
        <FileText size={18} />
        <span>{multiple ? "Файлы пока не выбраны." : "Файл пока не выбран."}</span>
      </div>
    );
  }
  return (
    <div className="files-list is-filled">
      <div className="file-summary">{multiple ? `Загружено файлов: ${files.length}` : "Файл выбран"}</div>
      <ul>
        {files.map((file, index) => (
          <li key={`${file.name}-${file.size}-${file.lastModified}-${index}`}>
            <div className="file-main">
              <FileText size={18} />
              <div>
                <b title={file.name}>{file.name}</b>
                <span>{fileExt(file.name)} · {fileSize(file.size)}</span>
              </div>
            </div>
            <StatusBadge status="done" />
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
      <div className="processing-steps" aria-label="Этапы обработки">
        <span>Извлечение позиций</span>
        <span>Сопоставление</span>
        <span>Excel-отчет</span>
      </div>
      <small>Сервис извлекает позиции, цены и сроки, затем сопоставляет товары с учетом ИИ и сохраненных подтверждений.</small>
    </section>
  );
}

function AdminHistoryPage() {
  const [filters, setFilters] = useState({ date_from: "", date_to: "", supplier: "", status: "", percent_min: "" });
  const [data, setData] = useState(null);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(true);

  const load = async (nextFilters = filters) => {
    setLoading(true);
    setError("");
    const params = new URLSearchParams();
    Object.entries(nextFilters).forEach(([key, value]) => value && params.set(key, value));
    try {
      const payload = await fetchJson(`/api/admin/history${params.toString() ? `?${params}` : ""}`);
      setData(payload);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, []);

  const runs = data?.runs || [];
  const stats = data?.stats || {};
  return (
    <>
      <PageHeader
        title="История обработок"
        text="Фильтры, отчеты, спорные строки и быстрый переход к деталям обработки."
        actions={<a className="btn btn-primary" href="/">Новая обработка</a>}
      />
      <MetricGrid
        items={[
          ["Обработок", stats.total_runs ?? "—"],
          ["Готовых отчетов", stats.done_runs ?? "—"],
          ["Строк базы", stats.purchase_rows ?? "—"],
          ["Самообучение", stats.learning_rows ?? "—"],
        ]}
      />
      <section className="panel filter-panel">
        <form className="filter-bar" onSubmit={(event) => { event.preventDefault(); load(filters); }}>
          <Field label="Дата от"><input className="input" type="date" value={filters.date_from} onChange={(e) => setFilters({ ...filters, date_from: e.target.value })} /></Field>
          <Field label="Дата до"><input className="input" type="date" value={filters.date_to} onChange={(e) => setFilters({ ...filters, date_to: e.target.value })} /></Field>
          <Field label="Поставщик"><input className="input" placeholder="Название или ИНН" value={filters.supplier} onChange={(e) => setFilters({ ...filters, supplier: e.target.value })} /></Field>
          <Field label="Статус"><input className="input" placeholder="Все" value={filters.status} onChange={(e) => setFilters({ ...filters, status: e.target.value })} /></Field>
          <Field label="% от"><input className="input" inputMode="decimal" placeholder="80" value={filters.percent_min} onChange={(e) => setFilters({ ...filters, percent_min: e.target.value })} /></Field>
          <button className="btn btn-primary" type="submit"><Search size={17} /> Фильтровать</button>
        </form>
      </section>
      {error && <Notice tone="error">{error}</Notice>}
      <section className="panel table-panel">
        <div className="panel-title">
          <div><h2>История заявок</h2><p>{loading ? "Загрузка..." : `Показано: ${runs.length}`}</p></div>
          <button className="btn btn-secondary" type="button" onClick={() => load()}><RefreshCw size={16} /> Обновить</button>
        </div>
        <div className="table-wrap">
          <table className="data-table">
            <thead><tr><th>Дата</th><th>Файлы</th><th>Поставщики</th><th>Сопоставлено</th><th>Статус</th><th>Действия</th></tr></thead>
            <tbody>
              {runs.map((run) => (
                <tr key={run.run_id}>
                  <td><b>{formatDateTime(run.created_at)}</b><span>{run.run_id}</span></td>
                  <td>{run.request_files?.[0] || "—"}<span>{run.offer_files?.length || 0} КП</span></td>
                  <td>{run.suppliers?.length ? run.suppliers.join(", ") : "—"}</td>
                  <td><b>{Number(run.match_percent || 0).toFixed(1)}%</b><span>спорные: {run.review_count ?? 0} · не найдены: {run.unmatched_count ?? 0}</span></td>
                  <td><StatusBadge status={run.status} /></td>
                  <td><div className="table-actions"><a href={`/review/${run.run_id}`}>Проверка</a><a href={`/download/${run.run_id}/summary.xlsx`}>Excel</a><a href={`/download/${run.run_id}/review.xlsx`}>Review</a></div></td>
                </tr>
              ))}
              {!loading && !runs.length && <tr><td colSpan="6" className="empty-cell">История пока пустая или ничего не найдено по фильтрам.</td></tr>}
            </tbody>
          </table>
        </div>
      </section>
    </>
  );
}

function PurchasesPage() {
  const [data, setData] = useState(null);
  const [notice, setNotice] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(true);
  const fileRef = useRef(null);

  const load = async () => {
    setLoading(true);
    setError("");
    try {
      setData(await fetchJson("/api/admin/purchases"));
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, []);

  const submit = async (event) => {
    event.preventDefault();
    setNotice("");
    setError("");
    const form = new FormData(event.currentTarget);
    try {
      const result = await fetchJson("/admin/import-purchases", { method: "POST", body: form });
      setNotice([result.message, ...(result.warnings || [])].join(" "));
      if (fileRef.current) fileRef.current.value = "";
      load();
    } catch (err) {
      setError(err.message);
    }
  };

  const rows = data?.rows || [];
  return (
    <>
      <PageHeader
        title="База прошлых закупок"
        text="Импортируйте Excel-базу, чтобы сервис показывал прошлых лидеров по цене и историю тендеров."
        actions={<a className="btn btn-primary" href="/">Новая обработка</a>}
      />
      <section className="panel split-panel">
        <div>
          <h2>Импорт Excel</h2>
          <p>Поддерживаются колонки: позиция, поставщик, ИНН, цена, дата закупки, единица, количество, сумма, источник. Если части колонок нет, будут импортированы доступные данные.</p>
        </div>
        <form className="import-form" onSubmit={submit}>
          <input ref={fileRef} className="input" type="file" name="purchase_base" accept=".xlsx,.xlsm" required />
          <button className="btn btn-primary" type="submit"><Archive size={17} /> Импортировать</button>
        </form>
      </section>
      {notice && <Notice>{notice}</Notice>}
      {error && <Notice tone="error">{error}</Notice>}
      <section className="panel table-panel">
        <div className="panel-title">
          <div><h2>Загружено строк: {data?.count ?? "—"}</h2><p>{loading ? "Загрузка..." : "Последние 300 строк базы"}</p></div>
          <button className="btn btn-secondary" type="button" onClick={load}><RefreshCw size={16} /> Обновить</button>
        </div>
        <div className="table-wrap">
          <table className="data-table">
            <thead><tr><th>Импорт</th><th>Позиция</th><th>Поставщик</th><th>Цена</th><th>Количество</th><th>Сумма</th><th>Дата</th></tr></thead>
            <tbody>
              {rows.map((row, index) => (
                <tr key={`${row.imported_at}-${index}`}>
                  <td>{formatDateTime(row.imported_at)}<span>{row.source_file}</span></td>
                  <td>{row.position_name}</td>
                  <td>{row.supplier || "—"}<span>{row.supplier_inn}</span></td>
                  <td>{row.price_text || "—"}</td>
                  <td>{row.qty ?? "—"} {row.unit || ""}</td>
                  <td>{row.total_text || "—"}</td>
                  <td>{row.purchase_date || "—"}</td>
                </tr>
              ))}
              {!loading && !rows.length && <tr><td colSpan="7" className="empty-cell">База прошлых закупок пока не импортирована.</td></tr>}
            </tbody>
          </table>
        </div>
      </section>
    </>
  );
}

function AuditPage() {
  const [data, setData] = useState(null);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(true);

  const load = async () => {
    setLoading(true);
    setError("");
    try {
      setData(await fetchJson("/api/admin/audit"));
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, []);
  const rows = data?.rows || [];

  return (
    <>
      <PageHeader title="Журнал действий" text="Технический учет входов, обработок, импортов и формирования отчетов." actions={<button className="btn btn-secondary" type="button" onClick={load}><RefreshCw size={16} /> Обновить</button>} />
      {error && <Notice tone="error">{error}</Notice>}
      <section className="panel table-panel">
        <div className="panel-title"><div><h2>Последние действия</h2><p>{loading ? "Загрузка..." : `Записей: ${rows.length}`}</p></div></div>
        <div className="table-wrap">
          <table className="data-table">
            <thead><tr><th>Дата</th><th>Пользователь</th><th>Действие</th><th>IP</th><th>Детали</th></tr></thead>
            <tbody>
              {rows.map((row, index) => (
                <tr key={`${row.created_at}-${index}`}>
                  <td>{formatDateTime(row.created_at)}</td>
                  <td>{row.username || "—"}</td>
                  <td><b>{row.action}</b><span>{row.object_type} {row.object_id}</span></td>
                  <td>{row.ip || "—"}</td>
                  <td><code>{row.details_json}</code></td>
                </tr>
              ))}
              {!loading && !rows.length && <tr><td colSpan="5" className="empty-cell">Журнал пока пустой.</td></tr>}
            </tbody>
          </table>
        </div>
      </section>
    </>
  );
}

function SecurityPage() {
  const [data, setData] = useState(null);
  const [notice, setNotice] = useState("");
  const [error, setError] = useState("");

  const load = async () => {
    try {
      setData(await fetchJson("/api/admin/security"));
    } catch (err) {
      setError(err.message);
    }
  };

  useEffect(() => { load(); }, []);

  const postForm = async (event, url) => {
    event.preventDefault();
    setNotice("");
    setError("");
    try {
      const result = await fetchJson(url, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams(new FormData(event.currentTarget)),
      });
      setNotice(result.message);
      event.currentTarget.reset();
      load();
    } catch (err) {
      setError(err.message);
    }
  };

  return (
    <>
      <PageHeader title="Настройки" text="Смена пароля, email восстановления и параметры доступа без регистрации пользователей." />
      {notice && <Notice>{notice}</Notice>}
      {error && <Notice tone="error">{error}</Notice>}
      <section className="security-grid">
        <div className="panel">
          <div className="panel-title">
            <div><h2>Доступ администратора</h2><p>Логин: {data?.admin_username || "admin"}</p></div>
          </div>
          <form className="form-stack" onSubmit={(event) => postForm(event, "/admin/change-password")}>
            <label>Текущий пароль<input className="input" type="password" name="current_password" autoComplete="current-password" required /></label>
            <label>Новый пароль<input className="input" type="password" name="new_password" autoComplete="new-password" minLength={8} required /></label>
            <label>Повторите новый пароль<input className="input" type="password" name="confirm_password" autoComplete="new-password" minLength={8} required /></label>
            <button className="btn btn-primary" type="submit"><KeyRound size={17} /> Сменить пароль</button>
          </form>
        </div>
        <div className="panel">
          <div className="panel-title">
            <div><h2>Email восстановления</h2><p>Текущий адрес: {data?.admin_email || "—"}</p></div>
          </div>
          <form className="form-stack" onSubmit={(event) => postForm(event, "/admin/change-email")}>
            <label>Новый email<input className="input" type="email" name="admin_email" defaultValue={data?.admin_email || ""} autoComplete="email" required /></label>
            <label>Текущий пароль<input className="input" type="password" name="current_password" autoComplete="current-password" required /></label>
            <button className="btn btn-primary" type="submit"><Mail size={17} /> Сохранить email</button>
          </form>
        </div>
      </section>
    </>
  );
}

function ReviewPage({ runId }) {
  const safeRunId = useMemo(() => encodeURIComponent(runId), [runId]);
  const [data, setData] = useState(null);
  const [error, setError] = useState("");
  const [query, setQuery] = useState("");
  const [submitting, setSubmitting] = useState(false);

  useEffect(() => {
    let active = true;
    setError("");
    fetchJson(`/api/review/${safeRunId}`)
      .then((payload) => { if (active) setData(payload); })
      .catch((err) => { if (active) setError(err.message); });
    return () => { active = false; };
  }, [safeRunId]);

  const rows = data?.rows || [];
  const filteredRows = rows.filter((row) => {
    const value = query.trim().toLowerCase();
    if (!value) return true;
    return [
      row.request_label,
      row.name,
      row.supplier,
      row.reason,
      row.status_label,
    ].join(" ").toLowerCase().includes(value);
  });
  const stats = data?.stats || {};
  const suppliersCount = data?.suppliers?.length ?? 0;

  const submit = () => {
    setSubmitting(true);
  };

  if (error) {
    return (
      <>
        <PageHeader title="Проверка сопоставлений" text="Не удалось открыть обработку." />
        <Notice tone="error">{error}</Notice>
      </>
    );
  }

  if (!data) {
    return (
      <>
        <PageHeader title="Проверка сопоставлений" text="Загружаем строки ручной проверки." />
        <section className="panel review-loading">
          <Clock3 size={18} />
          <span>Загрузка данных проверки...</span>
        </section>
      </>
    );
  }

  return (
    <>
      <PageHeader
        title="Проверка сопоставлений"
        text="Подтвердите спорные строки перед формированием итоговой Excel-сводки."
        actions={<a className="btn btn-secondary" href={`/download/${safeRunId}/review.xlsx`} download><FileCheck2 size={17} /> Файл проверки</a>}
      />

      <MetricGrid
        className="review-metrics"
        items={[
          ["Сопоставлено", `${Number(data.match_percent || 0).toFixed(1)}%`],
          ["Позиций", stats.request ?? "—"],
          ["На проверке", rows.length],
          ["Поставщиков", suppliersCount],
          ["Не найдено", stats.unmatched ?? "—"],
        ]}
      />

      {data.errors?.map((item, index) => (
        <Notice key={`${item}-${index}`} tone={item.includes("API-ключ") || item.includes("DeepSeek") ? "error" : "warning"} compact>
          {item}
        </Notice>
      ))}

      <form className="review-workspace" action={`/finalize/${safeRunId}`} method="post" onSubmit={submit}>
        <datalist id="unit-options">
          {["м2", "м3", "шт", "п.м", "м", "кг", "т", "упак"].map((unit) => <option key={unit} value={unit} />)}
        </datalist>
        <section className="panel review-panel">
          <div className="panel-title review-title">
            <div>
              <h2>Строки для проверки</h2>
              <p>{query ? `Показано ${filteredRows.length} из ${rows.length}` : `Всего строк: ${rows.length}`}</p>
            </div>
            <div className="review-search">
              <Search size={16} />
              <input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Поиск по позиции или поставщику" />
            </div>
          </div>

          <div className="review-table-wrap">
            <table className="review-data-table">
              <thead>
                <tr>
                  <th>№</th>
                  <th>Позиция заявки</th>
                  <th>Ед.</th>
                  <th>Поставщик</th>
                  <th>Позиция КП</th>
                  <th>Кол-во</th>
                  <th>Цена</th>
                  <th>Статус</th>
                  <th>Ручная корректировка</th>
                </tr>
              </thead>
              <tbody>
                {filteredRows.map((row) => (
                  <ReviewRow key={row.idx} row={row} requestOptions={data.request_options || []} />
                ))}
                {!filteredRows.length && (
                  <tr><td colSpan="9" className="empty-cell">Строки не найдены.</td></tr>
                )}
              </tbody>
            </table>
          </div>
        </section>

        <section className="review-actionbar">
          <div>
            <b>Проверено {Math.max((stats.comparable || 0) - rows.length, 0)} из {stats.comparable || stats.offers || 0} позиций</b>
            <span>После подтверждения будет сформирован итоговый Excel-отчет.</span>
          </div>
          <div className="review-actionbar-buttons">
            <a className="btn btn-secondary" href="/">Новая обработка</a>
            <button className="btn btn-primary" type="submit" disabled={submitting}>
              {submitting ? "Формируем..." : "Сформировать Excel-отчет"}
            </button>
          </div>
        </section>
      </form>
    </>
  );
}

function ReviewRow({ row, requestOptions }) {
  const [requestPos, setRequestPos] = useState(row.request_pos || "");
  const selectedOption = requestOptions.find((item) => item.pos === requestPos);
  const unit = selectedOption?.unit || row.request_unit || "—";

  return (
    <tr className={row.status_class === "unmatched" ? "is-soft-warning" : ""}>
      <td className="review-index">{row.idx + 1}</td>
      <td className="review-match-cell">
        <select
          className="input review-select"
          name={`match_${row.idx}`}
          value={requestPos}
          onChange={(event) => setRequestPos(event.target.value)}
          title={selectedOption?.label || row.request_label || "Не сопоставлять"}
        >
          <option value="">Не сопоставлять</option>
          {requestOptions.map((option) => (
            <option key={option.pos} value={option.pos}>{option.label}</option>
          ))}
        </select>
        {!!row.suggestions?.length && (
          <div className="review-suggestions">
            {row.suggestions.slice(0, 4).map((suggestion) => (
              <button
                key={`${row.idx}-${suggestion.pos}`}
                type="button"
                className="suggestion-chip"
                title={suggestion.label}
                onClick={() => setRequestPos(suggestion.pos)}
              >
                {suggestion.pos}
                <span>{suggestion.score}%</span>
              </button>
            ))}
          </div>
        )}
      </td>
      <td className="review-unit">{unit}</td>
      <td>{row.supplier || "—"}<span className="row-muted">строка {row.row_no || "—"}</span></td>
      <td className="review-offer-name" title={row.name}>{row.name || "—"}</td>
      <td>{row.qty ?? "—"} {row.unit || ""}</td>
      <td>{row.price ?? "—"}</td>
      <td><StatusBadge status={row.status_class || row.status} /></td>
      <td>
        <div className="review-adjust">
          <input className="input compact" type="text" name={`norm_qty_${row.idx}`} defaultValue={row.override_qty ?? ""} placeholder="40,698" />
          <input className="input compact unit" type="text" name={`norm_unit_${row.idx}`} defaultValue={row.override_unit || ""} placeholder="м3" list="unit-options" />
        </div>
        <textarea className="input note" name={`note_${row.idx}`} defaultValue={row.note || ""} placeholder="Примечание по расхождению" />
        <small>{row.reason || "проверить совпадение"}</small>
      </td>
    </tr>
  );
}

function DonePage({ runId }) {
  const safeRunId = useMemo(() => encodeURIComponent(runId), [runId]);
  const [data, setData] = useState(null);

  useEffect(() => {
    let active = true;
    fetchJson(`/api/review/${safeRunId}`)
      .then((payload) => { if (active) setData(payload); })
      .catch(() => {});
    return () => { active = false; };
  }, [safeRunId]);

  const stats = data?.stats || {};
  const suppliersCount = data?.suppliers?.length ?? "—";
  return (
    <>
      <PageHeader title="Excel-отчет готов" text="Сводка сформирована. Можно скачать отчет или открыть историю обработки." actions={<a className="btn btn-secondary" href="/">Новая обработка</a>} />
      <section className="panel done-panel done-panel-wide">
        <div className="done-hero">
          <span className="ready-badge"><CheckCircle2 size={16} /> Готово</span>
          <div>
            <h2>Excel-отчет готов</h2>
            <p>Найдено {stats.request ?? "—"} позиций, {(stats.review ?? 0) + (stats.unmatched ?? 0)} спорных совпадений, {suppliersCount} поставщиков.</p>
          </div>
          <a className="btn btn-primary" href={`/download/${safeRunId}/summary.xlsx`} download><Download size={19} /> Скачать Excel-отчет</a>
        </div>
        <MetricGrid
          items={[
            ["Позиций заявки", stats.request ?? "—"],
            ["Сопоставлено", data ? `${Number(data.match_percent || 0).toFixed(1)}%` : "—"],
            ["Спорные строки", (stats.review ?? 0) + (stats.unmatched ?? 0)],
            ["Поставщики", suppliersCount],
          ]}
        />
        <div className="done-actions">
          <a className="btn btn-secondary" href={`/download/${safeRunId}/review.xlsx`} download><FileCheck2 size={18} /> Файл проверки</a>
          <a className="btn btn-secondary" href={`/review/${safeRunId}`}><ArrowRight size={18} /> Вернуться к проверке</a>
          <a className="btn btn-secondary" href="/admin/history"><History size={18} /> История</a>
        </div>
      </section>
    </>
  );
}

function MetricGrid({ items, className = "" }) {
  return (
    <section className={`metric-grid ${className}`.trim()}>
      {items.map(([label, value]) => (
        <div className="metric-card" key={label}>
          <span>{label}</span>
          <b>{value}</b>
        </div>
      ))}
    </section>
  );
}

function Field({ label, children }) {
  return <label className="field"><span>{label}</span>{children}</label>;
}

function StatusBadge({ status }) {
  const key = status || "";
  const label = STATUS_LABELS[key] || key || "—";
  return <span className={`status-badge ${key}`}>{label}</span>;
}

function Notice({ children, tone = "success", compact = false }) {
  return (
    <div className={`notice ${tone} ${compact ? "is-compact" : ""}`}>
      <span className="notice-dot" aria-hidden="true" />
      <span>{children}</span>
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);
