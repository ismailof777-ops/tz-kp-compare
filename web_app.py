from __future__ import annotations

from email.parser import BytesParser
from email.policy import default as email_policy
import html
import io
import json
import mimetypes
import os
import shutil
import sys
import uuid
from dataclasses import asdict
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

from compare_tz_kp import (
    Match,
    RequestItem,
    SupplierItem,
    build_matches,
    clean_text,
    read_offer,
    read_request_xlsx,
    status_label,
    write_final,
    write_review,
)


ROOT = Path(__file__).resolve().parent
RUNS_DIR = ROOT / "outputs" / "runs"
MAX_UPLOAD_SIZE = 80 * 1024 * 1024


class UploadedFile:
    def __init__(self, filename: str, content: bytes):
        self.filename = filename
        self.file = io.BytesIO(content)


CSS = """
:root {
  color-scheme: light;
  --bg: #f6f8fb;
  --panel: #ffffff;
  --text: #172033;
  --muted: #667085;
  --line: #d9e0ea;
  --accent: #176b87;
  --accent-strong: #0f566c;
  --green: #d9f2df;
  --red: #f9d6d5;
  --yellow: #fff2c6;
}
* { box-sizing: border-box; }
body {
  margin: 0;
  background: var(--bg);
  color: var(--text);
  font: 14px/1.45 Arial, Helvetica, sans-serif;
}
.shell {
  width: min(1180px, calc(100% - 32px));
  margin: 0 auto;
  padding: 28px 0 40px;
}
.topbar {
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 16px;
  margin-bottom: 18px;
}
h1 {
  margin: 0;
  font-size: 24px;
  line-height: 1.15;
  letter-spacing: 0;
}
.subtitle {
  margin: 7px 0 0;
  color: var(--muted);
  max-width: 760px;
}
.product-intro {
  max-width: 920px;
  margin-bottom: 16px;
  padding: 18px;
  border: 1px solid var(--line);
  border-radius: 8px;
  background: #ffffff;
  box-shadow: 0 10px 24px rgba(16, 24, 40, 0.05);
}
.product-intro h2 {
  margin: 0;
  font-size: 20px;
  line-height: 1.25;
}
.product-intro p {
  margin: 8px 0 0;
  color: var(--muted);
  max-width: 760px;
}
.intro-points {
  display: flex;
  gap: 8px;
  flex-wrap: wrap;
  margin-top: 12px;
}
.intro-point {
  border: 1px solid var(--line);
  border-radius: 999px;
  background: #fbfcfe;
  padding: 7px 11px;
  color: var(--muted);
  font-size: 13px;
}
.intro-point b {
  display: inline;
  margin: 0;
  color: var(--text);
}
.intro-point span {
  display: inline;
  color: var(--muted);
}
.panel {
  background: var(--panel);
  border: 1px solid var(--line);
  border-radius: 8px;
  padding: 18px;
  box-shadow: 0 10px 24px rgba(16, 24, 40, 0.05);
}
.grid {
  display: grid;
  grid-template-columns: minmax(0, 1fr);
  gap: 16px;
}
.upload-panel {
  max-width: 920px;
}
.field { margin-bottom: 16px; }
label {
  display: block;
  font-weight: 700;
  margin-bottom: 7px;
}
select,
input[type="text"],
input[type="search"] {
  width: 100%;
  border: 1px solid var(--line);
  border-radius: 6px;
  background: #fff;
  padding: 10px;
  min-height: 42px;
}
select,
input[type="text"],
input[type="search"] { font: inherit; }
.match-input {
  min-width: 300px;
}
.match-help {
  margin-top: 5px;
  color: var(--muted);
  font-size: 12px;
}
.file-drop {
  position: relative;
  overflow: hidden;
  display: grid;
  grid-template-columns: auto minmax(0, 1fr);
  gap: 12px;
  align-items: center;
  min-height: 86px;
  border: 1px dashed #aebdca;
  border-radius: 8px;
  background: #fbfcfe;
  padding: 14px;
  transition: border-color .15s ease, background .15s ease, box-shadow .15s ease;
}
.file-drop:hover,
.file-drop.is-dragover {
  border-color: var(--accent);
  background: #f0f8fb;
  box-shadow: 0 0 0 3px rgba(23, 107, 135, 0.08);
}
.file-input {
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
  opacity: 0;
  cursor: pointer;
  z-index: 3;
}
.file-drop > :not(.file-input) {
  position: relative;
  z-index: 4;
  pointer-events: none;
}
.file-icon {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 38px;
  height: 38px;
  border-radius: 999px;
  background: #e8f4f8;
  color: var(--accent);
  font-size: 24px;
  line-height: 1;
  font-weight: 700;
}
.file-title {
  display: block;
  font-weight: 700;
}
.file-subtitle {
  display: block;
  margin-top: 3px;
  color: var(--muted);
  font-size: 13px;
}
.file-list {
  position: relative;
  z-index: 5;
  pointer-events: auto;
  margin-top: 8px;
  color: var(--muted);
  font-size: 13px;
}
.file-list.is-filled {
  color: var(--text);
}
.file-list b {
  display: block;
  margin-bottom: 4px;
}
.file-list ul {
  margin: 0;
  padding-left: 0;
  list-style: none;
}
.file-list li {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  margin: 2px 0;
  padding: 5px 6px 5px 8px;
  border: 1px solid var(--line);
  border-radius: 6px;
  background: #fff;
  overflow-wrap: anywhere;
}
.file-name {
  min-width: 0;
  overflow-wrap: anywhere;
}
.file-remove {
  flex: 0 0 auto;
  min-height: 26px;
  border: 1px solid #e0b7b5;
  border-radius: 5px;
  background: #fff5f5;
  color: #8a2420;
  padding: 0 8px;
  font: inherit;
  font-size: 12px;
  cursor: pointer;
}
.file-remove:hover {
  background: #ffe9e9;
}
.hint {
  margin-top: 6px;
  color: var(--muted);
  font-size: 13px;
}
.format-note {
  display: flex;
  gap: 8px;
  flex-wrap: wrap;
  margin: 2px 0 16px;
}
.pill {
  border: 1px solid var(--line);
  border-radius: 999px;
  padding: 5px 9px;
  color: var(--muted);
  background: #fbfcfe;
  font-size: 12px;
}
.actions {
  display: flex;
  align-items: center;
  gap: 10px;
  flex-wrap: wrap;
  margin-top: 14px;
}
.btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-height: 40px;
  padding: 0 15px;
  border: 1px solid var(--accent);
  border-radius: 6px;
  background: var(--accent);
  color: #fff;
  font-weight: 700;
  text-decoration: none;
  cursor: pointer;
}
.btn:hover { background: var(--accent-strong); border-color: var(--accent-strong); }
.btn[disabled] {
  cursor: wait;
  opacity: .75;
}
.btn.secondary {
  background: #fff;
  color: var(--accent);
}
.btn.secondary:hover { background: #eef8fb; }
.btn.stop {
  display: none;
  border-color: #d79c99;
  background: #fff;
  color: #8a2420;
}
.btn.stop.is-visible {
  display: inline-flex;
}
.btn.stop:hover { background: #fff0f0; }
.stats {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 10px;
}
.stat {
  border: 1px solid var(--line);
  border-radius: 6px;
  padding: 12px;
  background: #fbfcfe;
}
.stat b {
  display: block;
  font-size: 22px;
}
.notice {
  border-radius: 6px;
  border: 1px solid var(--line);
  padding: 12px;
  margin-bottom: 16px;
  background: #fff;
}
.notice.warn { border-color: #f4c76d; background: #fff8e6; }
.notice.error { border-color: #eda5a3; background: #fff0f0; }
.table-wrap {
  overflow: auto;
  border: 1px solid var(--line);
  border-radius: 8px;
  background: #fff;
}
.review-tools {
  display: grid;
  grid-template-columns: minmax(240px, 420px) minmax(0, 1fr);
  gap: 12px;
  align-items: end;
  margin-bottom: 12px;
}
.review-tools label {
  margin-bottom: 5px;
}
.review-count {
  color: var(--muted);
  font-size: 13px;
  padding-bottom: 11px;
}
table {
  border-collapse: collapse;
  width: 100%;
  min-width: 980px;
}
th, td {
  border-bottom: 1px solid var(--line);
  padding: 10px;
  text-align: left;
  vertical-align: top;
}
th {
  position: sticky;
  top: 0;
  background: #eaf3f8;
  z-index: 1;
  font-size: 13px;
}
td.small, th.small { width: 118px; }
.review-table th:nth-child(1),
.review-table td:nth-child(1) {
  position: sticky;
  left: 0;
  z-index: 2;
  width: 118px;
  min-width: 118px;
  background: #fff;
}
.review-table th:nth-child(2),
.review-table td:nth-child(2) {
  position: sticky;
  left: 118px;
  z-index: 2;
  width: 390px;
  min-width: 390px;
  background: #fff;
  box-shadow: 8px 0 12px rgba(16, 24, 40, 0.06);
}
.review-table th:nth-child(1),
.review-table th:nth-child(2) {
  z-index: 4;
  background: #eaf3f8;
}
.review-table .match-input {
  min-width: 0;
}
.status {
  display: inline-block;
  max-width: 112px;
  border: 1px solid transparent;
  border-radius: 6px;
  padding: 4px 6px;
  font-size: 11px;
  line-height: 1.2;
  font-weight: 700;
  text-align: center;
}
.status.review { background: #fff6d8; border-color: #f1d88d; color: #6f5200; }
.status.unmatched { background: #fdeaea; border-color: #f2b9b7; color: #8a2420; }
.status.auto { background: #e8f6ec; border-color: #bfe5c8; color: #236336; }
.status.manual { background: #eef4fa; border-color: #cbdcea; color: #244f73; }
.muted { color: var(--muted); }
.danger { background: var(--red); }
.empty {
  padding: 26px;
  text-align: center;
  color: var(--muted);
}
.processing {
  display: none;
  margin-top: 16px;
  border: 1px solid #b8d6e2;
  border-radius: 8px;
  background: #f1fbff;
  padding: 14px;
}
.processing.is-visible { display: block; }
.processing-head {
  display: flex;
  justify-content: space-between;
  gap: 16px;
  margin-bottom: 10px;
  font-weight: 700;
}
.processing-head span {
  color: var(--muted);
  font-weight: 400;
}
.progress {
  overflow: hidden;
  height: 9px;
  border-radius: 999px;
  background: #d5e8ef;
}
.progress-bar {
  width: 42%;
  height: 100%;
  border-radius: 999px;
  background: var(--accent);
  animation: progress-slide 1.1s ease-in-out infinite;
}
@keyframes progress-slide {
  0% { transform: translateX(-120%); }
  50% { transform: translateX(80%); }
  100% { transform: translateX(260%); }
}
@media (max-width: 840px) {
  .shell { width: min(100% - 20px, 1180px); padding-top: 18px; }
  .topbar, .grid { display: block; }
  .upload-panel { max-width: none; }
  .panel { padding: 14px; }
  .product-intro { padding: 14px; }
  .intro-points { display: flex; }
  .stats { grid-template-columns: 1fr; margin-top: 14px; }
  .review-tools { grid-template-columns: 1fr; }
  .review-count { padding-bottom: 0; }
  h1 { font-size: 21px; }
  .btn { width: 100%; }
}
"""


def esc(value) -> str:
    return html.escape("" if value is None else str(value), quote=True)


def page(title: str, body: str) -> bytes:
    html_doc = f"""<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{esc(title)}</title>
  <style>{CSS}</style>
</head>
<body>
  <main class="shell">{body}</main>
  <script>
    const uploadForm = document.querySelector('[data-upload-form]');
    if (uploadForm) {{
      const requestInput = uploadForm.querySelector('input[name="request"]');
      const offersInput = uploadForm.querySelector('input[name="offers"]');
      const formMessage = uploadForm.querySelector('[data-form-message]');
      const submitButton = uploadForm.querySelector('[data-submit]');
      const stopButton = uploadForm.querySelector('[data-stop]');
      const processing = uploadForm.querySelector('[data-processing]');
      let activeController = null;
      const showMessage = (text) => {{
        if (!formMessage) return;
        formMessage.textContent = text;
        formMessage.style.display = 'block';
      }};
      const resetProcessingState = () => {{
        activeController = null;
        if (submitButton) {{
          submitButton.disabled = false;
          submitButton.textContent = 'Обработать файлы';
        }}
        if (stopButton) stopButton.classList.remove('is-visible');
        if (processing) processing.classList.remove('is-visible');
      }};
      const removeInputFile = (input, indexToRemove) => {{
        const transfer = new DataTransfer();
        Array.from(input.files || []).forEach((file, index) => {{
          if (index !== indexToRemove) transfer.items.add(file);
        }});
        input.files = transfer.files;
        input.dispatchEvent(new Event('change', {{ bubbles: true }}));
      }};
      const renderFileList = (input) => {{
        const target = uploadForm.querySelector('[data-file-list="' + input.id + '"]');
        if (!target) return;
        const files = Array.from(input.files || []);
        target.textContent = '';
        target.classList.toggle('is-filled', files.length > 0);
        if (!files.length) {{
          target.textContent = input.multiple ? 'Файлы пока не выбраны.' : 'Файл пока не выбран.';
          return;
        }}
        const summary = document.createElement('b');
        summary.textContent = input.multiple
          ? 'Выбрано файлов: ' + files.length
          : 'Выбран файл:';
        target.appendChild(summary);
        const list = document.createElement('ul');
        files.forEach((file, index) => {{
          const item = document.createElement('li');
          const name = document.createElement('span');
          name.className = 'file-name';
          name.textContent = file.name;
          const remove = document.createElement('button');
          remove.className = 'file-remove';
          remove.type = 'button';
          remove.textContent = 'Удалить';
          remove.addEventListener('click', (event) => {{
            event.preventDefault();
            event.stopPropagation();
            removeInputFile(input, index);
          }});
          item.appendChild(name);
          item.appendChild(remove);
          list.appendChild(item);
        }});
        target.appendChild(list);
      }};
      const fileKey = (file) => [file.name, file.size, file.lastModified].join('|');
      const setInputFiles = (input, fileList, append = false) => {{
        const transfer = new DataTransfer();
        const files = [];
        if (append && input.multiple) {{
          files.push(...Array.from(input.files || []));
        }}
        files.push(...Array.from(fileList || []));
        const seen = new Set();
        files.slice(0, input.multiple ? undefined : 1).forEach((file) => {{
          const key = fileKey(file);
          if (seen.has(key)) return;
          seen.add(key);
          transfer.items.add(file);
        }});
        input.files = transfer.files;
        input.dispatchEvent(new Event('change', {{ bubbles: true }}));
      }};
      [requestInput, offersInput].forEach((input) => {{
        if (!input) return;
        input.addEventListener('change', () => {{
          if (formMessage) formMessage.style.display = 'none';
          renderFileList(input);
        }});
        renderFileList(input);
      }});
      uploadForm.querySelectorAll('[data-drop-zone]').forEach((zone) => {{
        const input = zone.querySelector('input[type="file"]');
        if (!input) return;
        ['dragenter', 'dragover'].forEach((eventName) => {{
          zone.addEventListener(eventName, (event) => {{
            event.preventDefault();
            zone.classList.add('is-dragover');
          }});
        }});
        ['dragleave', 'drop'].forEach((eventName) => {{
          zone.addEventListener(eventName, (event) => {{
            event.preventDefault();
            zone.classList.remove('is-dragover');
          }});
        }});
        zone.addEventListener('drop', (event) => {{
          setInputFiles(input, event.dataTransfer.files, input.multiple);
        }});
      }});
      stopButton?.addEventListener('click', () => {{
        if (activeController) activeController.abort();
        resetProcessingState();
        showMessage('Обработка остановлена. Можно изменить файлы и запустить заново.');
      }});
      uploadForm.addEventListener('submit', async (event) => {{
        event.preventDefault();
        if (!requestInput || !requestInput.files.length) {{
          showMessage('Выберите файл заявки / ТЗ.');
          requestInput && requestInput.focus();
          return;
        }}
        if (!offersInput || offersInput.files.length < 2) {{
          showMessage('Выберите минимум два КП или счета поставщиков.');
          offersInput && offersInput.focus();
          return;
        }}
        if (formMessage) formMessage.style.display = 'none';
        activeController = new AbortController();
        if (submitButton) {{
          submitButton.disabled = true;
          submitButton.textContent = 'Обрабатываем...';
        }}
        if (stopButton) stopButton.classList.add('is-visible');
        if (processing) processing.classList.add('is-visible');
        try {{
          const response = await fetch(uploadForm.action, {{
            method: 'POST',
            body: new FormData(uploadForm),
            signal: activeController.signal,
          }});
          if (response.redirected) {{
            window.location.href = response.url;
            return;
          }}
          const responseHtml = await response.text();
          document.open();
          document.write(responseHtml);
          document.close();
        }} catch (error) {{
          if (error.name !== 'AbortError') {{
            resetProcessingState();
            showMessage('Не удалось обработать файлы. Проверьте подключение и попробуйте еще раз.');
          }}
        }}
      }});
    }}
    const reviewSearch = document.querySelector('[data-review-search]');
    const reviewRows = Array.from(document.querySelectorAll('[data-review-row]'));
    const reviewCount = document.querySelector('[data-review-count]');
    if (reviewSearch && reviewRows.length) {{
      const rowText = (row) => (row.textContent + ' ' + (row.querySelector('input')?.value || '')).toLowerCase();
      const updateReviewSearch = () => {{
        const query = reviewSearch.value.trim().toLowerCase();
        let shown = 0;
        reviewRows.forEach((row) => {{
          const visible = !query || rowText(row).includes(query);
          row.style.display = visible ? '' : 'none';
          if (visible) shown += 1;
        }});
        if (reviewCount) {{
          reviewCount.textContent = query
            ? 'Показано строк: ' + shown + ' из ' + reviewRows.length
            : 'Всего строк для проверки: ' + reviewRows.length;
        }}
      }};
      reviewSearch.addEventListener('input', updateReviewSearch);
      reviewRows.forEach((row) => row.querySelector('input')?.addEventListener('change', updateReviewSearch));
      updateReviewSearch();
    }}
  </script>
</body>
</html>"""
    return html_doc.encode("utf-8")


def safe_filename(filename: str, fallback: str) -> str:
    raw = Path(filename or fallback).name.replace("\x00", "")
    cleaned = "".join(ch if ch.isalnum() or ch in " ._-()[]{}" else "_" for ch in raw).strip()
    return cleaned or fallback


def parse_multipart_upload(headers, body: bytes) -> dict[str, list[UploadedFile]]:
    content_type = headers.get("Content-Type", "")
    if not content_type.lower().startswith("multipart/form-data"):
        return {}

    message = BytesParser(policy=email_policy).parsebytes(
        b"Content-Type: "
        + content_type.encode("utf-8", errors="ignore")
        + b"\r\nMIME-Version: 1.0\r\n\r\n"
        + body
    )
    files: dict[str, list[UploadedFile]] = {}
    for part in message.iter_parts():
        disposition = part.get_content_disposition()
        if disposition != "form-data":
            continue
        field_name = part.get_param("name", header="content-disposition")
        filename = part.get_filename()
        if not field_name or not filename:
            continue
        payload = part.get_payload(decode=True) or b""
        files.setdefault(field_name, []).append(UploadedFile(clean_text(filename), payload))
    return files


def request_item_from_dict(data: dict) -> RequestItem:
    return RequestItem(
        pos=clean_text(data.get("pos")),
        name=clean_text(data.get("name")),
        specs=clean_text(data.get("specs")),
        unit=clean_text(data.get("unit")),
        qty=data.get("qty"),
    )


def supplier_item_from_dict(data: dict) -> SupplierItem:
    return SupplierItem(
        supplier=clean_text(data.get("supplier")),
        source=clean_text(data.get("source")),
        row_no=clean_text(data.get("row_no")),
        name=clean_text(data.get("name")),
        qty=data.get("qty"),
        unit=clean_text(data.get("unit")),
        price=data.get("price"),
        total=data.get("total"),
        delivery=clean_text(data.get("delivery")),
    )


def match_from_dict(data: dict) -> Match:
    supplier_item = supplier_item_from_dict(data["supplier_item"])
    return Match(
        supplier_item=supplier_item,
        request_pos=clean_text(data.get("request_pos")),
        score=float(data.get("score") or 0),
        status=clean_text(data.get("status")) or "review",
        reason=clean_text(data.get("reason") or ""),
    )


def save_state(run_dir: Path, request_items: list[RequestItem], matches: list[Match], errors: list[str]) -> None:
    state = {
        "request_items": [asdict(item) for item in request_items],
        "matches": [
            {
                "supplier_item": asdict(match.supplier_item),
                "request_pos": match.request_pos,
                "score": match.score,
                "status": match.status,
                "reason": match.reason,
            }
            for match in matches
        ],
        "errors": errors,
    }
    (run_dir / "state.json").write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def load_state(run_dir: Path) -> tuple[list[RequestItem], list[Match], list[str]]:
    state = json.loads((run_dir / "state.json").read_text(encoding="utf-8"))
    request_items = [request_item_from_dict(item) for item in state["request_items"]]
    matches = [match_from_dict(item) for item in state["matches"]]
    return request_items, matches, state.get("errors", [])


def stats_for(request_items: list[RequestItem], matches: list[Match]) -> dict[str, int]:
    return {
        "request": len(request_items),
        "offers": len(matches),
        "auto": sum(1 for match in matches if match.status == "auto" and match.request_pos),
        "review": sum(1 for match in matches if match.status == "review" and match.request_pos),
        "unmatched": sum(1 for match in matches if not match.request_pos),
    }


def resolve_request_pos(raw_value: str, request_items: list[RequestItem]) -> str | None:
    value = clean_text(raw_value)
    if value.lower() in {"", "-", "нет", "не сопоставлять"}:
        return None
    by_pos = {item.pos: item.pos for item in request_items}
    if value in by_pos:
        return value
    for item in request_items:
        prefix = f"{item.pos} - "
        if value.startswith(prefix):
            return item.pos
    value_lower = value.lower()
    matches = [
        item.pos
        for item in request_items
        if value_lower in item.pos.lower() or value_lower in item.name.lower()
    ]
    return matches[0] if len(matches) == 1 else None


def render_home(error: str = "") -> bytes:
    error_html = f'<div class="notice error">{esc(error)}</div>' if error else ""
    body = f"""
<div class="topbar">
  <div>
    <h1>Сравнение заявки и КП</h1>
  </div>
</div>
{error_html}
<section class="product-intro">
  <h2>Сервис сравнивает заявку с КП поставщиков и собирает итоговую Excel-сводку.</h2>
  <p>Загрузите ТЗ/заявку и счета поставщиков. Система найдет совпадающие позиции, покажет спорные строки для проверки и выделит лучшие и худшие цены в готовом отчете.</p>
  <div class="intro-points">
    <div class="intro-point"><b>Сопоставление:</b> <span>заявка и КП</span></div>
    <div class="intro-point"><b>Проверка:</b> <span>спорные строки</span></div>
    <div class="intro-point"><b>Excel:</b> <span>цены, сроки, подсветка</span></div>
  </div>
</section>
<section class="grid">
  <form class="panel upload-panel" action="/process" method="post" enctype="multipart/form-data" data-upload-form novalidate>
    <div class="format-note">
      <span class="pill">Заявка: .xlsx</span>
      <span class="pill">КП: .xlsx или текстовый .pdf</span>
      <span class="pill">На выходе: Excel-сводка</span>
    </div>
    <div class="field">
      <label for="request">Заявка / ТЗ (.xlsx)</label>
      <div class="file-drop" data-drop-zone>
        <input class="file-input" id="request" name="request" type="file" accept=".xlsx" required>
        <span class="file-icon">+</span>
        <div>
          <span class="file-title">Перетащите файл заявки сюда</span>
          <span class="file-subtitle">или нажмите, чтобы выбрать .xlsx на компьютере</span>
          <div class="file-list" data-file-list="request">Файл пока не выбран.</div>
        </div>
      </div>
      <div class="hint">Файл со списком позиций, единицами и количеством.</div>
    </div>
    <div class="field">
      <label for="offers">КП / счета поставщиков (.xlsx, .pdf)</label>
      <div class="file-drop" data-drop-zone>
        <input class="file-input" id="offers" name="offers" type="file" accept=".xlsx,.pdf,.xls" multiple required>
        <span class="file-icon">+</span>
        <div>
          <span class="file-title">Перетащите КП или счета сюда</span>
          <span class="file-subtitle">или нажмите, чтобы выбрать несколько файлов</span>
          <div class="file-list" data-file-list="offers">Файлы пока не выбраны.</div>
        </div>
      </div>
      <div class="hint">Выберите два или больше КП/счетов. Подходят файлы .xlsx и текстовые PDF.</div>
    </div>
    <div class="actions">
      <button class="btn" type="submit" data-submit>Обработать файлы</button>
      <button class="btn stop" type="button" data-stop>Остановить обработку</button>
    </div>
    <div class="notice error" data-form-message style="display:none; margin-top:12px; margin-bottom:0"></div>
    <div class="processing" data-processing>
      <div class="processing-head">
        <strong>Обрабатываем файлы</strong>
        <span>извлекаем позиции, цены и сроки</span>
      </div>
      <div class="progress"><div class="progress-bar"></div></div>
      <div class="hint">Это может занять до минуты, если файлов много или PDF большой.</div>
    </div>
  </form>
</section>
"""
    return page("Сравнение заявки и КП", body)


def render_review(run_id: str) -> bytes:
    run_dir = RUNS_DIR / run_id
    request_items, matches, errors = load_state(run_dir)
    stats = stats_for(request_items, matches)
    review_rows = [
        (idx, match)
        for idx, match in enumerate(matches)
        if match.status != "auto" or not match.request_pos or match.reason
    ]
    request_options = "\n".join(
        f'<option value="{esc(item.pos)} - {esc(item.name[:110])}"></option>'
        for item in request_items
    )
    warning_html = "".join(f'<div class="notice warn">{esc(error)}</div>' for error in errors)
    rows_html = ""
    for idx, match in review_rows:
        selected_item = next((item for item in request_items if item.pos == match.request_pos), None)
        selected_value = f"{selected_item.pos} - {selected_item.name[:110]}" if selected_item else ""
        match_input = f"""
<input class="match-input" type="text" name="match_{idx}" list="request-options" value="{esc(selected_value)}" placeholder="Начните вводить название или номер позиции">
<div class="match-help">Оставьте пустым, если позицию КП не нужно сопоставлять.</div>"""
        status_class = "unmatched" if not match.request_pos else match.status
        if status_class not in {"auto", "review", "unmatched", "manual"}:
            status_class = "review"
        rows_html += f"""
<tr data-review-row>
  <td class="small"><span class="status {status_class}">{esc(status_label(match.status))}</span></td>
  <td>{match_input}</td>
  <td>{esc(match.supplier_item.supplier)}</td>
  <td class="small">{esc(match.supplier_item.row_no)}</td>
  <td>{esc(match.supplier_item.name)}</td>
  <td class="small">{esc(match.supplier_item.qty)} {esc(match.supplier_item.unit)}</td>
  <td class="small">{esc(match.supplier_item.price)}</td>
  <td>{esc(match.reason or "проверить совпадение")}</td>
</tr>"""

    if rows_html:
        review_html = f"""
<form action="/finalize/{esc(run_id)}" method="post">
  <datalist id="request-options">
    {request_options}
  </datalist>
  <div class="notice warn">В поле "Позиция заявки" начните вводить название, бренд или номер позиции. Сервис покажет подсказки из заявки.</div>
  <div class="table-wrap">
    <table class="review-table">
      <thead>
        <tr>
          <th class="small">Статус</th>
          <th>Позиция заявки</th>
          <th>Поставщик</th>
          <th class="small">Строка</th>
          <th>Позиция КП</th>
          <th class="small">Кол-во</th>
          <th class="small">Цена</th>
          <th>Причина</th>
        </tr>
      </thead>
      <tbody>{rows_html}</tbody>
    </table>
  </div>
  <div class="actions">
    <button class="btn" type="submit">Сформировать Excel</button>
    <a class="btn secondary" href="/download/{esc(run_id)}/review.xlsx" download>Скачать файл проверки</a>
    <a class="btn secondary" href="/">Новая обработка</a>
  </div>
</form>"""
    else:
        review_html = f"""
<div class="panel empty">Спорных строк нет. Можно сразу скачать итоговый Excel.</div>
<div class="actions">
  <a class="btn" href="/download/{esc(run_id)}/summary.xlsx" download>Скачать Excel</a>
  <a class="btn secondary" href="/">Новая обработка</a>
</div>"""

    body = f"""
<div class="topbar">
  <div>
    <h1>Проверка совпадений</h1>
    <p class="subtitle">Исправьте спорные строки и оставьте лишние позиции КП несопоставленными.</p>
  </div>
</div>
{warning_html}
<section class="panel" style="margin-bottom:16px">
  <div class="stats">
    <div class="stat"><b>{stats["request"]}</b><span>позиций заявки</span></div>
    <div class="stat"><b>{stats["offers"]}</b><span>строк КП</span></div>
    <div class="stat"><b>{stats["auto"]}</b><span>автоматически</span></div>
    <div class="stat"><b>{stats["review"] + stats["unmatched"]}</b><span>требуют внимания</span></div>
  </div>
</section>
{review_html}
"""
    return page("Проверка совпадений", body)


def render_done(run_id: str) -> bytes:
    body = f"""
<div class="topbar">
  <div>
    <h1>Excel готов</h1>
    <p class="subtitle">Ручные правки применены. Итоговая сводка сформирована.</p>
  </div>
</div>
<section class="panel">
  <div class="actions">
    <a class="btn" href="/download/{esc(run_id)}/summary.xlsx" download>Скачать итоговый Excel</a>
    <a class="btn secondary" href="/download/{esc(run_id)}/review.xlsx" download>Скачать файл проверки</a>
    <a class="btn secondary" href="/">Новая обработка</a>
  </div>
</section>
"""
    return page("Excel готов", body)


class AppHandler(BaseHTTPRequestHandler):
    server_version = "TZKP/0.1"

    def send_html(self, content: bytes, status: HTTPStatus = HTTPStatus.OK) -> None:
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def redirect(self, location: str) -> None:
        self.send_response(HTTPStatus.SEE_OTHER)
        self.send_header("Location", location)
        self.end_headers()

    def do_GET(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        path = parsed.path
        if path == "/":
            self.send_html(render_home())
            return
        if path.startswith("/review/"):
            run_id = safe_filename(unquote(path.removeprefix("/review/")), "run")
            run_dir = RUNS_DIR / run_id
            if not (run_dir / "state.json").exists():
                self.send_html(render_home("Обработка не найдена. Загрузите файлы заново."), HTTPStatus.NOT_FOUND)
                return
            self.send_html(render_review(run_id))
            return
        if path.startswith("/done/"):
            run_id = safe_filename(unquote(path.removeprefix("/done/")), "run")
            self.send_html(render_done(run_id))
            return
        if path.startswith("/download/"):
            self.serve_download(path)
            return
        self.send_html(render_home("Страница не найдена."), HTTPStatus.NOT_FOUND)

    def do_HEAD(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        if parsed.path.startswith("/download/"):
            self.serve_download(parsed.path, send_body=False)
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        if parsed.path == "/process":
            self.handle_process()
            return
        if parsed.path.startswith("/finalize/"):
            run_id = safe_filename(unquote(parsed.path.removeprefix("/finalize/")), "run")
            self.handle_finalize(run_id)
            return
        self.send_html(render_home("Неверный адрес формы."), HTTPStatus.NOT_FOUND)

    def handle_process(self) -> None:
        length = int(self.headers.get("Content-Length") or 0)
        if length > MAX_UPLOAD_SIZE:
            self.send_html(render_home("Файлы слишком большие для локальной MVP-версии."), HTTPStatus.REQUEST_ENTITY_TOO_LARGE)
            return

        uploads = parse_multipart_upload(self.headers, self.rfile.read(length))
        request_fields = uploads.get("request", [])
        offer_fields = uploads.get("offers", [])
        request_field = request_fields[0] if request_fields else None
        if request_field is None or not request_field.filename:
            self.send_html(render_home("Загрузите файл заявки .xlsx."), HTTPStatus.BAD_REQUEST)
            return
        offer_fields = [field for field in offer_fields if field.filename]
        if len(offer_fields) < 2:
            self.send_html(render_home("Загрузите минимум два КП или счета."), HTTPStatus.BAD_REQUEST)
            return

        run_id = uuid.uuid4().hex[:12]
        run_dir = RUNS_DIR / run_id
        upload_dir = run_dir / "uploads"
        upload_dir.mkdir(parents=True, exist_ok=True)

        request_path = upload_dir / safe_filename(request_field.filename, "request.xlsx")
        with request_path.open("wb") as dst:
            shutil.copyfileobj(request_field.file, dst)

        offer_paths: list[Path] = []
        for idx, field in enumerate(offer_fields, start=1):
            target = upload_dir / safe_filename(field.filename, f"offer_{idx}")
            with target.open("wb") as dst:
                shutil.copyfileobj(field.file, dst)
            offer_paths.append(target)

        try:
            request_items = read_request_xlsx(request_path)
        except Exception as exc:  # noqa: BLE001
            self.send_html(render_home(f"Не удалось прочитать заявку: {exc}"), HTTPStatus.BAD_REQUEST)
            return

        supplier_items: list[SupplierItem] = []
        errors: list[str] = []
        for offer_path in offer_paths:
            try:
                supplier_items.extend(read_offer(offer_path))
            except Exception as exc:  # noqa: BLE001
                errors.append(str(exc))

        if not supplier_items:
            self.send_html(render_home("Не удалось извлечь позиции КП. Проверьте форматы файлов."), HTTPStatus.BAD_REQUEST)
            return

        matches = build_matches(request_items, supplier_items)
        write_review(run_dir / "review.xlsx", matches, request_items)
        write_final(run_dir / "summary.xlsx", request_items, matches)
        save_state(run_dir, request_items, matches, errors)
        self.redirect(f"/review/{run_id}")

    def handle_finalize(self, run_id: str) -> None:
        run_dir = RUNS_DIR / run_id
        if not (run_dir / "state.json").exists():
            self.send_html(render_home("Обработка не найдена. Загрузите файлы заново."), HTTPStatus.NOT_FOUND)
            return
        length = int(self.headers.get("Content-Length") or 0)
        raw_body = self.rfile.read(length).decode("utf-8", errors="replace")
        form = parse_qs(raw_body)
        request_items, matches, errors = load_state(run_dir)
        updated: list[Match] = []
        for idx, match in enumerate(matches):
            field_name = f"match_{idx}"
            if field_name in form:
                request_pos = resolve_request_pos(form[field_name][0], request_items)
                if request_pos:
                    updated.append(Match(match.supplier_item, request_pos, match.score, "manual", "подтверждено на проверке"))
                else:
                    updated.append(Match(match.supplier_item, None, match.score, "unmatched", "оставлено без сопоставления"))
            else:
                updated.append(match)

        write_review(run_dir / "review.xlsx", updated, request_items)
        write_final(run_dir / "summary.xlsx", request_items, updated)
        save_state(run_dir, request_items, updated, errors)
        self.redirect(f"/done/{run_id}")

    def serve_download(self, path: str, send_body: bool = True) -> None:
        parts = [part for part in path.split("/") if part]
        if len(parts) != 3:
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        _, run_id_raw, filename_raw = parts
        run_id = safe_filename(unquote(run_id_raw), "run")
        filename = safe_filename(unquote(filename_raw), "file")
        if filename not in {"summary.xlsx", "review.xlsx"}:
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        file_path = RUNS_DIR / run_id / filename
        if not file_path.exists():
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        data = file_path.read_bytes()
        content_type = mimetypes.guess_type(file_path.name)[0] or "application/octet-stream"
        download_name = f"{run_id}_{filename}"
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        self.send_header("Content-Disposition", f'attachment; filename="{download_name}"')
        self.end_headers()
        if send_body:
            self.wfile.write(data)

    def log_message(self, format: str, *args) -> None:  # noqa: A002
        sys.stderr.write("%s - %s\n" % (self.address_string(), format % args))


def main() -> None:
    host = os.environ.get("HOST") or ("0.0.0.0" if os.environ.get("PORT") else "127.0.0.1")
    port = int(os.environ.get("PORT", "8000"))
    if len(sys.argv) > 1:
        port = int(sys.argv[1])
    RUNS_DIR.mkdir(parents=True, exist_ok=True)
    server = ThreadingHTTPServer((host, port), AppHandler)
    print(f"Сервис запущен: http://{host}:{port}")
    print("Остановить: Ctrl+C")
    server.serve_forever()


if __name__ == "__main__":
    main()
