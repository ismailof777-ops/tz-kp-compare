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
import threading
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
    get_ai_warnings,
    match_score,
    read_offer,
    read_request_xlsx,
    status_label,
    write_final,
    write_review,
)


ROOT = Path(__file__).resolve().parent
RUNS_DIR = ROOT / "outputs" / "runs"
MAX_UPLOAD_SIZE = 80 * 1024 * 1024
AI_JOBS: dict[str, dict[str, object]] = {}
AI_JOBS_LOCK = threading.Lock()


class UploadedFile:
    def __init__(self, filename: str, content: bytes):
        self.filename = filename
        self.file = io.BytesIO(content)


CSS = """
:root {
  color-scheme: light;
  --bg: #f6f8fb;
  --panel: #ffffff;
  --panel-soft: #f8fafc;
  --text: #0f172a;
  --muted: #64748b;
  --line: #e2e8f0;
  --line-strong: #cbd5e1;
  --accent: #0f766e;
  --accent-strong: #0b5f59;
  --accent-soft: #ecfdf5;
  --blue: #2563eb;
  --green: #e5f6eb;
  --red: #fdeceb;
  --yellow: #fff6d6;
  --shadow: 0 18px 45px rgba(15, 23, 42, 0.06);
  --shadow-soft: 0 8px 24px rgba(15, 23, 42, 0.05);
  --font-sans: "Segoe UI", Inter, system-ui, -apple-system, BlinkMacSystemFont, "Helvetica Neue", Arial, sans-serif;
  --font-mono: ui-monospace, SFMono-Regular, Consolas, "Liberation Mono", "Courier New", monospace;
}
* { box-sizing: border-box; }
button,
input,
select,
textarea {
  font: inherit;
}
body {
  margin: 0;
  background: var(--bg);
  color: var(--text);
  font: 14px/1.5 var(--font-sans);
  font-variant-numeric: tabular-nums;
}
.site-header {
  width: min(1200px, calc(100% - 40px));
  margin: 0 auto;
  padding: 22px 0 0;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 18px;
}
.brand {
  display: inline-flex;
  align-items: center;
  gap: 10px;
  color: var(--text);
  font-weight: 800;
  text-decoration: none;
}
.brand-mark {
  width: 32px;
  height: 32px;
  border-radius: 10px;
  background: linear-gradient(135deg, var(--accent), #164e63);
  box-shadow: 0 10px 22px rgba(15, 118, 110, 0.18);
}
.header-links {
  display: flex;
  align-items: center;
  gap: 16px;
  color: var(--muted);
  font-size: 14px;
}
.header-links a {
  color: var(--muted);
  text-decoration: none;
}
.header-links a:hover {
  color: var(--accent);
}
.shell {
  width: min(1200px, calc(100% - 40px));
  margin: 0 auto;
  padding: 44px 0 44px;
}
.topbar {
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 16px;
  margin-bottom: 26px;
}
h1 {
  margin: 0;
  font-size: 42px;
  line-height: 1.12;
  letter-spacing: 0;
}
.subtitle {
  margin: 14px 0 0;
  color: var(--muted);
  max-width: 850px;
  font-size: 18px;
  line-height: 1.55;
}
.hero-badge {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  margin-top: 18px;
  padding: 8px 12px;
  border: 1px solid var(--line);
  border-radius: 999px;
  background: #fff;
  color: var(--muted);
  font-size: 13px;
  box-shadow: var(--shadow-soft);
}
.home-layout {
  display: grid;
  grid-template-columns: minmax(320px, 0.9fr) minmax(520px, 1.2fr);
  gap: 24px;
  align-items: start;
}
.product-intro {
  height: fit-content;
  align-self: start;
  position: sticky;
  top: 24px;
  padding: 24px;
  border: 1px solid var(--line);
  border-radius: 20px;
  background: #ffffff;
  box-shadow: var(--shadow);
}
.product-intro h2 {
  margin: 0;
  font-size: 22px;
  line-height: 1.2;
  letter-spacing: 0;
}
.product-intro p {
  margin: 12px 0 0;
  color: var(--muted);
  max-width: 620px;
  font-size: 15px;
}
.intro-points {
  display: grid;
  gap: 12px;
  margin-top: 20px;
}
.intro-point {
  border: 1px solid var(--line);
  border-radius: 16px;
  background: var(--panel-soft);
  padding: 14px 14px 14px 52px;
  color: var(--muted);
  font-size: 13px;
  position: relative;
}
.intro-point::before {
  content: attr(data-step);
  position: absolute;
  left: 14px;
  top: 14px;
  width: 26px;
  height: 26px;
  border-radius: 999px;
  background: #dff7f3;
  color: var(--accent);
  display: inline-flex;
  align-items: center;
  justify-content: center;
  font-weight: 800;
  font-size: 12px;
}
.intro-point b {
  display: block;
  margin: 0 0 2px;
  color: var(--text);
}
.intro-point span {
  display: block;
  color: var(--muted);
}
.panel {
  background: var(--panel);
  border: 1px solid var(--line);
  border-radius: 20px;
  padding: 30px;
  box-shadow: var(--shadow);
}
.grid {
  display: grid;
  grid-template-columns: minmax(0, 1fr);
  gap: 16px;
}
.upload-panel {
  width: 100%;
  align-self: start;
}
.field { margin-bottom: 24px; }
.field-head {
  margin-bottom: 12px;
}
.field-head h3 {
  margin: 0;
  font-size: 18px;
  line-height: 1.25;
}
.field-head p {
  margin: 6px 0 0;
  color: var(--muted);
  font-size: 14px;
}
label {
  display: block;
  font-weight: 700;
  margin-bottom: 8px;
}
select,
input[type="text"],
input[type="search"] {
  width: 100%;
  border: 1px solid var(--line);
  border-radius: 12px;
  background: #fff;
  padding: 11px 12px;
  min-height: 44px;
  color: var(--text);
  outline: none;
}
select,
input[type="text"],
input[type="search"] { font: inherit; }
select:focus,
input[type="text"]:focus,
input[type="search"]:focus {
  border-color: #99f6e4;
  box-shadow: 0 0 0 4px rgba(15, 118, 110, 0.1);
}
.match-input {
  min-width: 300px;
  text-overflow: ellipsis;
}
.match-help {
  color: var(--muted);
  font-size: 12px;
}
.match-actions {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 8px;
  margin-top: 7px;
}
.clear-match {
  border: 1px solid #fecaca;
  border-radius: 999px;
  background: #fff7f7;
  color: #991b1b;
  padding: 6px 10px;
  font-size: 12px;
  line-height: 1.2;
  font-weight: 700;
  cursor: pointer;
}
.clear-match:hover {
  background: #fee2e2;
}
.suggestions {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  margin-top: 10px;
}
.suggestion {
  border: 1px solid var(--line);
  border-radius: 999px;
  background: #fff;
  color: var(--text);
  padding: 7px 10px;
  font-size: 12px;
  line-height: 1.2;
  font-weight: 700;
  cursor: pointer;
}
.suggestion-label {
  color: var(--muted);
  font-weight: 600;
}
.suggestion:hover {
  border-color: #99f6e4;
  background: #f0fdfa;
  color: var(--accent-strong);
}
.with-tooltip {
  cursor: help;
}
.file-drop {
  position: relative;
  overflow: hidden;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 10px;
  min-height: 178px;
  border: 1px dashed var(--line-strong);
  border-radius: 18px;
  background: var(--panel-soft);
  padding: 26px;
  text-align: center;
  transition: border-color .15s ease, background .15s ease, box-shadow .15s ease;
}
.file-drop:hover,
.file-drop.is-dragover {
  border-color: var(--accent);
  background: #f0fdfa;
  box-shadow: 0 0 0 4px rgba(15, 118, 110, 0.08);
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
  width: 58px;
  height: 58px;
  border-radius: 999px;
  background: #dff7f3;
  color: var(--accent);
  font-size: 24px;
  line-height: 1;
  font-weight: 700;
}
.file-title {
  display: block;
  font-weight: 700;
  font-size: 16px;
}
.file-subtitle {
  display: block;
  margin-top: 2px;
  color: var(--muted);
  font-size: 13px;
}
.file-pick {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-height: 36px;
  margin-top: 4px;
  padding: 0 14px;
  border: 1px solid var(--line);
  border-radius: 999px;
  background: #fff;
  color: var(--accent);
  font-weight: 700;
  font-size: 13px;
}
.file-support {
  color: var(--muted);
  font-size: 12px;
}
.file-list {
  position: relative;
  z-index: 5;
  pointer-events: auto;
  margin-top: 12px;
  color: var(--muted);
  font-size: 13px;
}
.files-list {
  max-height: 292px;
  overflow-y: auto;
  padding-right: 4px;
}
.file-list.is-filled {
  color: var(--text);
}
.file-summary {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  margin-bottom: 8px;
  color: var(--text);
  font-weight: 700;
}
.file-list ul {
  margin: 0;
  padding-left: 0;
  list-style: none;
}
.file-list li {
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto;
  gap: 12px;
  margin: 8px 0;
  padding: 12px;
  border: 1px solid var(--line);
  border-radius: 14px;
  background: #fff;
  overflow-wrap: anywhere;
}
.file-name {
  min-width: 0;
  overflow-wrap: anywhere;
  color: var(--text);
  font-weight: 700;
}
.file-meta {
  display: flex;
  gap: 8px;
  flex-wrap: wrap;
  margin-top: 3px;
  color: var(--muted);
  font-size: 12px;
}
.file-status {
  display: inline-flex;
  align-items: center;
  min-height: 22px;
  padding: 0 8px;
  border-radius: 999px;
  background: #dcfce7;
  color: #166534;
  font-size: 12px;
  font-weight: 700;
}
.file-actions {
  display: flex;
  align-items: center;
  gap: 6px;
}
.file-remove {
  flex: 0 0 auto;
  min-height: 30px;
  border: 1px solid #fecaca;
  border-radius: 999px;
  background: #fff5f5;
  color: #8a2420;
  padding: 0 10px;
  font: inherit;
  font-size: 12px;
  cursor: pointer;
}
.file-replace {
  flex: 0 0 auto;
  min-height: 30px;
  border: 1px solid var(--line);
  border-radius: 999px;
  background: #fff;
  color: var(--accent);
  padding: 0 10px;
  font: inherit;
  font-size: 12px;
  cursor: pointer;
}
.file-replace:hover {
  background: var(--accent-soft);
}
.file-remove:hover {
  background: #ffe9e9;
}
.hint {
  margin-top: 6px;
  color: var(--muted);
  font-size: 13px;
}
.report-list {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 8px;
  margin: 18px 0 0;
  padding: 0;
  list-style: none;
}
.report-title {
  margin-top: 24px;
  font-size: 16px;
}
.report-list li {
  border: 1px solid var(--line);
  border-radius: 999px;
  background: #fff;
  padding: 7px 10px;
  color: var(--muted);
  font-size: 12px;
}
.consent-field {
  display: flex;
  gap: 9px;
  align-items: flex-start;
  margin: 4px 0 18px;
  color: var(--text);
  font-size: 13px;
  padding: 11px 12px;
  border: 1px solid var(--line);
  border-radius: 16px;
  background: var(--panel-soft);
}
.consent-field input {
  flex: 0 0 auto;
  width: 16px;
  height: 16px;
  margin: 2px 0 0;
}
.consent-field label {
  margin: 0;
  font-weight: 400;
}
a {
  color: var(--accent);
}
a:hover {
  color: var(--accent-strong);
}
.stepper {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 10px;
  margin: 0 0 26px;
}
.step {
  position: relative;
  display: flex;
  align-items: center;
  gap: 10px;
  min-height: 54px;
  padding: 10px 12px;
  border: 1px solid var(--line);
  border-radius: 16px;
  background: var(--panel-soft);
  color: var(--muted);
  font-size: 13px;
  font-weight: 700;
}
.step-number {
  width: 28px;
  height: 28px;
  border-radius: 999px;
  background: #fff;
  border: 1px solid var(--line);
  color: var(--accent);
  display: inline-flex;
  align-items: center;
  justify-content: center;
  font-weight: 800;
}
.step.is-active {
  border-color: #99f6e4;
  background: #f0fdfa;
  color: var(--text);
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
  min-height: 46px;
  padding: 0 18px;
  border: 1px solid var(--accent);
  border-radius: 14px;
  background: var(--accent);
  color: #fff;
  font-weight: 700;
  text-decoration: none;
  cursor: pointer;
}
.btn.primary-wide {
  width: 100%;
  min-height: 52px;
  font-size: 15px;
}
.cta-note {
  margin: 9px 0 0;
  color: var(--muted);
  font-size: 13px;
  text-align: center;
}
.btn,
.btn:visited,
.btn:hover,
.btn:focus {
  color: #fff;
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
.btn.secondary,
.btn.secondary:visited,
.btn.secondary:hover,
.btn.secondary:focus {
  color: var(--accent);
}
.btn.secondary:hover { background: var(--accent-soft); }
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
.review-summary {
  display: grid;
  grid-template-columns: repeat(7, minmax(112px, 1fr));
  gap: 12px;
  margin: 4px 0 18px;
}
.summary-item {
  display: grid;
  gap: 8px;
  min-height: 86px;
  border: 1px solid var(--line);
  border-radius: 18px;
  background: #fff;
  padding: 14px;
  color: var(--muted);
  font-size: 13px;
  box-shadow: var(--shadow-soft);
}
.summary-item b {
  color: var(--text);
  font-size: 24px;
  line-height: 1;
  font-variant-numeric: tabular-nums;
  white-space: nowrap;
}
.summary-item.attention {
  border-color: #fde68a;
  background: #fffbeb;
}
.summary-item.primary {
  border-color: #99f6e4;
  background: #f0fdfa;
}
.notice {
  border-radius: 16px;
  border: 1px solid var(--line);
  padding: 14px 16px;
  margin-bottom: 18px;
  background: #fff;
  color: var(--muted);
  box-shadow: var(--shadow-soft);
}
.notice.warn { border-color: #fde68a; background: #fffbeb; color: #854d0e; }
.notice.error { border-color: #fecaca; background: #fff1f2; color: #991b1b; }
.legal-page {
  max-width: 920px;
  margin: 0 auto;
}
.legal-page h2 {
  margin: 0 0 12px;
  font-size: 18px;
}
.legal-page h3 {
  margin: 22px 0 8px;
  padding-top: 18px;
  border-top: 1px solid var(--line);
  font-size: 15px;
}
.legal-page h2 + p,
.legal-page h3:first-of-type {
  border-top: 0;
  padding-top: 0;
}
.legal-page p,
.legal-page li {
  color: var(--muted);
  line-height: 1.65;
}
.legal-page ul {
  margin: 8px 0 0;
  padding-left: 20px;
}
.site-footer {
  width: min(1200px, calc(100% - 40px));
  margin: -18px auto 0;
  padding: 0 0 30px;
  color: var(--muted);
  font-size: 13px;
  display: flex;
  gap: 10px;
  align-items: center;
  flex-wrap: wrap;
}
.site-footer span::after {
  content: "·";
  margin-left: 10px;
  color: #98a2b3;
}
.cookie-banner {
  position: fixed;
  left: 16px;
  right: 16px;
  bottom: 16px;
  z-index: 50;
  display: none;
  align-items: center;
  justify-content: space-between;
  gap: 14px;
  max-width: 920px;
  margin: 0 auto;
  padding: 13px 14px;
  border: 1px solid var(--line);
  border-radius: 8px;
  background: #fff;
  box-shadow: 0 16px 36px rgba(16, 24, 40, 0.14);
}
.cookie-banner.is-visible {
  display: flex;
}
.cookie-text {
  min-width: 0;
  color: var(--muted);
  font-size: 13px;
}
.cookie-text b {
  display: block;
  margin-bottom: 3px;
  color: var(--text);
}
.cookie-actions {
  display: flex;
  gap: 8px;
  flex: 0 0 auto;
  flex-wrap: wrap;
}
.cookie-actions .btn {
  min-height: 34px;
  padding: 0 11px;
  font-size: 13px;
}
.table-wrap {
  width: 100%;
  max-width: 100%;
  overflow: auto;
  border: 1px solid var(--line);
  border-radius: 20px;
  background: #fff;
  box-shadow: var(--shadow);
}
.review-form {
  display: grid;
  gap: 18px;
  min-width: 0;
}
.review-card {
  min-width: 0;
  border: 1px solid var(--line);
  border-radius: 22px;
  background: #fff;
  padding: 20px;
  box-shadow: var(--shadow);
}
.review-card-head {
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 16px;
  margin-bottom: 16px;
}
.review-card-head h2 {
  margin: 0;
  font-size: 22px;
  line-height: 1.2;
}
.review-card-head p {
  margin: 6px 0 0;
  color: var(--muted);
  font-size: 14px;
  line-height: 1.5;
}
.review-badge {
  flex: 0 0 auto;
  display: inline-flex;
  align-items: center;
  min-height: 32px;
  padding: 0 12px;
  border: 1px solid #99f6e4;
  border-radius: 999px;
  background: #f0fdfa;
  color: var(--accent-strong);
  font-size: 13px;
  font-weight: 800;
}
.review-actionbar {
  min-width: 0;
  display: grid;
  grid-template-columns: minmax(260px, 1fr) auto;
  align-items: center;
  gap: 16px;
  border: 1px solid var(--line);
  border-radius: 20px;
  background: #fff;
  padding: 16px;
  box-shadow: var(--shadow);
}
.review-actionbar .actions {
  margin-top: 0;
  justify-content: flex-end;
  flex-wrap: nowrap;
}
.action-copy {
  color: var(--muted);
  font-size: 13px;
  line-height: 1.45;
}
.action-copy b {
  display: block;
  margin-bottom: 2px;
  color: var(--text);
  font-size: 14px;
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
.review-table {
  font-family: var(--font-sans);
}
.review-table td.small,
.review-table th.small,
.review-table .match-help,
.review-table .suggestion,
.summary-item b {
  font-family: var(--font-mono);
}
.summary-item b {
  font-family: var(--font-sans);
}
table {
  border-collapse: collapse;
  width: 100%;
  min-width: 1080px;
  font-size: 13px;
}
th, td {
  border-bottom: 1px solid var(--line);
  padding: 14px 12px;
  text-align: left;
  vertical-align: top;
}
th {
  position: sticky;
  top: 0;
  background: #f8fafc;
  z-index: 1;
  font-size: 12px;
  color: #475569;
  font-weight: 800;
  text-transform: uppercase;
  letter-spacing: 0;
}
tbody tr:hover td {
  background: #f8fafc;
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
  box-shadow: 10px 0 18px rgba(15, 23, 42, 0.05);
}
.review-table th:nth-child(1),
.review-table th:nth-child(2) {
  z-index: 4;
  background: #f8fafc;
}
.review-table .match-input {
  min-width: 0;
}
.status {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  max-width: 126px;
  border: 1px solid transparent;
  border-radius: 999px;
  padding: 6px 9px;
  font-size: 11px;
  line-height: 1.2;
  font-weight: 800;
  text-align: center;
}
.status.review { background: #fff6d8; border-color: #f1d88d; color: #6f5200; }
.status.unmatched { background: #fdeaea; border-color: #f2b9b7; color: #8a2420; }
.status.auto { background: #e8f6ec; border-color: #bfe5c8; color: #236336; }
.status.manual { background: #eef4fa; border-color: #cbdcea; color: #244f73; }
.status.service { background: #f0f2f4; border-color: #d9e0ea; color: #667085; }
.muted { color: var(--muted); }
.danger { background: var(--red); }
.empty {
  padding: 26px;
  text-align: center;
  color: var(--muted);
}
.empty h2 {
  margin: 12px 0 6px;
  color: var(--text);
  font-size: 22px;
}
.empty p {
  margin: 0 auto;
  max-width: 620px;
  color: var(--muted);
}
.ready-panel {
  display: grid;
  gap: 16px;
}
.ready-badge {
  display: inline-flex;
  width: fit-content;
  align-items: center;
  min-height: 32px;
  padding: 0 12px;
  border-radius: 999px;
  background: #dcfce7;
  color: #166534;
  font-weight: 800;
  font-size: 13px;
}
.ready-panel h2 {
  margin: 0;
  font-size: 24px;
}
.ready-panel p {
  margin: 0;
  color: var(--muted);
  font-size: 15px;
}
.ready-stats {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 10px;
}
.ready-stat {
  border: 1px solid var(--line);
  border-radius: 16px;
  background: var(--panel-soft);
  padding: 12px;
  color: var(--muted);
  font-size: 13px;
}
.ready-stat b {
  display: block;
  margin-top: 4px;
  color: var(--text);
  font: 800 22px/1 var(--font-mono);
}
.processing {
  display: none;
  margin-top: 24px;
  border: 1px solid var(--line);
  border-radius: 18px;
  background: var(--panel-soft);
  padding: 18px;
  box-shadow: var(--shadow-soft);
}
.processing.is-visible { display: block; }
.processing-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 14px;
  margin-bottom: 14px;
  font-weight: 700;
}
.processing-title {
  display: flex;
  align-items: center;
  gap: 10px;
  min-width: 0;
}
.processing-icon {
  width: 34px;
  height: 34px;
  border-radius: 999px;
  background: #dff7f3;
  color: var(--accent);
  display: inline-flex;
  align-items: center;
  justify-content: center;
  font-weight: 900;
}
.processing-head span {
  color: var(--accent);
  font-weight: 800;
}
.progress {
  overflow: hidden;
  height: 10px;
  border-radius: 999px;
  background: #dbeafe;
}
.progress-bar {
  width: 0%;
  height: 100%;
  border-radius: 999px;
  background: linear-gradient(90deg, var(--accent), #14b8a6);
  transition: width .25s ease;
}
.processing-note {
  margin-top: 12px;
  color: var(--muted);
  font-size: 13px;
  line-height: 1.55;
}
.processing-note[data-processing-message] {
  color: var(--text);
  font-weight: 600;
}
.ai-progress {
  display: none;
  width: min(520px, 100%);
  margin-top: 12px;
  border: 1px solid #b9dce3;
  border-radius: 8px;
  background: #f1fbfc;
  padding: 12px;
}
.ai-progress.is-visible { display: block; }
.ai-progress-line {
  display: flex;
  justify-content: space-between;
  gap: 12px;
  margin-bottom: 8px;
  color: var(--muted);
  font-size: 13px;
}
.ai-progress-track {
  overflow: hidden;
  height: 10px;
  border-radius: 999px;
  background: #d5e8ef;
}
.ai-progress-fill {
  width: 0%;
  height: 100%;
  border-radius: 999px;
  background: var(--accent);
  transition: width .25s ease;
}
.ai-progress-message {
  margin-top: 8px;
  color: var(--muted);
  font-size: 12px;
}
.summary-help {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 12px;
  margin: 0 0 18px;
  color: var(--muted);
  font-size: 13px;
  line-height: 1.55;
}
.summary-help div {
  border: 1px solid var(--line);
  border-radius: 18px;
  background: #fff;
  padding: 14px;
  box-shadow: var(--shadow-soft);
}
.summary-help b {
  color: var(--text);
  font-weight: 700;
}
@media (max-width: 1180px) {
  .review-summary {
    grid-template-columns: repeat(4, minmax(0, 1fr));
  }
  .review-actionbar {
    grid-template-columns: 1fr;
  }
  .review-actionbar .actions {
    justify-content: flex-start;
    flex-wrap: wrap;
  }
}
@media (max-width: 840px) {
  .site-header { width: min(100% - 24px, 1180px); padding-top: 14px; }
  .header-links { gap: 10px; }
  .shell { width: min(100% - 24px, 1180px); padding-top: 26px; }
  .topbar, .grid, .home-layout { display: block; }
  .home-layout { display: flex; flex-direction: column; }
  .grid { order: 1; }
  .product-intro { order: 2; position: static; }
  .upload-panel { max-width: none; }
  .panel { padding: 18px; border-radius: 18px; }
  .product-intro { padding: 18px; margin-top: 14px; }
  .product-intro h2 { font-size: 20px; }
  .intro-points { display: grid; }
  .stepper { grid-template-columns: 1fr; }
  .file-drop { min-height: 164px; padding: 20px; }
  .file-list li { grid-template-columns: 1fr; }
  .file-actions { justify-content: flex-start; }
  .report-list { grid-template-columns: 1fr; }
  .cookie-banner {
    align-items: stretch;
    flex-direction: column;
  }
  .cookie-actions .btn {
    flex: 1 1 130px;
  }
  .review-summary { grid-template-columns: repeat(2, minmax(0, 1fr)); }
  .ready-stats { grid-template-columns: 1fr; }
  .summary-help { grid-template-columns: 1fr; }
  .review-tools { grid-template-columns: 1fr; }
  .review-count { padding-bottom: 0; }
  .review-card { padding: 16px; border-radius: 18px; }
  .review-card-head,
  .review-actionbar {
    display: grid;
    grid-template-columns: 1fr;
  }
  .review-actionbar .actions {
    justify-content: stretch;
  }
  .review-badge {
    width: fit-content;
  }
  h1 { font-size: 30px; }
  .subtitle { font-size: 15px; }
  .btn { width: 100%; }
  .site-footer { width: min(100% - 24px, 1180px); }
}
"""


def esc(value) -> str:
    return html.escape("" if value is None else str(value), quote=True)


def set_ai_job(run_id: str, **updates: object) -> None:
    with AI_JOBS_LOCK:
        current = AI_JOBS.setdefault(
            run_id,
            {
                "state": "idle",
                "current": 0,
                "total": 0,
                "percent": 0,
                "message": "",
                "redirect": f"/review/{run_id}",
            },
        )
        current.update(updates)


def get_ai_job(run_id: str) -> dict[str, object]:
    with AI_JOBS_LOCK:
        payload = dict(
            AI_JOBS.get(
                run_id,
                {
                    "state": "idle",
                    "current": 0,
                    "total": 0,
                    "percent": 0,
                    "message": "ИИ-проверка не запущена",
                    "redirect": f"/review/{run_id}",
                },
            )
        )
        payload["run_id"] = run_id
        return payload


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
  <header class="site-header">
    <a class="brand" href="/">
      <span class="brand-mark" aria-hidden="true"></span>
      <span>Сравнение КП</span>
    </a>
    <nav class="header-links" aria-label="Навигация">
      <a href="/#how-it-works">Помощь</a>
      <a href="/privacy">Политика</a>
    </nav>
  </header>
  <main class="shell">{body}</main>
  <footer class="site-footer">
    <span>© approvemoscow.ru</span>
    <a href="/privacy">Политика обработки персональных данных</a>
  </footer>
  <div class="cookie-banner" data-cookie-banner role="region" aria-label="Уведомление о cookies">
    <div class="cookie-text">
      <b>Cookies и технические данные</b>
      <span>Сервис использует только необходимые технические данные для работы сайта. Аналитика и рекламные cookies не подключены.</span>
      <a href="/privacy">Подробнее</a>
    </div>
    <div class="cookie-actions">
      <button class="btn secondary" type="button" data-cookie-choice="necessary">Только необходимые</button>
      <button class="btn" type="button" data-cookie-choice="accepted">Понятно</button>
    </div>
  </div>
  <script>
    const cookieBanner = document.querySelector('[data-cookie-banner]');
    if (cookieBanner) {{
      const storageKey = 'approvemoscow_cookie_choice';
      let savedChoice = '';
      try {{
        savedChoice = window.localStorage.getItem(storageKey) || '';
      }} catch (error) {{
        savedChoice = '';
      }}
      if (!savedChoice) {{
        cookieBanner.classList.add('is-visible');
      }}
      cookieBanner.querySelectorAll('[data-cookie-choice]').forEach((button) => {{
        button.addEventListener('click', () => {{
          try {{
            window.localStorage.setItem(storageKey, button.dataset.cookieChoice || 'accepted');
          }} catch (error) {{
            // Если localStorage недоступен, просто скрываем баннер на текущей странице.
          }}
          cookieBanner.classList.remove('is-visible');
        }});
      }});
    }}
    const uploadForm = document.querySelector('[data-upload-form]');
    if (uploadForm) {{
      const requestInput = uploadForm.querySelector('input[name="request"]');
      const offersInput = uploadForm.querySelector('input[name="offers"]');
      const consentInput = uploadForm.querySelector('input[name="privacy_consent"]');
      const formMessage = uploadForm.querySelector('[data-form-message]');
      const submitButton = uploadForm.querySelector('[data-submit]');
      const stopButton = uploadForm.querySelector('[data-stop]');
      const processing = uploadForm.querySelector('[data-processing]');
      const processingFill = uploadForm.querySelector('[data-processing-fill]');
      const processingPercent = uploadForm.querySelector('[data-processing-percent]');
      const processingMessage = uploadForm.querySelector('[data-processing-message]');
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
          submitButton.textContent = 'Сформировать Excel-отчет';
        }}
        if (stopButton) stopButton.classList.remove('is-visible');
        if (processing) processing.classList.remove('is-visible');
        if (processingFill) processingFill.style.width = '0%';
        if (processingPercent) processingPercent.textContent = '0%';
      }};
      const setUploadProgress = (data) => {{
        const percent = Math.max(0, Math.min(100, Number(data.percent || 0)));
        if (processing) processing.classList.add('is-visible');
        if (processingFill) processingFill.style.width = percent + '%';
        if (processingPercent) processingPercent.textContent = percent + '%';
        if (processingMessage) processingMessage.textContent = data.message || 'Обработка файлов';
      }};
      const pollUploadProgress = async (runId) => {{
        const response = await fetch('/progress/' + encodeURIComponent(runId), {{ cache: 'no-store' }});
        const data = await response.json();
        setUploadProgress(data);
        if (data.state === 'done') {{
          window.location.href = data.redirect || ('/review/' + runId);
          return;
        }}
        if (data.state === 'error') {{
          resetProcessingState();
          showMessage(data.message || 'Не удалось обработать файлы.');
          return;
        }}
        setTimeout(() => pollUploadProgress(runId), 700);
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
        const fileSize = (size) => {{
          if (size >= 1024 * 1024) return (size / 1024 / 1024).toFixed(1).replace('.0', '') + ' МБ';
          if (size >= 1024) return Math.round(size / 1024) + ' КБ';
          return size + ' Б';
        }};
        const fileExt = (name) => {{
          const part = String(name || '').split('.').pop() || '';
          return part ? part.toUpperCase() : 'ФАЙЛ';
        }};
        const summary = document.createElement('div');
        summary.className = 'file-summary';
        summary.textContent = input.multiple
          ? 'Загружено файлов: ' + files.length
          : 'Файл выбран';
        target.appendChild(summary);
        const list = document.createElement('ul');
        files.forEach((file, index) => {{
          const item = document.createElement('li');
          const details = document.createElement('div');
          const name = document.createElement('span');
          name.className = 'file-name';
          name.textContent = file.name;
          const meta = document.createElement('div');
          meta.className = 'file-meta';
          const format = document.createElement('span');
          format.textContent = fileExt(file.name);
          const size = document.createElement('span');
          size.textContent = fileSize(file.size);
          const status = document.createElement('span');
          status.className = 'file-status';
          status.textContent = 'Готово';
          meta.appendChild(format);
          meta.appendChild(size);
          meta.appendChild(status);
          details.appendChild(name);
          details.appendChild(meta);
          const actions = document.createElement('div');
          actions.className = 'file-actions';
          if (!input.multiple) {{
            const replace = document.createElement('button');
            replace.className = 'file-replace';
            replace.type = 'button';
            replace.textContent = 'Заменить';
            replace.addEventListener('click', (event) => {{
              event.preventDefault();
              event.stopPropagation();
              input.click();
            }});
            actions.appendChild(replace);
          }}
          const remove = document.createElement('button');
          remove.className = 'file-remove';
          remove.type = 'button';
          remove.textContent = 'Удалить';
          remove.addEventListener('click', (event) => {{
            event.preventDefault();
            event.stopPropagation();
            removeInputFile(input, index);
          }});
          actions.appendChild(remove);
          item.appendChild(details);
          item.appendChild(actions);
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
        if (!consentInput || !consentInput.checked) {{
          showMessage('Подтвердите согласие на обработку персональных данных.');
          consentInput && consentInput.focus();
          return;
        }}
        if (formMessage) formMessage.style.display = 'none';
        activeController = new AbortController();
        if (submitButton) {{
          submitButton.disabled = true;
          submitButton.textContent = 'Формируем Excel-отчет...';
        }}
        if (stopButton) stopButton.classList.add('is-visible');
        if (processing) processing.classList.add('is-visible');
        try {{
          const response = await fetch(uploadForm.action, {{
            method: 'POST',
            headers: {{ 'Accept': 'application/json' }},
            body: new FormData(uploadForm),
            signal: activeController.signal,
          }});
          const contentType = response.headers.get('content-type') || '';
          if (contentType.includes('application/json')) {{
            const data = await response.json();
            if (!response.ok || data.state === 'error') {{
              resetProcessingState();
              showMessage(data.message || 'Не удалось обработать файлы.');
              return;
            }}
            setUploadProgress(data);
            pollUploadProgress(data.run_id);
            return;
          }}
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
    const requestOptions = Array.from(document.querySelectorAll('#request-options option'));
    const updateMatchTitle = (input) => {{
      const option = requestOptions.find((item) => item.value === input.value);
      input.title = option?.dataset.full || input.value || 'Оставить без сопоставления';
    }};
    document.querySelectorAll('.match-input').forEach((input) => {{
      updateMatchTitle(input);
      input.addEventListener('input', () => updateMatchTitle(input));
      input.addEventListener('change', () => updateMatchTitle(input));
    }});
    document.querySelectorAll('[data-suggest]').forEach((button) => {{
      button.addEventListener('click', () => {{
        const cell = button.closest('td');
        const input = cell?.querySelector('.match-input');
        if (!input) return;
        input.value = button.dataset.suggest || '';
        updateMatchTitle(input);
        input.dispatchEvent(new Event('change', {{ bubbles: true }}));
      }});
    }});
    document.querySelectorAll('[data-clear-match]').forEach((button) => {{
      button.addEventListener('click', () => {{
        const cell = button.closest('td');
        const input = cell?.querySelector('.match-input');
        if (!input) return;
        input.value = '';
        updateMatchTitle(input);
        input.dispatchEvent(new Event('change', {{ bubbles: true }}));
      }});
    }});
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


def multipart_has_field(headers, body: bytes, expected_name: str) -> bool:
    content_type = headers.get("Content-Type", "")
    if not content_type.lower().startswith("multipart/form-data"):
        return False

    message = BytesParser(policy=email_policy).parsebytes(
        b"Content-Type: "
        + content_type.encode("utf-8", errors="ignore")
        + b"\r\nMIME-Version: 1.0\r\n\r\n"
        + body
    )
    for part in message.iter_parts():
        if part.get_content_disposition() != "form-data":
            continue
        if part.get_param("name", header="content-disposition") == expected_name:
            return True
    return False


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
        invoice_no=clean_text(data.get("invoice_no")),
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
    comparable = [match for match in matches if match.status != "service"]
    matched = sum(1 for match in comparable if match.request_pos)
    return {
        "request": len(request_items),
        "offers": len(matches),
        "comparable": len(comparable),
        "service": sum(1 for match in matches if match.status == "service"),
        "auto": sum(1 for match in comparable if match.status == "auto" and match.request_pos),
        "review": sum(1 for match in comparable if match.status == "review" and match.request_pos),
        "unmatched": sum(1 for match in comparable if not match.request_pos),
        "matched": matched,
    }


def top_request_suggestions(match: Match, request_items: list[RequestItem], limit: int = 5) -> list[tuple[RequestItem, float]]:
    suggestions: list[tuple[RequestItem, float]] = []
    for item in request_items:
        score, _reason = match_score(item.name, match.supplier_item.name)
        if item.pos == match.request_pos:
            score = max(score, match.score)
        suggestions.append((item, score))
    suggestions.sort(key=lambda pair: pair[1], reverse=True)
    return suggestions[:limit]


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
    <h1>Сравнение ТЗ и КП поставщиков</h1>
    <p class="subtitle">Загрузите заявку и коммерческие предложения. Сервис сопоставит позиции, найдет минимальные цены и сформирует Excel-сводку для проверки закупки.</p>
    <div class="hero-badge">XLSX · PDF · Excel-отчет</div>
  </div>
</div>
{error_html}
<section class="home-layout">
<section class="product-intro info-panel" id="how-it-works">
  <h2>Как это работает</h2>
  <div class="intro-points">
    <div class="intro-point" data-step="1"><b>Загрузите заявку</b> <span>Excel-файл со списком позиций, количеством и единицами измерения.</span></div>
    <div class="intro-point" data-step="2"><b>Добавьте КП поставщиков</b> <span>Можно загрузить несколько Excel/PDF файлов от разных поставщиков.</span></div>
    <div class="intro-point" data-step="3"><b>Получите Excel-отчет</b> <span>Сервис выделит совпадения, минимальные цены и спорные строки.</span></div>
  </div>
  <h2 class="report-title">В отчете будет</h2>
  <ul class="report-list">
    <li>позиции</li>
    <li>поставщики</li>
    <li>цены за единицу</li>
    <li>итоговые суммы</li>
    <li>минимальные предложения</li>
    <li>спорные совпадения</li>
  </ul>
</section>
<section class="grid">
  <form class="panel upload-panel" action="/process" method="post" enctype="multipart/form-data" data-upload-form novalidate>
    <div class="stepper" aria-label="Этапы обработки">
      <div class="step is-active"><span class="step-number">1</span><span>Заявка</span></div>
      <div class="step"><span class="step-number">2</span><span>КП поставщиков</span></div>
      <div class="step"><span class="step-number">3</span><span>Excel-отчет</span></div>
    </div>
    <div class="field">
      <div class="field-head">
        <h3>Заявка / ТЗ</h3>
        <p>Загрузите Excel-файл со списком позиций, количеством и единицами измерения.</p>
      </div>
      <div class="file-drop" data-drop-zone>
        <input class="file-input" id="request" name="request" type="file" accept=".xlsx" required>
        <span class="file-icon">↑</span>
        <span class="file-title">Перетащите файл заявки сюда</span>
        <span class="file-subtitle">или выберите .xlsx на компьютере</span>
        <span class="file-pick">Выбрать файл</span>
        <span class="file-support">Поддерживается .xlsx</span>
      </div>
      <div class="file-list" data-file-list="request">Файл пока не выбран.</div>
    </div>
    <div class="field">
      <div class="field-head">
        <h3>КП / счета поставщиков</h3>
        <p>Добавьте два или больше КП или счетов от поставщиков.</p>
      </div>
      <div class="file-drop" data-drop-zone>
        <input class="file-input" id="offers" name="offers" type="file" accept=".xlsx,.pdf,.xls" multiple required>
        <span class="file-icon">↑</span>
        <span class="file-title">Перетащите КП или счета сюда</span>
        <span class="file-subtitle">или выберите несколько файлов</span>
        <span class="file-pick">Выбрать файлы</span>
        <span class="file-support">Поддерживаются .xlsx и текстовые PDF</span>
      </div>
      <div class="file-list files-list" data-file-list="offers">Файлы пока не выбраны.</div>
    </div>
    <div class="consent-field">
      <input id="privacy_consent" name="privacy_consent" type="checkbox" value="yes" required>
      <label for="privacy_consent">Я согласен с <a href="/privacy" target="_blank" rel="noopener">обработкой персональных данных</a></label>
    </div>
    <div class="actions">
      <button class="btn primary-wide" type="submit" data-submit>Сформировать Excel-отчет</button>
      <button class="btn stop" type="button" data-stop>Остановить обработку</button>
    </div>
    <div class="cta-note">Обычно обработка занимает 1–3 минуты.</div>
    <div class="notice error" data-form-message style="display:none; margin-top:12px; margin-bottom:0"></div>
    <div class="processing" data-processing>
      <div class="processing-head">
        <div class="processing-title">
          <span class="processing-icon">↻</span>
          <strong>Формирование Excel-отчета</strong>
        </div>
        <span data-processing-percent>0%</span>
      </div>
      <div class="progress"><div class="progress-bar" data-processing-fill></div></div>
      <div class="processing-note" data-processing-message>Загружаем файлы и готовим их к распознаванию.</div>
      <div class="processing-note">Сервис извлекает позиции, цены и сроки, затем сопоставляет товары с учетом ИИ. Если файлов много или PDF большой, обработка может занять несколько минут.</div>
    </div>
  </form>
</section>
</section>
"""
    return page("Сравнение заявки и КП", body)


def render_review(run_id: str) -> bytes:
    run_dir = RUNS_DIR / run_id
    request_items, matches, errors = load_state(run_dir)
    stats = stats_for(request_items, matches)
    match_percent = round((stats["matched"] / stats["comparable"]) * 100, 1) if stats["comparable"] else 0
    review_rows = [
        (idx, match)
        for idx, match in enumerate(matches)
        if match.status != "service" and (match.status != "auto" or not match.request_pos or match.reason)
    ]
    request_options = "\n".join(
        f'<option value="{esc(item.pos)} - {esc(item.name[:110])}" data-full="{esc(item.pos)} - {esc(item.name)}"></option>'
        for item in request_items
    )
    warning_html = "".join(f'<div class="notice warn">{esc(error)}</div>' for error in errors)
    rows_html = ""
    for idx, match in review_rows:
        selected_item = next((item for item in request_items if item.pos == match.request_pos), None)
        selected_value = f"{selected_item.pos} - {selected_item.name[:110]}" if selected_item else ""
        selected_title = f"{selected_item.pos} - {selected_item.name}" if selected_item else "Оставить без сопоставления"
        suggestion_buttons = ""
        suggestions = top_request_suggestions(match, request_items)
        if suggestions:
            suggestion_buttons = '<div class="suggestions" aria-label="Варианты сопоставления">'
            for suggestion, score in suggestions:
                value = f"{suggestion.pos} - {suggestion.name[:110]}"
                full = f"{suggestion.pos} - {suggestion.name}"
                suggestion_buttons += (
                    f'<button class="suggestion" type="button" data-suggest="{esc(value)}" '
                    f'title="{esc(full)}">{esc(suggestion.pos)} <span class="suggestion-label">совпадение {round(score * 100)}%</span></button>'
                )
            suggestion_buttons += "</div>"
        match_input = f"""
<input class="match-input with-tooltip" type="text" name="match_{idx}" list="request-options" value="{esc(selected_value)}" title="{esc(selected_title)}" placeholder="Начните вводить название или номер позиции">
<div class="match-actions">
  <button class="clear-match" type="button" data-clear-match>Не сопоставлять</button>
  <span class="match-help">Если строка КП лишняя или не относится к заявке.</span>
</div>
{suggestion_buttons}"""
        status_class = "unmatched" if not match.request_pos else match.status
        if status_class not in {"auto", "review", "unmatched", "manual", "service"}:
            status_class = "review"
        rows_html += f"""
<tr data-review-row>
  <td class="small"><span class="status {status_class}">{esc(status_label(match.status))}</span></td>
  <td>{match_input}</td>
  <td>{esc(match.supplier_item.supplier)}</td>
  <td class="small">{esc(match.supplier_item.row_no)}</td>
  <td class="with-tooltip" title="{esc(match.supplier_item.name)}">{esc(match.supplier_item.name)}</td>
  <td class="small">{esc(match.supplier_item.qty)} {esc(match.supplier_item.unit)}</td>
  <td class="small">{esc(match.supplier_item.price)}</td>
  <td>{esc(match.reason or "проверить совпадение")}</td>
</tr>"""

    if rows_html:
        review_html = f"""
<form class="review-form" action="/finalize/{esc(run_id)}" method="post">
  <datalist id="request-options">
    {request_options}
  </datalist>
  <div class="review-card">
    <div class="review-card-head">
      <div>
        <h2>Строки для проверки</h2>
        <p>Подтвердите предложенную позицию заявки или выберите подходящий вариант из подсказок под полем.</p>
      </div>
      <span class="review-badge">{len(review_rows)} к проверке</span>
    </div>
    <div class="notice warn">Процент рядом с подсказкой показывает похожесть позиции КП на позицию заявки. Слабые совпадения отправляются на проверку, а не подтверждаются автоматически.</div>
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
  </div>
  <div class="review-actionbar">
    <div class="action-copy">
      <b>После проверки будет сформирован итоговый Excel-отчет</b>
      Сервис применит ручные правки, сохранит спорные решения и пересоберет сводку по поставщикам.
    </div>
    <div class="actions">
      <button class="btn" type="submit">Сформировать Excel-отчет</button>
      <a class="btn secondary" href="/download/{esc(run_id)}/review.xlsx" download>Скачать файл проверки</a>
      <a class="btn secondary" href="/">Новая обработка</a>
    </div>
  </div>
</form>"""
    else:
        review_html = f"""
<div class="review-card empty">
  <span class="ready-badge">Готово</span>
  <h2>Спорных строк нет</h2>
  <p>Все товарные позиции сопоставлены автоматически. Можно сразу скачать итоговый Excel-отчет.</p>
</div>
<div class="review-actionbar">
  <div class="action-copy">
    <b>Excel-отчет готов к скачиванию</b>
    Итоговая сводка уже сформирована по данным заявки и КП поставщиков.
  </div>
  <div class="actions">
    <a class="btn" href="/download/{esc(run_id)}/summary.xlsx" download>Скачать Excel-отчет</a>
    <a class="btn secondary" href="/">Новая обработка</a>
  </div>
</div>"""

    body = f"""
<div class="topbar">
  <div>
    <h1>Проверка совпадений</h1>
    <p class="subtitle">Исправьте спорные строки и оставьте лишние позиции КП несопоставленными.</p>
  </div>
</div>
{warning_html}
<div class="review-summary" aria-label="Краткая сводка обработки">
  <div class="summary-item"><span>Заявка</span><b>{stats["request"]}</b></div>
  <div class="summary-item"><span>КП</span><b>{stats["offers"]}</b></div>
  <div class="summary-item"><span>Товарные</span><b>{stats["comparable"]}</b></div>
  <div class="summary-item primary"><span>Сопоставлено</span><b>{match_percent}%</b></div>
  <div class="summary-item"><span>Авто</span><b>{stats["auto"]}</b></div>
  <div class="summary-item attention"><span>К проверке</span><b>{stats["review"] + stats["unmatched"]}</b></div>
  <div class="summary-item"><span>Доставка/услуги</span><b>{stats["service"]}</b></div>
</div>
<div class="summary-help">
  <div><b>Товарные</b> - строки КП с материалами и товарами. Именно они участвуют в проценте сопоставления.</div>
  <div><b>Доставка/услуги</b> - доставка, разгрузка и транспортные услуги. Они сохраняются в Excel отдельным листом и не смешиваются с товарами.</div>
  <div><b>К проверке</b> - строки, где система предложила совпадение, но человеку лучше подтвердить выбор перед финальным Excel.</div>
  <div><b>Подсказки</b> - варианты под полем показывают возможные совпадения. Чем выше процент, тем вероятнее совпадение.</div>
</div>
{review_html}
"""
    return page("Проверка совпадений", body)


def render_done(run_id: str) -> bytes:
    stats_html = ""
    try:
        request_items, matches, _errors = load_state(RUNS_DIR / run_id)
        stats = stats_for(request_items, matches)
        suppliers = {match.supplier_item.supplier for match in matches if match.supplier_item.supplier}
        stats_html = f"""
  <div class="ready-stats">
    <div class="ready-stat">Позиции заявки<b>{stats["request"]}</b></div>
    <div class="ready-stat">Спорные совпадения<b>{stats["review"] + stats["unmatched"]}</b></div>
    <div class="ready-stat">Поставщики<b>{len(suppliers)}</b></div>
  </div>"""
    except Exception:
        stats_html = ""
    body = f"""
<div class="topbar">
  <div>
    <h1>Excel готов</h1>
    <p class="subtitle">Ручные правки применены. Итоговая сводка сформирована.</p>
  </div>
</div>
<section class="panel ready-panel">
  <span class="ready-badge">Готово</span>
  <div>
    <h2>Excel-отчет готов</h2>
    <p>Файл можно скачать и использовать для проверки закупки.</p>
  </div>
  {stats_html}
  <div class="actions">
    <a class="btn" href="/download/{esc(run_id)}/summary.xlsx" download>Скачать Excel-отчет</a>
    <a class="btn secondary" href="/download/{esc(run_id)}/review.xlsx" download>Скачать файл проверки</a>
    <a class="btn secondary" href="/">Новая обработка</a>
  </div>
</section>
"""
    return page("Excel готов", body)


def render_privacy() -> bytes:
    body = """
<div class="topbar">
  <div>
    <h1>Политика обработки персональных данных</h1>
    <p class="subtitle">Документ описывает, какие данные обрабатывает сервис approvemoscow.ru при сравнении заявки/ТЗ и КП поставщиков.</p>
  </div>
</div>
<section class="panel legal-page">
  <h2>1. Общие положения</h2>
  <p>Настоящая политика применяется к сайту approvemoscow.ru и сервису сравнения заявки/ТЗ с коммерческими предложениями поставщиков. Используя сервис и загружая файлы, пользователь подтверждает согласие с условиями обработки данных.</p>

  <h3>2. Какие данные могут обрабатываться</h3>
  <ul>
    <li>сведения из загружаемых файлов заявки, ТЗ, КП, счетов и PDF-документов;</li>
    <li>контактные данные, реквизиты организаций, ФИО, телефоны, адреса электронной почты, если они содержатся в файлах;</li>
    <li>технические данные запроса: IP-адрес, дата и время обращения, сведения о браузере и системные журналы сервера.</li>
  </ul>

  <h3>3. Цели обработки</h3>
  <p>Данные используются для загрузки и распознавания документов, сопоставления позиций заявки и КП, формирования Excel-отчета, диагностики ошибок и обеспечения работоспособности сервиса.</p>

  <h3>4. Правовые основания и согласие</h3>
  <p>Обработка выполняется на основании согласия пользователя, выраженного отметкой чекбокса перед отправкой файлов, а также для исполнения действия, запрошенного пользователем в сервисе.</p>

  <h3>5. Передача третьим лицам</h3>
  <p>Для сопоставления позиций сервис может использовать внешние API обработки текста. В такие API могут передаваться фрагменты данных из загруженных документов, необходимые для сравнения позиций. Сервис не продает персональные данные и не передает их третьим лицам для рекламных целей.</p>

  <h3>6. Хранение и защита</h3>
  <p>Файлы и результаты обработки хранятся на сервере сервиса в объеме, необходимом для выполнения обработки и скачивания результата. Доступ к серверу ограничивается техническими средствами администрирования.</p>

  <h3>7. Cookies и аналитика</h3>
  <p>Сайт показывает уведомление о cookies и технических данных. Сервис не подключает рекламную аналитику и не использует cookies для отслеживания пользователей. Выбор в уведомлении сохраняется локально в браузере, чтобы не показывать баннер повторно.</p>

  <h3>8. Права пользователя</h3>
  <p>Пользователь может запросить информацию об обработке данных, уточнение или удаление загруженных материалов и результатов обработки, если такие данные сохраняются на сервере.</p>

  <h3>9. Контакты</h3>
  <p>Для вопросов по обработке персональных данных используйте контактный канал владельца сайта approvemoscow.ru.</p>
</section>
"""
    return page("Политика обработки персональных данных", body)


class AppHandler(BaseHTTPRequestHandler):
    server_version = "TZKP/0.1"

    def send_html(self, content: bytes, status: HTTPStatus = HTTPStatus.OK) -> None:
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def send_json(self, payload: dict, status: HTTPStatus = HTTPStatus.OK) -> None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def wants_json(self) -> bool:
        return "application/json" in self.headers.get("Accept", "").lower()

    def send_process_error(self, message: str, status: HTTPStatus = HTTPStatus.BAD_REQUEST) -> None:
        if self.wants_json():
            self.send_json({"state": "error", "message": message}, status)
            return
        self.send_html(render_home(message), status)

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
        if path == "/privacy":
            self.send_html(render_privacy())
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
        if path.startswith("/progress/"):
            run_id = safe_filename(unquote(path.removeprefix("/progress/")), "run")
            self.send_json(get_ai_job(run_id))
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
        if parsed.path.startswith("/rerun-ai/"):
            run_id = safe_filename(unquote(parsed.path.removeprefix("/rerun-ai/")), "run")
            self.handle_rerun_ai(run_id)
            return
        self.send_html(render_home("Неверный адрес формы."), HTTPStatus.NOT_FOUND)

    def handle_process(self) -> None:
        length = int(self.headers.get("Content-Length") or 0)
        if length > MAX_UPLOAD_SIZE:
            self.send_process_error("Файлы слишком большие для локальной MVP-версии.", HTTPStatus.REQUEST_ENTITY_TOO_LARGE)
            return

        body = self.rfile.read(length)
        if not multipart_has_field(self.headers, body, "privacy_consent"):
            self.send_process_error("Подтвердите согласие на обработку персональных данных.")
            return

        uploads = parse_multipart_upload(self.headers, body)
        request_fields = uploads.get("request", [])
        offer_fields = uploads.get("offers", [])
        request_field = request_fields[0] if request_fields else None
        if request_field is None or not request_field.filename:
            self.send_process_error("Загрузите файл заявки .xlsx.")
            return
        offer_fields = [field for field in offer_fields if field.filename]
        if len(offer_fields) < 2:
            self.send_process_error("Загрузите минимум два КП или счета.")
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

        def update_progress(percent: int, message: str) -> None:
            set_ai_job(
                run_id,
                state="running",
                current=percent,
                total=100,
                percent=percent,
                message=message,
                redirect=f"/review/{run_id}",
            )

        def update_ai_progress(current: int, total: int, message: str) -> None:
            ai_percent = round((current / total) * 42) if total else 0
            update_progress(min(90, 48 + ai_percent), message)

        def worker() -> None:
            try:
                update_progress(8, "Читаем файл заявки")
                request_items = read_request_xlsx(request_path)

                supplier_items: list[SupplierItem] = []
                errors: list[str] = []
                for idx, offer_path in enumerate(offer_paths, start=1):
                    update_progress(12 + round((idx - 1) / max(1, len(offer_paths)) * 24), f"Извлекаем позиции КП: файл {idx} из {len(offer_paths)}")
                    try:
                        parsed_items = read_offer(offer_path)
                        if parsed_items:
                            supplier_items.extend(parsed_items)
                        else:
                            errors.append(f"{offer_path.name}: позиции КП не найдены. Возможно, это скан или нестандартный PDF.")
                    except Exception as exc:  # noqa: BLE001
                        errors.append(str(exc))

                if not supplier_items:
                    details = " ".join(errors[:5])
                    raise ValueError(f"Не удалось извлечь позиции КП. {details}".strip())

                update_progress(42, "Сопоставляем товары по названиям, синонимам и единицам")
                matches = build_matches(request_items, supplier_items, progress_callback=update_ai_progress)
                errors.extend(get_ai_warnings())

                update_progress(93, "Формируем файл проверки и итоговый Excel")
                write_review(run_dir / "review.xlsx", matches, request_items)
                write_final(run_dir / "summary.xlsx", request_items, matches)
                save_state(run_dir, request_items, matches, errors)

                comparable = [match for match in matches if match.status != "service"]
                total = len(comparable)
                matched = sum(1 for match in comparable if match.request_pos)
                percent = round((matched / total) * 100, 1) if total else 0
                set_ai_job(
                    run_id,
                    state="done",
                    current=100,
                    total=100,
                    percent=100,
                    message=f"Готово. Сопоставлено {percent}% товарных строк",
                    redirect=f"/review/{run_id}",
                )
            except Exception as exc:  # noqa: BLE001
                set_ai_job(
                    run_id,
                    state="error",
                    percent=0,
                    message=f"Ошибка обработки: {exc}",
                    redirect="/",
                )

        update_progress(3, "Файлы загружены. Запускаем обработку")
        threading.Thread(target=worker, daemon=True).start()
        self.send_json(get_ai_job(run_id))

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

    def handle_rerun_ai(self, run_id: str) -> None:
        run_dir = RUNS_DIR / run_id
        if not (run_dir / "state.json").exists():
            self.send_json({"state": "error", "message": "Обработка не найдена"}, HTTPStatus.NOT_FOUND)
            return
        existing = get_ai_job(run_id)
        if existing.get("state") == "running":
            self.send_json(existing)
            return

        def update_progress(current: int, total: int, message: str) -> None:
            percent = round((current / total) * 100) if total else 0
            set_ai_job(
                run_id,
                state="running",
                current=current,
                total=total,
                percent=percent,
                message=message,
                redirect=f"/review/{run_id}",
            )

        def worker() -> None:
            try:
                set_ai_job(run_id, state="running", current=0, total=0, percent=0, message="Готовим файл к ИИ-проверке")
                request_items, matches, errors = load_state(run_dir)
                supplier_items = [match.supplier_item for match in matches]
                updated = build_matches(request_items, supplier_items, progress_callback=update_progress)
                clean_errors = [error for error in errors if not error.startswith("DeepSeek ")]
                clean_errors.extend(get_ai_warnings())
                write_review(run_dir / "review.xlsx", updated, request_items)
                write_final(run_dir / "summary.xlsx", request_items, updated)
                save_state(run_dir, request_items, updated, clean_errors)
                comparable = [match for match in updated if match.status != "service"]
                total = len(comparable)
                matched = sum(1 for match in comparable if match.request_pos)
                percent = round((matched / total) * 100, 1) if total else 0
                set_ai_job(
                    run_id,
                    state="done",
                    current=total,
                    total=total,
                    percent=100,
                    message=f"Готово. Сопоставлено {percent}%",
                    redirect=f"/review/{run_id}",
                )
            except Exception as exc:  # noqa: BLE001
                set_ai_job(run_id, state="error", percent=0, message=f"Ошибка ИИ-проверки: {exc}", redirect=f"/review/{run_id}")

        set_ai_job(run_id, state="running", current=0, total=0, percent=0, message="Запускаем ИИ-проверку", redirect=f"/review/{run_id}")
        threading.Thread(target=worker, daemon=True).start()
        self.send_json(get_ai_job(run_id))

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
        try:
            sys.stderr.write("%s - %s\n" % (self.address_string(), format % args))
        except Exception:
            pass


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
