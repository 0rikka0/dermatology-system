# -*- coding: utf-8 -*-
"""
Dermatology Pathology Reporting System - Web Server & RESTful API
Uses standard library ThreadingHTTPServer.
"""
import os
import sys
import re
import html
import json
import threading
import urllib.parse
import mimetypes
import socket
import subprocess
from datetime import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler
from socketserver import ThreadingMixIn

from app.config import load_config, save_config, update_config
from app.instance_lock import APP_ID, remove_runtime_info
from app.path_manager import (
    ensure_workspace_structure,
    get_database_path,
    get_patient_word_path,
    get_final_report_path,
    get_patient_photo_dir,
    ensure_patient_photo_dir,
    rename_patient_files_and_folders,
    find_clinical_photo_folder,
    get_clinical_photo_root_dir,
    get_archive_root_dir,
    archive_patient_report,
    save_special_draft_queue,
    sync_review_overwrites,
    locate_patient_all_paths,
    delete_patient_all_files,
    sync_status_file_operations,
    sync_updated_report_to_special_archives,
    cleanup_stale_output_reports,
    find_patient_review_folder,
    sync_review_photos_to_making_and_siblings,
    find_patient_archive_dir,
    find_patient_archive_path,
    find_patient_ihc_archive_path,
    create_ihc_report_package,
    archive_ihc_report,
    revoke_ihc_archive_files,
    delete_ihc_report_files,
    revoke_patient_archive_files,
    revoke_patient_approval_files,
    get_signing_sheet_dir
)
from app.database import (
    init_db,
    get_patient,
    check_duplicate,
    get_next_pathology_no,
    insert_patient,
    update_patient,
    get_today_patients,
    search_patients,
    update_patient_report_making,
    mark_patient_archived,
    mark_patient_approved,
    update_patient_status,
    delete_patient,
    get_patients_for_report_making,
    get_patients_for_review,
    get_patients_for_ihc,
    get_ihc_report,
    upsert_ihc_report,
    mark_ihc_archived,
    revoke_ihc_archive_db,
    delete_ihc_report_and_clear_status,
    update_ihc_paths,
    get_output_regular_patients,
    get_output_special_approved_patients,
    get_output_ihc_patients,
    revoke_patient_archive_db,
    revoke_patient_approval_db,
    get_patients_for_signing_sheet,
    get_today_batch_archivable_patients
)
from app.template_manager import (
    load_text_templates,
    search_templates,
    recommend_templates,
    parse_photo_filename,
    format_pathology_diagnosis
)
from app.doc_generator import generate_report_file, update_report_file
from app.doc_preview import render_preview_html, extract_docx_details
from app.signing_sheet_generator import (
    generate_signing_sheet_excel,
    generate_signing_sheet_word,
    render_signing_sheet_html,
    compute_surgery_date_str
)

def normalize_report_date(val):
    """Normalize a date string ('YYYY-MM-DD' or 'YYYY年M月D日') into (cn_str, raw_str)."""
    s = str(val or "").strip()
    if not s:
        return "", ""
    m_iso = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})$", s)
    if m_iso:
        y, m, d = m_iso.groups()
        return f"{y}年{int(m):02d}月{int(d):02d}日", f"{y}-{int(m):02d}-{int(d):02d}"
    m_cn = re.match(r"^(\d{4})年(\d{1,2})月(\d{1,2})日?$", s)
    if m_cn:
        y, m, d = m_cn.groups()
        return f"{y}年{int(m):02d}月{int(d):02d}日", f"{y}-{int(m):02d}-{int(d):02d}"
    return s, ""


def _scan_photos_dir(photo_dir, image_count=1):
    """Scan a pathology photo directory and compute parsed photos, auto diagnosis and template recommendations."""
    photos = []
    parsed_items = []
    diagnoses_found = []

    if photo_dir and os.path.exists(photo_dir):
        try:
            entries = sorted(os.listdir(photo_dir))
        except Exception:
            entries = []
        for fn in entries:
            p_info = parse_photo_filename(fn)
            if p_info:
                full_p = os.path.join(photo_dir, fn)
                p_info["path"] = full_p
                p_info["url"] = f"/api/photo_file?path={urllib.parse.quote(full_p)}"
                photos.append(p_info)
                parsed_items.append((p_info["indices"], p_info["diagnosis"]))
                if p_info["diagnosis"]:
                    diagnoses_found.append(p_info["diagnosis"])

    auto_diag = format_pathology_diagnosis(parsed_items, image_count=image_count)
    recs = recommend_templates(diagnoses_found)
    return photos, auto_diag, recs


def _remap_selected_images_dir(selected_images, old_dir, new_dir):
    """Rewrite selected_images path prefixes when a photo directory moves."""
    if not old_dir or not new_dir or old_dir == new_dir or not isinstance(selected_images, list):
        return selected_images
    old_norm = os.path.normpath(old_dir)
    new_norm = os.path.normpath(new_dir)
    updated_imgs = []
    for item in selected_images:
        if isinstance(item, dict) and item.get("path"):
            item_copy = dict(item)
            p_norm = os.path.normpath(item_copy["path"])
            if p_norm.startswith(old_norm):
                item_copy["path"] = new_norm + p_norm[len(old_norm):]
            updated_imgs.append(item_copy)
        elif isinstance(item, str):
            p_norm = os.path.normpath(item)
            updated_imgs.append(new_norm + p_norm[len(old_norm):] if p_norm.startswith(old_norm) else item)
        else:
            updated_imgs.append(item)
    return updated_imgs


def _build_ihc_preview_data(patient, ihc):
    """Merge patient demographics with independent IHC report fields for preview/rendering."""
    doc_path = ihc.get("ihc_doc_path")
    if not doc_path or not os.path.exists(doc_path):
        doc_path = ihc.get("archive_file")
    prev_data = dict(patient)
    for k in ("report_doctor", "review_doctor", "report_date", "report_date_raw",
              "microscopic_desc", "pathology_diag", "selected_images", "image_count", "biopsy_sites"):
        val = ihc.get(k)
        if val not in (None, ""):
            prev_data[k] = val
    return prev_data, doc_path or ""


def _attach_review_photo_dir(work_dir, patient_dict):
    """Attach copied 特殊患者归档/(待审核|重切|完成审核)/.../病理照片 path info onto a patient dict."""
    if not patient_dict or not work_dir:
        return patient_dict
    info = find_patient_review_folder(
        work_dir,
        patient_dict.get("pathology_no", ""),
        patient_dict.get("name", ""),
        review_status=patient_dict.get("review_status", ""),
        is_pending=patient_dict.get("is_pending_review"),
        is_recut=patient_dict.get("is_needs_recut")
    )
    out = dict(patient_dict)
    out["review_photo_dir"] = info.get("photo_dir", "") if info else ""
    out["review_photo_category"] = info.get("category", "") if info else ""
    return out


def _canonicalize_review_selected_images(selected_images, making_photo_dir):
    """Ensure selected_images points to stable paths in 未归档报告/病理照片/ when matching files exist."""
    if not isinstance(selected_images, list) or not making_photo_dir:
        return selected_images
    canonical = []
    for item in selected_images:
        if isinstance(item, dict) and item.get("path"):
            item_copy = dict(item)
            bn = os.path.basename(item_copy["path"])
            cand = os.path.join(making_photo_dir, bn)
            if bn and os.path.exists(cand):
                item_copy["path"] = cand
            canonical.append(item_copy)
        elif isinstance(item, str) and item:
            bn = os.path.basename(item)
            cand = os.path.join(making_photo_dir, bn)
            canonical.append(cand if (bn and os.path.exists(cand)) else item)
        else:
            canonical.append(item)
    return canonical


def print_word_documents(file_paths, page1_only=True):
    """
    Print one or more .docx files via Microsoft Word (or WPS Writer fallback) COM automation.
    When page1_only=True, uses wdPrintFromTo (Range=3, From='1', To='1') so only Page 1 prints.
    Returns (printed_paths, failed_paths, error_message).
    """
    valid_paths = [os.path.normpath(p) for p in file_paths if p and os.path.exists(p)]
    if not valid_paths:
        return [], list(file_paths), "未找到可打印的 Word 文档文件"

    ps_script = r"""
$ErrorActionPreference = 'Stop'
$paths = @($env:DERM_PRINT_PATHS | ConvertFrom-Json)
$page1Only = ($env:DERM_PRINT_PAGE1 -eq '1')
$app = $null
foreach ($progId in @('Word.Application', 'KWPS.Application', 'wps.Application')) {
    try {
        $app = New-Object -ComObject $progId
        if ($app) { break }
    } catch {
        $app = $null
    }
}
if (-not $app) {
    Write-Output '{"printed":[],"failed":[],"error":"未检测到可用的 Word 或 WPS 打印组件"}'
    exit 2
}
$printed = @()
$failed = @()
$lastErr = ''
try {
    $app.Visible = $false
    $app.DisplayAlerts = 0
    $missing = [System.Type]::Missing
    foreach ($p in $paths) {
        $doc = $null
        try {
            $fn = [string]$p
            $cc = $false
            $ro = $true
            try {
                $doc = $app.Documents.Open([ref]$fn, [ref]$cc, [ref]$ro)
            } catch {
                $doc = $app.Documents.Open($p)
            }
            $bg = $false
            $ap = $false
            if ($page1Only) {
                $rng = 3
                $from = '1'
                $to = '1'
                $outFn = $missing
                try {
                    # wdPrintFromTo = 3, From='1', To='1'
                    $doc.PrintOut([ref]$bg, [ref]$ap, [ref]$rng, [ref]$outFn, [ref]$from, [ref]$to)
                } catch {
                    # Fallback positional call for WPS compatibility
                    $doc.PrintOut($false, $false, 3, '', '1', '1')
                }
            } else {
                try {
                    $doc.PrintOut([ref]$bg)
                } catch {
                    $doc.PrintOut($false)
                }
            }
            $printed += $p
        } catch {
            $failed += $p
            $lastErr = $_.Exception.Message
        } finally {
            if ($doc) {
                $saveChanges = $false
                try { $doc.Close([ref]$saveChanges) } catch { try { $doc.Close($false) } catch {} }
                try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null } catch {}
            }
        }
    }
} finally {
    try { $app.Quit() } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}
}
@{ printed = $printed; failed = $failed; error = $lastErr } | ConvertTo-Json -Compress
"""
    env = os.environ.copy()
    env["DERM_PRINT_PATHS"] = json.dumps(valid_paths, ensure_ascii=False)
    env["DERM_PRINT_PAGE1"] = "1" if page1_only else "0"
    timeout_sec = max(25, 12 + 8 * len(valid_paths))
    try:
        res = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_script],
            env=env,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout_sec
        )
        out_lines = [l.strip() for l in (res.stdout or "").splitlines() if l.strip().startswith("{")]
        if out_lines:
            parsed = json.loads(out_lines[-1])
            printed = parsed.get("printed") or []
            failed = parsed.get("failed") or []
            err = parsed.get("error") or ""
            return printed, failed, err
        err_out = (res.stderr or res.stdout or "").strip()
        return [], valid_paths, err_out or f"Word 打印进程退出码: {res.returncode}"
    except Exception as e:
        return [], valid_paths, str(e)

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WEB_DIR = os.path.join(BASE_DIR, "web")

class ThreadedHTTPServer(ThreadingMixIn, HTTPServer):
    daemon_threads = True
    allow_reuse_address = False

    def server_bind(self):
        if os.name == "nt" and hasattr(socket, "SO_EXCLUSIVEADDRUSE"):
            try:
                self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
            except OSError:
                pass
        super().server_bind()

class PathologyRequestHandler(SimpleHTTPRequestHandler):
    """Handles HTTP requests for Static files and REST API."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=WEB_DIR, **kwargs)

    def log_message(self, format, *args):
        # Concise logging with silent pythonw.exe (sys.stdout is None) protection
        try:
            if sys.stdout is not None:
                sys.stdout.write(f"[{datetime.now().strftime('%H:%M:%S')}] {args[0]} - {args[1]}\n")
        except Exception:
            pass

    def send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()
        self.wfile.write(body)

    def send_error_json(self, message, status=400):
        self.send_json({"success": False, "error": message}, status=status)

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def read_json_body(self):
        content_length = int(self.headers.get("Content-Length", 0))
        if content_length > 0:
            raw = self.rfile.read(content_length).decode("utf-8")
            return json.loads(raw)
        return {}

    def end_headers(self):
        # Prevent browser caching of static JS/CSS in local desktop application
        if hasattr(self, 'path') and any(self.path.endswith(ext) for ext in ['.js', '.css', '.html']):
            self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
            self.send_header("Pragma", "no-cache")
            self.send_header("Expires", "0")
        super().end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        query = urllib.parse.parse_qs(parsed.query, keep_blank_values=True)

        # Favicon icon handler (returns micro-svg icon to avoid 404)
        if path == "/favicon.ico":
            self.send_response(200)
            self.send_header("Content-Type", "image/svg+xml")
            self.send_header("Cache-Control", "max-age=86400")
            self.end_headers()
            svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><text y=".9em" font-size="90">🔬</text></svg>'
            self.wfile.write(svg.encode("utf-8"))
            return

        # API routing
        if path.startswith("/api/"):
            try:
                self.handle_api_get(path, query)
            except Exception as e:
                self.send_error_json(f"服务器内部错误: {str(e)}", status=500)
            return

        # Serve index.html with dynamic status detection
        if path in ["/", "", "/index.html"]:
            html_path = os.path.join(WEB_DIR, "index.html")
            if os.path.exists(html_path):
                try:
                    with open(html_path, "r", encoding="utf-8") as f:
                        html_content = f.read()

                    cfg = load_config()
                    is_initialized = cfg.get("initialized", False)
                    work_dir = cfg.get("work_dir", "")
                    rep_doc = cfg.get("report_doctor", "张医生")
                    rev_doc = cfg.get("review_doctor", "李主任")

                    # If not initialized, directly display setup wizard modal on page load
                    if not is_initialized:
                        html_content = html_content.replace(
                            'id="setup-wizard-modal" class="wizard-overlay" style="display: none;"',
                            'id="setup-wizard-modal" class="wizard-overlay" style="display: flex;"'
                        )
                        if work_dir:
                            html_content = re.sub(
                                r'id="wizard-work-dir" class="form-control" value="[^"]*"',
                                lambda _: f'id="wizard-work-dir" class="form-control" value="{html.escape(work_dir)}"',
                                html_content
                            )
                        if rep_doc:
                            html_content = re.sub(
                                r'id="wizard-report-doc" class="form-control" value="[^"]*"',
                                lambda _: f'id="wizard-report-doc" class="form-control" value="{html.escape(rep_doc)}"',
                                html_content
                            )
                        if rev_doc:
                            html_content = re.sub(
                                r'id="wizard-review-doc" class="form-control" value="[^"]*"',
                                lambda _: f'id="wizard-review-doc" class="form-control" value="{html.escape(rev_doc)}"',
                                html_content
                            )
                    else:
                        html_content = html_content.replace(
                            'id="setup-wizard-modal" class="wizard-overlay" style="display: flex;"',
                            'id="setup-wizard-modal" class="wizard-overlay" style="display: none;"'
                        )

                    body = html_content.encode("utf-8")
                    self.send_response(200)
                    self.send_header("Content-Type", "text/html; charset=utf-8")
                    self.send_header("Content-Length", str(len(body)))
                    self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
                    self.send_header("Pragma", "no-cache")
                    self.send_header("Expires", "0")
                    self.end_headers()
                    self.wfile.write(body)
                    return
                except Exception as e:
                    print(f"[Server] 动态渲染 index.html 失败: {e}")

            self.path = "/index.html"
            return super().do_GET()

        return super().do_GET()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        if path.startswith("/api/"):
            try:
                self.handle_api_post(path)
            except Exception as e:
                self.send_error_json(f"服务器处理失败: {str(e)}", status=500)
            return
        self.send_error(404, "Not Found")

    def do_PUT(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        if path.startswith("/api/"):
            try:
                self.handle_api_put(path)
            except Exception as e:
                self.send_error_json(f"服务器处理失败: {str(e)}", status=500)
            return
        self.send_error(404, "Not Found")

    def do_DELETE(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        if path.startswith("/api/"):
            try:
                self.handle_api_delete(path)
            except Exception as e:
                self.send_error_json(f"删除失败: {str(e)}", status=500)
            return
        self.send_error(404, "Not Found")

    # ----------------- API Handlers -----------------

    def handle_api_get(self, path, query):
        cfg = load_config()
        work_dir = cfg.get("work_dir")
        db_path = get_database_path(work_dir) if work_dir else None

        if path == "/api/status":
            today_str = datetime.now().strftime("%Y-%m-%d")
            self.send_json({
                "success": True,
                "app_id": APP_ID,
                "pid": os.getpid(),
                "initialized": cfg.get("initialized", False),
                "work_dir": cfg.get("work_dir", ""),
                "report_doctor": cfg.get("report_doctor", "张医生"),
                "review_doctor": cfg.get("review_doctor", "李主任"),
                "today": today_str
            })
            return

        if path == "/api/config":
            self.send_json({"success": True, "config": cfg})
            return

        # Serve local image file for WebUI thumbnail and preview
        if path == "/api/photo_file":
            p_file = query.get("path", [""])[0]
            if not p_file or not os.path.exists(p_file) or not os.path.isfile(p_file):
                self.send_error(404, "Photo not found")
                return

            # Security: Path traversal protection
            real_p = os.path.realpath(p_file).lower()
            allowed = False
            if work_dir and real_p.startswith(os.path.realpath(work_dir).lower()):
                allowed = True
            elif real_p.startswith(os.path.realpath(BASE_DIR).lower()):
                allowed = True

            if not allowed:
                self.send_error(403, "Access to path forbidden")
                return

            # Check file size limit (100MB)
            try:
                f_size = os.path.getsize(p_file)
            except Exception:
                f_size = 0
            if f_size > 100 * 1024 * 1024:
                self.send_error(413, "Photo file exceeds maximum limit (100MB)")
                return

            mime, _ = mimetypes.guess_type(p_file)
            mime = mime or "application/octet-stream"
            try:
                self.send_response(200)
                self.send_header("Content-Type", mime)
                self.send_header("Content-Length", str(f_size))
                self.send_header("Cache-Control", "max-age=3600")
                self.end_headers()
                with open(p_file, "rb") as f:
                    while True:
                        chunk = f.read(65536)
                        if not chunk:
                            break
                        self.wfile.write(chunk)
            except Exception as e:
                pass
            return

        # Text templates search
        if path == "/api/templates":
            cat = query.get("category", ["全部"])[0] or "全部"
            kw = query.get("keyword", [""])[0]
            tmpls = search_templates(keyword=kw, category=cat)
            self.send_json({"success": True, "templates": tmpls, "count": len(tmpls)})
            return

        # Text templates recommendation
        if path == "/api/templates/recommend":
            raw_diags = query.get("diagnoses", [])
            diag_list = []
            for item in raw_diags:
                for part in re.split(r"[,;；\n]+", str(item)):
                    part = part.strip()
                    if part:
                        diag_list.append(part)
            recs = recommend_templates(diag_list)
            self.send_json({"success": True, "recommended": recs, "count": len(recs)})
            return

        # Next pathology sequence number
        if path == "/api/next_no":
            if not cfg.get("initialized"):
                self.send_error_json("系统尚未初始化工作目录")
                return
            prefix = query.get("prefix", [""])[0]
            date_str = query.get("date", [None])[0]
            next_no = get_next_pathology_no(db_path, prefix=prefix, date_str=date_str)
            self.send_json({"success": True, "next_no": next_no})
            return

        # Check pathology number duplicate
        if path == "/api/check_no":
            if not cfg.get("initialized"):
                self.send_error_json("系统尚未初始化工作目录")
                return
            no = query.get("no", [""])[0].strip()
            if not no:
                self.send_json({"success": True, "exists": False})
                return
            patient = check_duplicate(db_path, no)
            self.send_json({
                "success": True,
                "exists": patient is not None,
                "patient": patient
            })
            return

        # Today's / Entry Patient List
        if path == "/api/patients/today":
            if not cfg.get("initialized"):
                self.send_json({"success": True, "patients": []})
                return
            if "report_date" in query or "submission_date" in query or "keyword" in query or "filter_mode" in query:
                rep_date = query.get("report_date", [""])[0]
                sub_date = query.get("submission_date", [""])[0]
                kw = query.get("keyword", [""])[0]
                patients = get_today_patients(db_path, report_date=rep_date, submission_date=sub_date, keyword=kw)
            else:
                patients = get_today_patients(db_path)
            self.send_json({"success": True, "patients": patients, "count": len(patients)})
            return

        # Patients list for Report Making workspace
        if path == "/api/patients/report_making":
            if not cfg.get("initialized"):
                self.send_json({"success": True, "patients": []})
                return
            d_raw = query.get("date", [None])[0]
            rep_date = query.get("report_date", [None])[0]
            sub_date = query.get("submission_date", [None])[0]
            kw = query.get("keyword", [None])[0]
            patients = get_patients_for_report_making(
                db_path, date_raw=d_raw, keyword=kw, report_date=rep_date, submission_date=sub_date
            )
            self.send_json({"success": True, "patients": patients, "count": len(patients)})
            return

        # Patients list for Review workspace
        if path == "/api/reviews/list":
            if not cfg.get("initialized"):
                self.send_json({"success": True, "patients": []})
                return
            inc = query.get("include_approved", ["0"])[0]
            include_approved = inc in ["1", "true", "True"]
            rep_date = query.get("report_date", [None])[0]
            sub_date = query.get("submission_date", [None])[0]
            kw = query.get("keyword", [None])[0]
            patients = get_patients_for_review(
                db_path, include_approved=include_approved,
                report_date=rep_date, submission_date=sub_date, keyword=kw
            )
            enriched = [_attach_review_photo_dir(work_dir, p) for p in patients]
            self.send_json({"success": True, "patients": enriched, "count": len(enriched)})
            return

        # Copied pathology photos for Review workspace (/api/reviews/<pathology_no>/photos)
        if path.startswith("/api/reviews/") and path.endswith("/photos"):
            if not cfg.get("initialized"):
                self.send_error_json("系统尚未初始化工作目录")
                return
            parts = path.strip("/").split("/")
            if len(parts) == 4 and parts[3] == "photos":
                pathology_no = parts[2].strip().upper()
                patient = get_patient(db_path, pathology_no)
                if not patient:
                    self.send_error_json(f"未找到病理号为 [{pathology_no}] 的患者", status=404)
                    return
                rev_info = find_patient_review_folder(
                    work_dir,
                    pathology_no,
                    patient["name"],
                    review_status=patient.get("review_status", ""),
                    is_pending=patient.get("is_pending_review"),
                    is_recut=patient.get("is_needs_recut")
                )
                is_copied = bool(rev_info and rev_info.get("photo_dir"))
                photo_dir = rev_info["photo_dir"] if is_copied else get_patient_photo_dir(work_dir, pathology_no, patient["name"])
                img_cnt_param = query.get("image_count", [None])[0]
                eff_img_cnt = img_cnt_param if (img_cnt_param is not None and str(img_cnt_param).strip() != "") else patient.get("image_count")
                photos, auto_diag, recs = _scan_photos_dir(photo_dir, image_count=eff_img_cnt)
                self.send_json({
                    "success": True,
                    "is_copied_dir": is_copied,
                    "category": rev_info.get("category", "") if rev_info else "",
                    "folder_dir": rev_info.get("folder_dir", "") if rev_info else "",
                    "photo_dir": photo_dir,
                    "photos": photos,
                    "count": len(photos),
                    "auto_diagnosis": auto_diag,
                    "recommended_templates": recs
                })
                return

        # Patients list for IHC workspace
        if path == "/api/ihc/list":
            if not cfg.get("initialized"):
                self.send_json({"success": True, "patients": []})
                return
            rep_date = query.get("report_date", [None])[0]
            sub_date = query.get("submission_date", [None])[0]
            kw = query.get("keyword", [None])[0]
            patients = get_patients_for_ihc(
                db_path, report_date=rep_date, submission_date=sub_date, keyword=kw
            )
            self.send_json({"success": True, "patients": patients, "count": len(patients)})
            return

        # IHC patient sub-resources (/api/ihc/<pathology_no>/photos | /preview)
        if path.startswith("/api/ihc/"):
            parts = path.strip("/").split("/")
            if len(parts) == 4:
                pathology_no = parts[2].strip().upper()
                patient = get_patient(db_path, pathology_no)
                if not patient:
                    self.send_error_json(f"未找到病理号为 [{pathology_no}] 的患者", status=404)
                    return
                ihc = get_ihc_report(db_path, pathology_no)
                if parts[3] == "photos":
                    if not ihc or not ihc.get("ihc_created") or not ihc.get("ihc_photo_dir"):
                        self.send_json({
                            "success": True,
                            "ihc_created": False,
                            "photo_dir": "",
                            "photos": [],
                            "count": 0,
                            "auto_diagnosis": "",
                            "recommended_templates": []
                        })
                        return
                    img_cnt_param = query.get("image_count", [None])[0]
                    eff_img_cnt = img_cnt_param if (img_cnt_param is not None and str(img_cnt_param).strip() != "") else ihc.get("image_count", patient.get("image_count"))
                    photos, auto_diag, recs = _scan_photos_dir(ihc["ihc_photo_dir"], image_count=eff_img_cnt)
                    self.send_json({
                        "success": True,
                        "ihc_created": True,
                        "photo_dir": ihc["ihc_photo_dir"],
                        "photos": photos,
                        "count": len(photos),
                        "auto_diagnosis": auto_diag,
                        "recommended_templates": recs
                    })
                    return

                if parts[3] == "preview":
                    if not ihc or not ihc.get("ihc_created"):
                        self.send_error_json("该患者尚未创建免疫组化报告", status=400)
                        return
                    prev_data, doc_path = _build_ihc_preview_data(patient, ihc)
                    preview_html = render_preview_html(prev_data, doc_path=doc_path)
                    self.send_json({"success": True, "html": preview_html, "patient": prev_data, "doc_path": doc_path})
                    return

        # Search patients
        if path == "/api/patients/search":
            if not cfg.get("initialized"):
                self.send_json({"success": True, "patients": []})
                return
            query_params = {
                "start_date": query.get("start_date", [""])[0],
                "end_date": query.get("end_date", [""])[0],
                "report_start_date": query.get("report_start_date", [""])[0],
                "report_end_date": query.get("report_end_date", [""])[0],
                "prefix": query.get("prefix", ["ALL"])[0],
                "archived": query.get("archived", ["ALL"])[0],
                "review_status": query.get("review_status", ["ALL"])[0],
                "keyword": query.get("keyword", [""])[0],
                "name": query.get("name", [""])[0],
                "pathology_no": query.get("pathology_no", [""])[0],
                "inpatient_no": query.get("inpatient_no", [""])[0],
                "outpatient_no": query.get("outpatient_no", [""])[0]
            }
            results = search_patients(db_path, query_params)
            self.send_json({"success": True, "patients": results, "count": len(results)})
            return

        def _parse_prefixes_param(pfx_str):
            if pfx_str is None:
                return None
            if pfx_str == "__EMPTY__":
                return []
            if pfx_str in ("ALL", "所有", ""):
                return None
            return [p.strip() for p in pfx_str.split(",") if p.strip()]

        # Report Output Workspace Patients (3.1)
        if path == "/api/output/patients":
            if not cfg.get("initialized"):
                self.send_json({"success": True, "patients": [], "count": 0})
                return
            out_type = query.get("type", ["regular"])[0]
            prefix_val = query.get("prefix", [None])[0]
            prefixes_str = query.get("prefixes", [None])[0]
            prefixes_list = _parse_prefixes_param(prefixes_str)
            today_only_val = query.get("today_only", ["0"])[0] in ["1", "true", "True"]
            rep_start = query.get("report_start_date", [None])[0]
            rep_end = query.get("report_end_date", [None])[0]
            sub_start = query.get("sub_start_date", [None])[0]
            sub_end = query.get("sub_end_date", [None])[0]

            if out_type == "special":
                fn = get_output_special_approved_patients
            elif out_type == "ihc":
                fn = get_output_ihc_patients
            else:
                fn = get_output_regular_patients
            patients = fn(
                db_path,
                prefix=prefix_val,
                today_only=today_only_val,
                prefixes=prefixes_list,
                report_start_date=rep_start,
                report_end_date=rep_end,
                sub_start_date=sub_start,
                sub_end_date=sub_end
            )
            self.send_json({"success": True, "patients": patients, "count": len(patients)})
            return

        # Signing Sheet Workspace Candidates (3.2)
        if path == "/api/signing/patients":
            if not cfg.get("initialized"):
                self.send_json({"success": True, "patients": [], "count": 0})
                return
            s_date = query.get("start_date", [None])[0] or query.get("sub_start_date", [None])[0]
            e_date = query.get("end_date", [None])[0] or query.get("sub_end_date", [None])[0]
            rep_s = query.get("report_start_date", [None])[0]
            rep_e = query.get("report_end_date", [None])[0]
            pfx_str = query.get("prefixes", ["ALL"])[0]
            prefixes_list = _parse_prefixes_param(pfx_str)

            patients = get_patients_for_signing_sheet(
                db_path, start_date=s_date, end_date=e_date, prefixes=prefixes_list,
                report_start_date=rep_s, report_end_date=rep_e
            )
            self.send_json({"success": True, "patients": patients, "count": len(patients)})
            return

        # Patient detail or sub-resource
        if path.startswith("/api/patients/"):
            parts = path.strip("/").split("/")
            if len(parts) >= 3:
                pathology_no = parts[2]
                patient = get_patient(db_path, pathology_no)
                if not patient:
                    self.send_error_json(f"未找到病理号为 [{pathology_no}] 的患者", status=404)
                    return

                # Photos of patient in 病理照片/<病理号>_<姓名>
                if len(parts) == 4 and parts[3] == "photos":
                    photo_dir = get_patient_photo_dir(work_dir, pathology_no, patient["name"])
                    img_cnt_param = query.get("image_count", [None])[0]
                    eff_img_cnt = img_cnt_param if (img_cnt_param is not None and str(img_cnt_param).strip() != "") else patient.get("image_count")
                    photos, auto_diag, recs = _scan_photos_dir(photo_dir, image_count=eff_img_cnt)

                    self.send_json({
                        "success": True,
                        "photo_dir": photo_dir,
                        "photos": photos,
                        "count": len(photos),
                        "auto_diagnosis": auto_diag,
                        "recommended_templates": recs
                    })
                    return

                # Preview
                if len(parts) == 4 and parts[3] == "preview":
                    kind = query.get("kind", ["normal"])[0]
                    if kind == "ihc":
                        ihc = get_ihc_report(db_path, pathology_no)
                        if not ihc or not ihc.get("ihc_created"):
                            self.send_error_json("该患者尚未创建免疫组化报告", status=400)
                            return
                        prev_data, doc_path = _build_ihc_preview_data(patient, ihc)
                        preview_html = render_preview_html(prev_data, doc_path=doc_path)
                        self.send_json({"success": True, "html": preview_html, "patient": prev_data, "doc_path": doc_path})
                        return
                    word_path = patient.get("final_report_path")
                    if not word_path or not os.path.exists(word_path):
                        word_path = patient.get("word_path") or get_patient_word_path(work_dir, pathology_no, patient["name"])
                    preview_html = render_preview_html(patient, doc_path=word_path)
                    self.send_json({"success": True, "html": preview_html, "patient": patient})
                    return

                # All related paths for patient (Requirement 3)
                if len(parts) == 4 and parts[3] == "paths":
                    ihc = get_ihc_report(db_path, pathology_no)
                    paths_list = locate_patient_all_paths(work_dir, patient, ihc_data=ihc)
                    self.send_json({"success": True, "paths": paths_list, "patient": patient})
                    return

                # Details
                self.send_json({"success": True, "patient": patient})
                return

        self.send_error_json("未知的GET API接口", status=404)

    def handle_api_post(self, path):
        cfg = load_config()
        work_dir = cfg.get("work_dir")
        db_path = get_database_path(work_dir) if work_dir else None
        body = self.read_json_body()

        # 1. System Initialization
        if path == "/api/init":
            chosen_work_dir = body.get("work_dir", "").strip() or cfg.get("work_dir")
            report_doc = body.get("report_doctor", "张医生").strip()
            review_doc = body.get("review_doctor", "李主任").strip()

            # Ensure workspace folders
            struct_result = ensure_workspace_structure(chosen_work_dir)
            target_db_path = get_database_path(chosen_work_dir)
            init_db(target_db_path)

            # Update configuration
            new_cfg = update_config({
                "initialized": True,
                "work_dir": chosen_work_dir,
                "report_doctor": report_doc,
                "review_doctor": review_doc
            })

            self.send_json({
                "success": True,
                "message": "系统初始化完成！工作目录与数据库已就绪。",
                "config": new_cfg,
                "structure": struct_result
            })
            return

        # 2. Update Configuration
        if path == "/api/config":
            updates = {}
            if "report_doctor" in body:
                updates["report_doctor"] = body["report_doctor"].strip()
            if "review_doctor" in body:
                updates["review_doctor"] = body["review_doctor"].strip()
            if "work_dir" in body and body["work_dir"].strip():
                new_work_dir = body["work_dir"].strip()
                ensure_workspace_structure(new_work_dir)
                init_db(get_database_path(new_work_dir))
                updates["work_dir"] = new_work_dir

            saved = update_config(updates)
            self.send_json({"success": True, "config": saved})
            return

        # 3. Create Patient
        if path == "/api/patients":
            if not cfg.get("initialized"):
                self.send_error_json("系统尚未初始化工作目录")
                return

            pathology_no = body.get("pathology_no", "").strip().upper()
            name = body.get("name", "").strip()
            if not pathology_no:
                self.send_error_json("病理号不能为空")
                return
            if not name:
                self.send_error_json("患者姓名不能为空")
                return

            # Check duplicate primary key
            existing = check_duplicate(db_path, pathology_no)
            if existing:
                self.send_error_json(f"病理号 [{pathology_no}] 已存在！对应患者：{existing['name']}", status=409)
                return

            # Default doctors from settings if not explicitly passed
            if not body.get("report_doctor"):
                body["report_doctor"] = cfg.get("report_doctor", "张医生")
            if not body.get("review_doctor"):
                body["review_doctor"] = cfg.get("review_doctor", "李主任")

            # 1.2.2 Create Photo folder
            photo_dir = ensure_patient_photo_dir(work_dir, pathology_no, name)
            body["photo_dir"] = photo_dir

            # 1.2.1 Generate Word Report file
            word_path = get_patient_word_path(work_dir, pathology_no, name)
            body["word_path"] = word_path
            templates_dir = cfg.get("templates_dir")

            generate_report_file(templates_dir, word_path, body)

            # 1.3 Insert to database
            created_patient = insert_patient(db_path, body)

            self.send_json({
                "success": True,
                "message": f"患者 [{name}] (病理号: {pathology_no}) 录入成功！",
                "patient": created_patient
            })
            return

        # 4. Open Word file in desktop application
        if path == "/api/open_file":
            file_path = body.get("path", "").strip()
            if not file_path or not os.path.exists(file_path):
                self.send_error_json("文件不存在或路径无效")
                return
            try:
                os.startfile(file_path)
                self.send_json({"success": True, "message": "已在本地程序中打开文件"})
            except Exception as e:
                self.send_error_json(f"打开文件失败: {e}")
            return

        # 5. Open directory in Windows Explorer
        if path == "/api/open_dir":
            dir_path = body.get("path", "").strip()
            if not dir_path or not os.path.exists(dir_path):
                self.send_error_json("文件夹不存在或路径无效")
                return
            try:
                subprocess.Popen(f'explorer "{os.path.normpath(dir_path)}"')
                self.send_json({"success": True, "message": "已在资源管理器中打开文件夹"})
            except Exception as e:
                self.send_error_json(f"打开文件夹失败: {e}")
            return

        # 6. Find clinical photos folder
        if path == "/api/clinical_photos/find":
            p_name = body.get("name", "").strip()
            p_no = body.get("pathology_no", "").strip()
            clin_res = find_clinical_photo_folder(work_dir, p_name, p_no)
            target_path = clin_res["path"]
            try:
                subprocess.Popen(f'explorer "{os.path.normpath(target_path)}"')
                if clin_res.get("multiple"):
                    matches_str = "、".join(clin_res.get("matches", []))
                    self.send_json({
                        "success": True,
                        "found": False,
                        "multiple": True,
                        "matches": clin_res.get("matches", []),
                        "path": target_path,
                        "message": f"提示：检测到存在多个匹配的临床照片文件夹 ({matches_str})，已为您在资源管理器中打开临床照片母目录，请手动确认！"
                    })
                elif clin_res.get("found"):
                    self.send_json({
                        "success": True,
                        "found": True,
                        "multiple": False,
                        "path": target_path,
                        "message": f"已在资源管理器中打开患者 [{p_name}] 临床照片文件夹"
                    })
                else:
                    self.send_json({
                        "success": False,
                        "found": False,
                        "multiple": False,
                        "path": target_path,
                        "message": f"未找到患者 [{p_name}] 专属临床照片文件夹，已直接打开临床照片根目录"
                    })
            except Exception as e:
                self.send_error_json(f"打开资源管理器失败: {e}")
            return

        # 7. Save Final Pathology Report (2.7)
        if path == "/api/reports/save":
            if not cfg.get("initialized"):
                self.send_error_json("系统尚未初始化工作目录")
                return

            p_no = body.get("pathology_no", "").strip().upper()
            patient = get_patient(db_path, p_no)
            if not patient:
                self.send_error_json(f"未找到病理号为 [{p_no}] 的患者", status=404)
                return

            path_diag = body.get("pathology_diag", "").strip()
            micro_desc = body.get("microscopic_desc", "").strip()
            selected_images = body.get("selected_images", [])
            image_count = int(body.get("image_count") or patient.get("image_count", 1))
            biopsy_sites = body.get("biopsy_sites") or patient.get("biopsy_sites", [])

            # Format final output report path: 病理号_姓名_病理诊断.docx
            final_report_path = get_final_report_path(work_dir, p_no, patient["name"], path_diag)

            data_for_doc = dict(patient)
            data_for_doc.update({
                "microscopic_desc": micro_desc,
                "pathology_diag": path_diag,
                "selected_images": selected_images,
                "image_count": image_count,
                "biopsy_sites": biopsy_sites,
                "report_doctor": body.get("report_doctor") or patient.get("report_doctor") or cfg.get("report_doctor", "张医生"),
                "review_doctor": body.get("review_doctor") or patient.get("review_doctor") or cfg.get("review_doctor", "李主任"),
                "report_date": body.get("report_date") or patient.get("report_date")
            })

            templates_dir = cfg.get("templates_dir")
            try:
                generate_report_file(templates_dir, final_report_path, data_for_doc)
                cleanup_stale_output_reports(work_dir, p_no, patient["name"], keep_path=final_report_path)
            except PermissionError as pe:
                self.send_error_json(str(pe), status=409)
                return

            # Update database
            update_patient_report_making(db_path, p_no, {
                "pathology_diag": path_diag,
                "microscopic_desc": micro_desc,
                "selected_images": selected_images,
                "final_report_path": final_report_path,
                "is_pending_review": 1 if body.get("is_pending_review") else 0,
                "is_needs_ihc": 1 if body.get("is_needs_ihc") else 0,
                "is_needs_recut": 1 if body.get("is_needs_recut") else 0,
                "report_doctor": data_for_doc["report_doctor"],
                "review_doctor": data_for_doc["review_doctor"]
            })

            # Sync image_count and biopsy_sites if modified
            if image_count != patient["image_count"] or biopsy_sites != patient["biopsy_sites"]:
                update_patient(db_path, p_no, {
                    **patient,
                    "image_count": image_count,
                    "biopsy_sites": biopsy_sites
                })

            updated_p = get_patient(db_path, p_no)
            # Synchronize special patient draft queue (待审核/重切) or 完成审核 draft, NO clinical photos copied
            try:
                if updated_p.get("review_status") == "完成审核" and not updated_p.get("is_pending_review") and not updated_p.get("is_needs_recut"):
                    sync_updated_report_to_special_archives(work_dir, updated_p, final_report_path)
                    save_res = {"success": True, "duplicate_clinical_photos": False}
                else:
                    save_res = save_special_draft_queue(work_dir, updated_p, final_report_path)
            except PermissionError as pe:
                self.send_error_json(str(pe), status=409)
                return

            updated_p = _attach_review_photo_dir(work_dir, updated_p)
            self.send_json({
                "success": True,
                "message": f"报告制作完成并已输出！文件: {os.path.basename(final_report_path)}",
                "final_report_path": final_report_path,
                "patient": updated_p,
                "duplicate_clinical_photos": False
            })
            return

        def _archive_single_patient(patient, source="maker", allow_word_fallback=True):
            p_no = patient["pathology_no"]
            rev_st = str(patient.get("review_status", ""))
            is_approved = (rev_st == "完成审核")
            is_pending = int(patient.get("is_pending_review", 0)) == 1 or "待审核" in rev_st
            is_recut = int(patient.get("is_needs_recut", 0)) == 1 or "重切" in rev_st
            if source == "maker" and not is_approved and (is_pending or is_recut):
                return (
                    400,
                    "标记为“待审核”或“重切”的报告禁止在报告制作工作区直接归档！已在保存时自动放入特殊患者待审队列，请移步【报告审核】工作区完成专家审核后再进行归档。",
                    None,
                )

            report_file = patient.get("final_report_path")
            if allow_word_fallback and (not report_file or not os.path.exists(report_file)):
                report_file = patient.get("word_path")

            if not report_file or not os.path.exists(report_file):
                return 400, "未找到该患者生成的Word报告，请先制作并保存报告", None

            try:
                arch_res = archive_patient_report(work_dir, patient, report_file, is_from_review=(source == "review" or is_approved))
            except PermissionError as pe:
                return 409, str(pe), None
            if not arch_res.get("success"):
                return 400, arch_res.get("error", "归档失败"), None

            mark_patient_archived(db_path, p_no, arch_res["archive_dir"])
            updated_p = _attach_review_photo_dir(work_dir, get_patient(db_path, p_no))
            return 200, "", {"arch_res": arch_res, "patient": updated_p}

        # 8. Archive Patient Report (2.8 & 3)
        if path == "/api/reports/archive":
            if not cfg.get("initialized"):
                self.send_error_json("系统尚未初始化工作目录")
                return

            p_no = body.get("pathology_no", "").strip().upper()
            source = body.get("source", "maker")
            patient = get_patient(db_path, p_no)
            if not patient:
                self.send_error_json(f"未找到病理号为 [{p_no}] 的患者", status=404)
                return

            st_code, err_msg, res_data = _archive_single_patient(patient, source=source, allow_word_fallback=True)
            if st_code != 200:
                self.send_error_json(err_msg, status=st_code)
                return

            arch_res = res_data["arch_res"]
            updated_p = res_data["patient"]
            self.send_json({
                "success": True,
                "message": f"患者 [{patient['name']}] 报告已成功归档入【已归档报告】！",
                "archive_dir": arch_res["archive_dir"],
                "special_paths": arch_res.get("special_paths", []),
                "patient": updated_p,
                "duplicate_clinical_photos": False
            })
            return

        # 8.1 Batch Archive Today's Saved-Unarchived Reports in Report Making
        if path == "/api/reports/batch_archive_today":
            if not cfg.get("initialized"):
                self.send_error_json("系统尚未初始化工作目录")
                return

            today_str = (body.get("date") or "").strip() or datetime.now().strftime("%Y-%m-%d")
            dry_run = bool(body.get("dry_run", False))
            candidates = [
                p for p in get_today_batch_archivable_patients(db_path, today_str=today_str)
                if p.get("final_report_path") and os.path.exists(p.get("final_report_path"))
            ]

            if dry_run:
                self.send_json({
                    "success": True,
                    "today": today_str,
                    "count": len(candidates),
                    "candidates": [
                        {
                            "pathology_no": p["pathology_no"],
                            "name": p["name"],
                            "review_status": p.get("review_status", ""),
                            "is_needs_ihc": p.get("is_needs_ihc", 0),
                            "archived": p.get("archived", 0)
                        }
                        for p in candidates
                    ]
                })
                return

            if not candidates:
                self.send_json({
                    "success": True,
                    "today": today_str,
                    "archived_count": 0,
                    "failed_count": 0,
                    "archived_nos": [],
                    "failed": [],
                    "message": "今日暂无可归档的已保存未归档报告（普通及仅免疫组化）"
                })
                return

            archived_nos = []
            failed = []
            for p in candidates:
                st_code, err_msg, _ = _archive_single_patient(p, source="maker", allow_word_fallback=False)
                if st_code == 200:
                    archived_nos.append(p["pathology_no"])
                else:
                    failed.append({
                        "pathology_no": p["pathology_no"],
                        "name": p.get("name", ""),
                        "error": err_msg
                    })

            if archived_nos and not failed:
                msg = f"今日已成功批量归档 {len(archived_nos)} 份已保存报告！"
            elif archived_nos and failed:
                msg = f"批量归档部分完成：成功 {len(archived_nos)} 份，失败 {len(failed)} 份"
            else:
                msg = f"批量归档失败（{len(failed)} 份未能归档）：{failed[0]['error']}"

            self.send_json({
                "success": len(archived_nos) > 0,
                "today": today_str,
                "archived_count": len(archived_nos),
                "failed_count": len(failed),
                "archived_nos": archived_nos,
                "failed": failed,
                "message": msg
            }, status=200 if archived_nos else 409)
            return

        def _save_review_report_internal(req_body):
            p_no = req_body.get("pathology_no", "").strip().upper()
            patient = get_patient(db_path, p_no)
            if not patient:
                return 404, f"未找到病理号为 [{p_no}] 的患者", None

            path_diag = req_body.get("pathology_diag", patient.get("pathology_diag", "")).strip()
            micro_desc = req_body.get("microscopic_desc", patient.get("microscopic_desc", "")).strip()
            selected_images = req_body.get("selected_images")
            if selected_images is None:
                selected_images = patient.get("selected_images", [])
            image_count = int(req_body.get("image_count") or patient.get("image_count", 1))
            biopsy_sites = req_body.get("biopsy_sites")
            if biopsy_sites is None:
                biopsy_sites = patient.get("biopsy_sites", [])

            rep_doc = req_body.get("report_doctor") or patient.get("report_doctor") or "张医生"
            rev_doc = req_body.get("review_doctor") or patient.get("review_doctor") or "李主任"
            rep_date_cn, rep_date_raw = normalize_report_date(req_body.get("report_date"))
            if not rep_date_cn:
                rep_date_cn = patient.get("report_date", "")
            if not rep_date_raw:
                rep_date_raw = patient.get("report_date_raw", "")

            # Sync any new pathology photos from copied review directory -> 未归档报告/病理照片 & siblings
            rev_info = find_patient_review_folder(
                work_dir, p_no, patient["name"],
                review_status=patient.get("review_status", ""),
                is_pending=patient.get("is_pending_review"),
                is_recut=patient.get("is_needs_recut")
            )
            making_photo_dir = get_patient_photo_dir(work_dir, p_no, patient["name"])
            sync_review_photos_to_making_and_siblings(work_dir, patient, rev_info.get("photo_dir", ""))
            selected_images = _canonicalize_review_selected_images(selected_images, making_photo_dir)

            final_report_path = get_final_report_path(work_dir, p_no, patient["name"], path_diag)

            data_for_doc = dict(patient)
            data_for_doc.update({
                "microscopic_desc": micro_desc,
                "pathology_diag": path_diag,
                "selected_images": selected_images,
                "image_count": image_count,
                "biopsy_sites": biopsy_sites,
                "report_doctor": rep_doc,
                "review_doctor": rev_doc,
                "report_date": rep_date_cn
            })

            templates_dir = cfg.get("templates_dir")
            try:
                generate_report_file(templates_dir, final_report_path, data_for_doc)
                cleanup_stale_output_reports(work_dir, p_no, patient["name"], keep_path=final_report_path)
            except PermissionError as pe:
                return 409, str(pe), None

            # Update database (if already archived, mark archived=2 so user re-archives in Report Making)
            is_pending_val = req_body.get("is_pending_review", patient.get("is_pending_review", 0))
            is_ihc_val = req_body.get("is_needs_ihc", patient.get("is_needs_ihc", 0))
            is_recut_val = req_body.get("is_needs_recut", patient.get("is_needs_recut", 0))
            update_patient_report_making(db_path, p_no, {
                "pathology_diag": path_diag,
                "microscopic_desc": micro_desc,
                "selected_images": selected_images,
                "final_report_path": final_report_path,
                "is_pending_review": 1 if is_pending_val else 0,
                "is_needs_ihc": 1 if is_ihc_val else 0,
                "is_needs_recut": 1 if is_recut_val else 0,
                "report_doctor": rep_doc,
                "review_doctor": rev_doc,
                "report_date": rep_date_cn,
                "report_date_raw": rep_date_raw
            })
            if image_count != patient["image_count"] or biopsy_sites != patient["biopsy_sites"]:
                cur_p = get_patient(db_path, p_no)
                update_patient(db_path, p_no, {
                    **cur_p,
                    "image_count": image_count,
                    "biopsy_sites": biopsy_sites
                })

            updated_p = get_patient(db_path, p_no)
            overwritten_paths = []
            try:
                if updated_p.get("review_status") == "完成审核" and not updated_p.get("is_pending_review") and not updated_p.get("is_needs_recut"):
                    sync_res = sync_review_overwrites(work_dir, updated_p, final_report_path)
                    overwritten_paths = sync_res.get("overwritten", [])
                else:
                    save_res = save_special_draft_queue(work_dir, updated_p, final_report_path)
                    for d in save_res.get("saved_dirs", []):
                        overwritten_paths.append(os.path.join(d, f"[病理报告]{os.path.basename(d)}.docx"))
            except PermissionError as pe:
                return 409, str(pe), None

            updated_p = _attach_review_photo_dir(work_dir, updated_p)
            return 200, "", {
                "overwritten": overwritten_paths,
                "final_report_path": final_report_path,
                "patient": updated_p
            }

        # 9. Save & Sync Review Overwrite (2.9 & 3)
        if path == "/api/reviews/save":
            if not cfg.get("initialized"):
                self.send_error_json("系统尚未初始化工作目录")
                return

            st_code, err_msg, res_data = _save_review_report_internal(body)
            if st_code != 200:
                self.send_error_json(err_msg, status=st_code)
                return

            updated_p = res_data["patient"]
            rearchive_hint = "（该患者此前已归档，请前往【报告制作】模块点击“一键归档”以更新正式归档报告）" if updated_p.get("archived") == 2 else ""
            self.send_json({
                "success": True,
                "message": f"审核内容已保存并同步草稿目录！{rearchive_hint}",
                "overwritten": res_data["overwritten"],
                "final_report_path": res_data["final_report_path"],
                "patient": updated_p,
                "duplicate_clinical_photos": False
            })
            return

        # 9.1 Save Review Edits and Print Page 1 Only
        if path == "/api/reviews/print":
            if not cfg.get("initialized"):
                self.send_error_json("系统尚未初始化工作目录")
                return

            st_code, err_msg, res_data = _save_review_report_internal(body)
            if st_code != 200:
                self.send_error_json(err_msg, status=st_code)
                return

            final_report_path = res_data["final_report_path"]
            updated_p = res_data["patient"]
            printed, failed, print_err = print_word_documents([final_report_path], page1_only=True)
            print_ok = (len(printed) == 1 and not failed)

            self.send_json({
                "success": True,
                "message": "审核报告已保存，已调用Word发送第1页至默认打印机！" if print_ok else f"审核报告已保存，但调用Word打印失败: {print_err or '请检查Word及默认打印机'}",
                "direct_printed": print_ok,
                "print_error": print_err,
                "overwritten": res_data["overwritten"],
                "final_report_path": final_report_path,
                "patient": updated_p
            })
            return

        # 10. Mark Review Complete (2.9)
        if path == "/api/reviews/approve":
            if not cfg.get("initialized"):
                self.send_error_json("系统尚未初始化工作目录")
                return

            p_no = body.get("pathology_no", "").strip().upper()
            patient = get_patient(db_path, p_no)
            if not patient:
                self.send_error_json(f"未找到病理号为 [{p_no}] 的患者", status=404)
                return

            approved_p = _attach_review_photo_dir(work_dir, mark_patient_approved(db_path, p_no))
            self.send_json({
                "success": True,
                "message": f"患者 [{patient['name']}] (病理号: {p_no}) 已标注为【完成审核】，请至【报告制作】模块完成归档！",
                "patient": approved_p
            })
            return

        # 10.1 Create IHC Report Package (copy archived report & pathology photos into 特殊患者归档/免疫组化, no clinical photos)
        if path == "/api/ihc/create":
            if not cfg.get("initialized"):
                self.send_error_json("系统尚未初始化工作目录")
                return

            p_no = body.get("pathology_no", "").strip().upper()
            patient = get_patient(db_path, p_no)
            if not patient:
                self.send_error_json(f"未找到病理号为 [{p_no}] 的患者", status=404)
                return

            # Locate original archived report in 已归档报告/ (strictly required)
            arch_file = find_patient_archive_path(work_dir, p_no, patient["name"])
            if (not arch_file or not os.path.exists(arch_file)) and patient.get("archive_dir"):
                cand = patient.get("archive_dir")
                if os.path.isfile(cand) and not cand.endswith("_免疫组化报告.docx"):
                    arch_file = cand

            if not arch_file or not os.path.exists(arch_file):
                self.send_error_json(
                    "该患者原报告尚未归档至【已归档报告】，请先完成原报告归档后再创建免疫组化报告！",
                    status=400
                )
                return

            try:
                pkg = create_ihc_report_package(work_dir, patient, arch_file)
            except PermissionError as pe:
                self.send_error_json(str(pe), status=409)
                return

            doc_details = extract_docx_details(arch_file) or {}
            ihc_saved = upsert_ihc_report(db_path, p_no, {
                "ihc_created": 1,
                "ihc_dir": pkg["ihc_dir"],
                "ihc_doc_path": pkg["ihc_doc_path"],
                "ihc_photo_dir": pkg["ihc_photo_dir"],
                "source_archive_path": arch_file,
                "source_archive_filename": pkg["source_archive_filename"],
                "report_doctor": patient.get("report_doctor") or cfg.get("report_doctor", "张医生"),
                "review_doctor": patient.get("review_doctor") or cfg.get("review_doctor", "李主任"),
                "report_date": patient.get("report_date", ""),
                "report_date_raw": patient.get("report_date_raw", ""),
                "microscopic_desc": patient.get("microscopic_desc") or doc_details.get("microscopic_desc", ""),
                "pathology_diag": patient.get("pathology_diag") or doc_details.get("pathology_diag", ""),
                "selected_images": pkg["selected_images"],
                "image_count": patient.get("image_count", 1),
                "biopsy_sites": patient.get("biopsy_sites", []),
                "archived": 0,
                "archive_file": ""
            }, mode="create")

            updated_p = dict(patient)
            updated_p["ihc_report"] = ihc_saved
            self.send_json({
                "success": True,
                "message": f"已为患者 [{patient['name']}] 创建独立的免疫组化报告草稿与照片目录！",
                "ihc_report": ihc_saved,
                "patient": updated_p
            })
            return

        # 10.2 Save Independent IHC Report (only updates 特殊患者归档/免疫组化 docx & ihc_reports)
        if path == "/api/ihc/save":
            if not cfg.get("initialized"):
                self.send_error_json("系统尚未初始化工作目录")
                return

            p_no = body.get("pathology_no", "").strip().upper()
            patient = get_patient(db_path, p_no)
            if not patient:
                self.send_error_json(f"未找到病理号为 [{p_no}] 的患者", status=404)
                return

            ihc = get_ihc_report(db_path, p_no)
            if not ihc or not ihc.get("ihc_created") or not ihc.get("ihc_doc_path"):
                self.send_error_json("请先点击【创建免疫组化报告】后再保存", status=400)
                return

            path_diag = body.get("pathology_diag", ihc.get("pathology_diag", "")).strip()
            micro_desc = body.get("microscopic_desc", ihc.get("microscopic_desc", "")).strip()
            selected_images = body.get("selected_images")
            if selected_images is None:
                selected_images = ihc.get("selected_images", [])
            image_count = int(body.get("image_count") or ihc.get("image_count") or patient.get("image_count", 1))
            biopsy_sites = body.get("biopsy_sites")
            if biopsy_sites is None:
                biopsy_sites = ihc.get("biopsy_sites") or patient.get("biopsy_sites", [])

            rep_doc = body.get("report_doctor") or ihc.get("report_doctor") or cfg.get("report_doctor", "张医生")
            rev_doc = body.get("review_doctor") or ihc.get("review_doctor") or cfg.get("review_doctor", "李主任")
            rep_date_cn, rep_date_raw = normalize_report_date(body.get("report_date"))
            if not rep_date_cn:
                rep_date_cn = ihc.get("report_date") or patient.get("report_date", "")
            if not rep_date_raw:
                rep_date_raw = ihc.get("report_date_raw") or patient.get("report_date_raw", "")

            data_for_doc = dict(patient)
            data_for_doc.update({
                "microscopic_desc": micro_desc,
                "pathology_diag": path_diag,
                "selected_images": selected_images,
                "image_count": image_count,
                "biopsy_sites": biopsy_sites,
                "report_doctor": rep_doc,
                "review_doctor": rev_doc,
                "report_date": rep_date_cn
            })

            templates_dir = cfg.get("templates_dir")
            try:
                generate_report_file(templates_dir, ihc["ihc_doc_path"], data_for_doc)
            except PermissionError as pe:
                self.send_error_json(str(pe), status=409)
                return

            ihc_saved = upsert_ihc_report(db_path, p_no, {
                "report_doctor": rep_doc,
                "review_doctor": rev_doc,
                "report_date": rep_date_cn,
                "report_date_raw": rep_date_raw,
                "microscopic_desc": micro_desc,
                "pathology_diag": path_diag,
                "selected_images": selected_images,
                "image_count": image_count,
                "biopsy_sites": biopsy_sites
            }, mode="save")

            updated_p = dict(patient)
            updated_p["ihc_report"] = ihc_saved
            self.send_json({
                "success": True,
                "message": f"免疫组化报告已独立保存（未改动原报告）！文件: {os.path.basename(ihc['ihc_doc_path'])}",
                "ihc_doc_path": ihc["ihc_doc_path"],
                "ihc_report": ihc_saved,
                "patient": updated_p
            })
            return

        # 10.3 Archive Independent IHC Report (document only -> 已归档报告/<原文件名>_免疫组化报告.docx, no move to 完成审核)
        if path == "/api/ihc/archive":
            if not cfg.get("initialized"):
                self.send_error_json("系统尚未初始化工作目录")
                return

            p_no = body.get("pathology_no", "").strip().upper()
            patient = get_patient(db_path, p_no)
            if not patient:
                self.send_error_json(f"未找到病理号为 [{p_no}] 的患者", status=404)
                return

            ihc = get_ihc_report(db_path, p_no)
            if not ihc or not ihc.get("ihc_created") or not ihc.get("ihc_doc_path") or not os.path.exists(ihc["ihc_doc_path"]):
                self.send_error_json("未找到可归档的免疫组化报告，请先创建并保存免疫组化报告", status=400)
                return

            try:
                arch_res = archive_ihc_report(work_dir, patient, ihc)
            except PermissionError as pe:
                self.send_error_json(str(pe), status=409)
                return

            if not arch_res.get("success"):
                self.send_error_json(arch_res.get("error", "免疫组化报告归档失败"), status=400)
                return

            ihc_archived = mark_ihc_archived(db_path, p_no, arch_res["archive_file"])
            updated_p = dict(patient)
            updated_p["ihc_report"] = ihc_archived
            self.send_json({
                "success": True,
                "message": f"免疫组化报告已归档至【已归档报告】：{arch_res['filename']}",
                "archive_file": arch_res["archive_file"],
                "filename": arch_res["filename"],
                "ihc_report": ihc_archived,
                "patient": updated_p
            })
            return

        # 10.4 Delete Independent IHC Report (Recycle Bin) and Revoke Patient IHC Status
        if path == "/api/ihc/delete":
            if not cfg.get("initialized"):
                self.send_error_json("系统尚未初始化工作目录")
                return

            p_no = body.get("pathology_no", "").strip().upper()
            patient = get_patient(db_path, p_no)
            if not patient:
                self.send_error_json(f"未找到病理号为 [{p_no}] 的患者", status=404)
                return

            ihc = get_ihc_report(db_path, p_no)
            try:
                del_res = delete_ihc_report_files(work_dir, patient, ihc_data=ihc)
            except PermissionError as pe:
                self.send_error_json(str(pe), status=409)
                return

            updated_p = delete_ihc_report_and_clear_status(db_path, p_no)
            self.send_json({
                "success": True,
                "message": f"已将患者 [{patient['name']}] 的免疫组化报告及目录移入回收站，并撤销其免疫组化状态！",
                "removed": del_res.get("removed", []),
                "patient": updated_p
            })
            return

        # 11. Open Explorer with item selected (Requirement 2 & 3)
        if path == "/api/open_explorer_select":
            target_path = body.get("path", "").strip()
            if not target_path or not os.path.exists(target_path):
                self.send_error_json("路径不存在或无效", status=404)
                return
            try:
                norm = os.path.normpath(target_path)
                if os.path.isfile(norm):
                    subprocess.Popen(f'explorer /select,"{norm}"')
                else:
                    subprocess.Popen(f'explorer "{norm}"')
                self.send_json({"success": True, "message": "已在资源管理器中定位打开！"})
            except Exception as e:
                self.send_error_json(f"打开资源管理器失败: {e}")
            return

        # 12. Save and Print Page 1 (Requirement 2)
        if path == "/api/reports/print":
            if not cfg.get("initialized"):
                self.send_error_json("系统尚未初始化工作目录")
                return

            p_no = body.get("pathology_no", "").strip().upper()
            patient = get_patient(db_path, p_no)
            if not patient:
                self.send_error_json(f"未找到病理号为 [{p_no}] 的患者", status=404)
                return

            path_diag = body.get("pathology_diag", "").strip()
            micro_desc = body.get("microscopic_desc", "").strip()
            selected_images = body.get("selected_images", [])
            image_count = int(body.get("image_count") or patient.get("image_count", 1))
            biopsy_sites = body.get("biopsy_sites") or patient.get("biopsy_sites", [])

            final_report_path = get_final_report_path(work_dir, p_no, patient["name"], path_diag)

            data_for_doc = dict(patient)
            data_for_doc.update({
                "microscopic_desc": micro_desc,
                "pathology_diag": path_diag,
                "selected_images": selected_images,
                "image_count": image_count,
                "biopsy_sites": biopsy_sites,
                "report_doctor": body.get("report_doctor") or patient.get("report_doctor") or cfg.get("report_doctor", "张医生"),
                "review_doctor": body.get("review_doctor") or patient.get("review_doctor") or cfg.get("review_doctor", "李主任"),
                "report_date": body.get("report_date") or patient.get("report_date")
            })

            templates_dir = cfg.get("templates_dir")
            try:
                generate_report_file(templates_dir, final_report_path, data_for_doc)
                cleanup_stale_output_reports(work_dir, p_no, patient["name"], keep_path=final_report_path)
            except PermissionError as pe:
                self.send_error_json(str(pe), status=409)
                return

            # Update DB (and mark archived=2 if already archived)
            update_patient_report_making(db_path, p_no, {
                "pathology_diag": path_diag,
                "microscopic_desc": micro_desc,
                "selected_images": selected_images,
                "final_report_path": final_report_path,
                "is_pending_review": 1 if body.get("is_pending_review") else 0,
                "is_needs_ihc": 1 if body.get("is_needs_ihc") else 0,
                "is_needs_recut": 1 if body.get("is_needs_recut") else 0,
                "report_doctor": data_for_doc["report_doctor"],
                "review_doctor": data_for_doc["review_doctor"]
            })

            updated_p = get_patient(db_path, p_no)
            try:
                if updated_p.get("review_status") == "完成审核" and not updated_p.get("is_pending_review") and not updated_p.get("is_needs_recut"):
                    sync_updated_report_to_special_archives(work_dir, updated_p, final_report_path)
                else:
                    save_special_draft_queue(work_dir, updated_p, final_report_path)
            except PermissionError as pe:
                self.send_error_json(str(pe), status=409)
                return

            # Silent Word printing of Page 1 only
            printed, failed, print_err = print_word_documents([final_report_path], page1_only=True)
            print_ok = (len(printed) == 1 and not failed)

            updated_p = get_patient(db_path, p_no)
            self.send_json({
                "success": True,
                "message": "报告已保存，已调用Word发送第1页至默认打印机！" if print_ok else f"报告已保存，但调用Word打印失败: {print_err or '请检查Word及默认打印机'}",
                "direct_printed": print_ok,
                "print_error": print_err,
                "final_report_path": final_report_path,
                "patient": updated_p
            })
            return

        # 13. Update Patient Status & Sync Files (Requirement 3)
        if path.startswith("/api/patients/") and path.endswith("/status"):
            parts = path.strip("/").split("/")
            pathology_no = parts[2].strip().upper()
            patient = get_patient(db_path, pathology_no)
            if not patient:
                self.send_error_json(f"未找到病理号为 [{pathology_no}] 的患者", status=404)
                return

            new_rev = body.get("review_status")
            new_arch = body.get("archived")

            # Sync file operations
            sync_status_file_operations(
                work_dir,
                patient,
                new_rev if new_rev is not None else patient.get("review_status", ""),
                new_arch if new_arch is not None else patient.get("archived", 0)
            )

            # Update database
            updated_p = update_patient_status(db_path, pathology_no, review_status=new_rev, archived=new_arch)
            self.send_json({
                "success": True,
                "message": f"患者 [{patient['name']}] 状态已成功更新并同步相关文件！",
                "patient": updated_p
            })
            return

        # 14. Revoke Archive (Requirement 3.1)
        if path == "/api/output/revoke_archive":
            p_no = body.get("pathology_no", "").strip().upper()
            report_type = body.get("report_type", "")
            patient = get_patient(db_path, p_no)
            if not patient:
                self.send_error_json(f"未找到病理号为 [{p_no}] 的患者", status=404)
                return

            if report_type == "ihc":
                ihc = get_ihc_report(db_path, p_no)
                revoke_ihc_archive_files(work_dir, patient, ihc)
                updated_ihc = revoke_ihc_archive_db(db_path, p_no)
                updated_p = dict(patient)
                updated_p["ihc_report"] = updated_ihc
                self.send_json({
                    "success": True,
                    "message": f"已成功撤销患者 [{patient['name']}] 的免疫组化报告归档，并从已归档目录移除了对应免疫组化报告！",
                    "patient": updated_p
                })
                return

            revoke_patient_archive_files(work_dir, patient)
            updated_p = revoke_patient_archive_db(db_path, p_no)
            self.send_json({
                "success": True,
                "message": f"已成功撤销患者 [{patient['name']}] 的归档状态，并从归档目录中移除了对应文件！",
                "patient": updated_p
            })
            return

        # 15. Revoke Approval (Requirement 3.1)
        if path == "/api/output/revoke_approval":
            p_no = body.get("pathology_no", "").strip().upper()
            patient = get_patient(db_path, p_no)
            if not patient:
                self.send_error_json(f"未找到病理号为 [{p_no}] 的患者", status=404)
                return

            updated_p = revoke_patient_approval_db(db_path, p_no)
            revoke_patient_approval_files(work_dir, patient, updated_p.get("review_status", ""))
            self.send_json({
                "success": True,
                "message": f"已成功撤销患者 [{patient['name']}] 的完成审核状态，患者已返回待审核列表！",
                "patient": updated_p
            })
            return

        # 16. Open Archive Dir (Requirement 3.1)
        if path == "/api/output/open_archive_dir":
            p_no = body.get("pathology_no", "").strip().upper()
            report_type = body.get("report_type", "")
            patient = get_patient(db_path, p_no)
            if not patient:
                self.send_error_json(f"未找到病理号为 [{p_no}] 的患者", status=404)
                return

            if report_type == "ihc":
                ihc = get_ihc_report(db_path, p_no)
                arch_path = (ihc or {}).get("archive_file", "")
                if not arch_path or not os.path.exists(arch_path):
                    arch_path = find_patient_ihc_archive_path(work_dir, p_no, patient["name"])
            else:
                arch_path = find_patient_archive_path(work_dir, p_no, patient["name"])
            if arch_path and os.path.exists(arch_path):
                norm = os.path.normpath(arch_path)
                if os.path.isfile(norm):
                    subprocess.Popen(f'explorer /select,"{norm}"')
                else:
                    subprocess.Popen(f'explorer "{norm}"')
                self.send_json({
                    "success": True,
                    "message": f"已在资源管理器中定位归档: {os.path.basename(norm)}",
                    "path": norm
                })
            else:
                fallback = get_archive_root_dir(work_dir)
                os.makedirs(fallback, exist_ok=True)
                subprocess.Popen(["explorer.exe", os.path.normpath(fallback)])
                self.send_json({
                    "success": True,
                    "message": "该患者归档文件尚未生成，已打开“已归档报告”根目录",
                    "path": fallback
                })
            return

        # 17. Batch Print Page 1 of Reports in Output List (Requirement 3.1 & Print Requirement 1)
        if path == "/api/output/batch_print":
            p_nos = body.get("pathology_nos", [])
            report_type = body.get("report_type", "")
            if not p_nos:
                self.send_error_json("未指定需要打印的报告病理号")
                return

            doc_map = []
            missing_nos = []
            for p_no in p_nos:
                patient = get_patient(db_path, p_no)
                if not patient:
                    missing_nos.append(p_no)
                    continue
                if report_type == "ihc":
                    ihc = get_ihc_report(db_path, p_no)
                    word_p = (ihc or {}).get("archive_file", "")
                    if not word_p or not os.path.exists(word_p):
                        word_p = find_patient_ihc_archive_path(work_dir, p_no, patient["name"])
                    if not word_p or not os.path.exists(word_p):
                        word_p = (ihc or {}).get("ihc_doc_path", "")
                else:
                    word_p = patient.get("final_report_path")
                    if not word_p or not os.path.exists(word_p):
                        diag = patient.get("pathology_diag") or ""
                        word_p = get_final_report_path(work_dir, p_no, patient["name"], diag)
                    if not word_p or not os.path.exists(word_p):
                        arch_p = patient.get("archive_dir")
                        if arch_p and os.path.isfile(arch_p) and arch_p.lower().endswith(".docx") and not arch_p.endswith("_免疫组化报告.docx"):
                            word_p = arch_p
                    if not word_p or not os.path.exists(word_p):
                        word_p = patient.get("word_path")

                if word_p and os.path.exists(word_p):
                    doc_map.append((p_no, os.path.normpath(word_p)))
                else:
                    missing_nos.append(p_no)

            printed_paths, failed_paths, print_err = print_word_documents([wp for _, wp in doc_map], page1_only=True)
            printed_set = {os.path.normcase(p) for p in printed_paths}
            printed_list = [p_no for p_no, wp in doc_map if os.path.normcase(wp) in printed_set]
            failed_list = [p_no for p_no, wp in doc_map if os.path.normcase(wp) not in printed_set] + missing_nos

            direct_ok = (len(printed_list) > 0 and len(failed_list) == 0)
            self.send_json({
                "success": True,
                "total": len(p_nos),
                "printed_count": len(printed_list),
                "failed_count": len(failed_list),
                "direct_printed": direct_ok,
                "print_error": print_err,
                "printed": printed_list,
                "failed": failed_list,
                "message": (
                    f"批量Word打印完成：已发送 {len(printed_list)} 份报告第1页至默认打印机！"
                    if direct_ok else
                    f"Word打印未完全成功（成功 {len(printed_list)} / 失败 {len(failed_list)}）: {print_err or '请检查Word及默认打印机'}"
                )
            })
            return

        # 18. Generate Signing Sheet (Requirement 3.2)
        if path == "/api/signing/generate":
            p_nos = body.get("pathology_nos", [])
            start_date = body.get("start_date")
            end_date = body.get("end_date")
            report_start_date = body.get("report_start_date")
            report_end_date = body.get("report_end_date")
            prefixes = body.get("prefixes")
            custom_filename = body.get("filename")

            if p_nos:
                patients = []
                for p_no in p_nos:
                    p = get_patient(db_path, p_no)
                    if p:
                        patients.append(p)
            else:
                patients = get_patients_for_signing_sheet(
                    db_path,
                    start_date=start_date,
                    end_date=end_date,
                    prefixes=prefixes,
                    report_start_date=report_start_date,
                    report_end_date=report_end_date
                )

            if not patients:
                self.send_error_json("所选筛选范围内无匹配的患者数据，无法生成签名单")
                return

            file_start = start_date or report_start_date
            file_end = end_date or report_end_date
            gen_res = generate_signing_sheet_excel(
                work_dir,
                patients,
                start_date=file_start,
                end_date=file_end,
                prefixes=prefixes,
                custom_filename=custom_filename
            )

            if start_date or end_date:
                date_range_str = f"{start_date} 至 {end_date}" if (start_date and end_date and start_date != end_date) else (start_date or end_date)
            else:
                sub_dates = sorted({
                    str(p.get("submission_date_raw") or "").strip()[:10]
                    for p in patients
                    if str(p.get("submission_date_raw") or "").strip()
                })
                if sub_dates:
                    date_range_str = sub_dates[0] if sub_dates[0] == sub_dates[-1] else f"{sub_dates[0]} 至 {sub_dates[-1]}"
                elif report_start_date or report_end_date:
                    date_range_str = f"{report_start_date} 至 {report_end_date}" if (report_start_date and report_end_date and report_start_date != report_end_date) else (report_start_date or report_end_date)
                else:
                    date_range_str = datetime.now().strftime("%Y-%m-%d")

            word_custom_name = os.path.splitext(gen_res["filename"])[0] + ".docx"
            word_res = generate_signing_sheet_word(
                work_dir,
                patients,
                start_date=start_date,
                end_date=end_date,
                prefixes=prefixes,
                custom_filename=word_custom_name,
                date_range_str=date_range_str
            )
            surgery_date_str = word_res.get("surgery_date") or compute_surgery_date_str(start_date=start_date, end_date=end_date, patients=patients)
            html_prev = render_signing_sheet_html(patients, date_range_str=date_range_str, surgery_date_str=surgery_date_str)

            self.send_json({
                "success": True,
                "message": f"签名单（Excel + Word）已生成，包含 {len(patients)} 位患者！",
                "file_path": gen_res["file_path"],
                "filename": gen_res["filename"],
                "word_path": word_res["file_path"],
                "word_filename": word_res["filename"],
                "surgery_date": surgery_date_str,
                "count": len(patients),
                "html": html_prev
            })
            return

        # 19. Print Signing Sheet (Requirement 5.2 & Print Requirement 1: print Word signing sheet)
        if path == "/api/signing/print":
            word_path = body.get("word_path", "").strip()
            f_path = word_path or body.get("file_path", "").strip()
            if f_path and f_path.lower().endswith(".xlsx"):
                candidate_docx = os.path.splitext(f_path)[0] + ".docx"
                if os.path.exists(candidate_docx):
                    f_path = candidate_docx

            printed, failed, print_err = print_word_documents([f_path], page1_only=False)
            direct_printed = (len(printed) == 1 and not failed)

            self.send_json({
                "success": True,
                "direct_printed": direct_printed,
                "print_error": print_err,
                "message": "Word签名单已发送至默认打印机！" if direct_printed else f"Word打印失败: {print_err or '请检查Word及默认打印机'}"
            })
            return

        # 20. Exit System
        if path == "/api/system/exit":
            self.send_json({"success": True, "message": "系统正在安全退出..."})
            def _delayed_exit():
                import time
                time.sleep(0.5)
                remove_runtime_info(expected_pid=os.getpid())
                os._exit(0)
            threading.Thread(target=_delayed_exit, daemon=True).start()
            return

        self.send_error_json("未知的POST API接口", status=404)

    def handle_api_delete(self, path):
        cfg = load_config()
        work_dir = cfg.get("work_dir")
        db_path = get_database_path(work_dir) if work_dir else None

        # Delete patient (Requirement 3)
        if path.startswith("/api/patients/"):
            pathology_no = path.strip("/").split("/")[2].strip().upper()
            patient = get_patient(db_path, pathology_no)
            if not patient:
                self.send_error_json(f"未找到病理号为 [{pathology_no}] 的患者", status=404)
                return

            # Delete all related files and folders
            del_res = delete_patient_all_files(work_dir, patient)

            # Delete from SQLite database
            delete_patient(db_path, pathology_no)

            self.send_json({
                "success": True,
                "message": f"患者 [{patient['name']}] (病理号: {pathology_no}) 的病历及全部相关文件已彻底删除！",
                "deleted_files": del_res.get("deleted_items", []),
                "errors": del_res.get("errors", [])
            })
            return

        self.send_error_json("未知的DELETE API接口", status=404)

    def handle_api_put(self, path):
        cfg = load_config()
        work_dir = cfg.get("work_dir")
        db_path = get_database_path(work_dir) if work_dir else None
        body = self.read_json_body()

        # Update Patient
        if path.startswith("/api/patients/"):
            old_no = path.strip("/").split("/")[2].strip().upper()
            existing = get_patient(db_path, old_no)
            if not existing:
                self.send_error_json(f"未找到病理号为 [{old_no}] 的患者", status=404)
                return

            new_no = body.get("pathology_no", old_no).strip().upper()
            new_name = body.get("name", existing["name"]).strip()
            if new_no != old_no and get_patient(db_path, new_no):
                conflict = get_patient(db_path, new_no)
                self.send_error_json(f"无法修改病理号：目标病理号 [{new_no}] 已被患者 [{conflict['name']}] 使用", status=409)
                return

            # Handle rename of files and folders (excluding 已归档报告/)
            rename_res = rename_patient_files_and_folders(
                work_dir,
                old_no,
                existing["name"],
                new_no,
                new_name
            )
            body["word_path"] = rename_res["word_path"]
            body["photo_dir"] = rename_res["photo_dir"]

            # Update selected_images paths if photo folder moved
            old_photo_dir = rename_res.get("old_photo_dir", "")
            new_photo_dir = rename_res.get("photo_dir", "")
            selected_images = _remap_selected_images_dir(existing.get("selected_images", []), old_photo_dir, new_photo_dir)
            body["selected_images"] = selected_images

            # Parse updated submission_date if pathology_no changed
            from app.database import parse_pathology_no
            parsed_no = parse_pathology_no(new_no)
            if parsed_no.get("valid"):
                body.setdefault("submission_date", parsed_no["submission_date"])
                body.setdefault("submission_date_raw", parsed_no["submission_date_raw"])

            # Synchronize changes into draft Word document (病理报告录入)
            templates_dir = cfg.get("templates_dir")
            draft_doc_data = dict(existing)
            draft_doc_data.update(body)
            update_report_file(templates_dir, rename_res["word_path"], draft_doc_data)

            # Check if patient already has a saved output report or archive
            was_archived = existing.get("archived") in (1, 2)
            renamed_final = rename_res.get("final_report_path", "")
            existing_final = existing.get("final_report_path", "")
            has_saved_report = bool(renamed_final or (existing_final and os.path.exists(existing_final)) or was_archived)

            target_final_path = ""
            if has_saved_report:
                if renamed_final:
                    target_final_path = renamed_final
                elif existing_final and os.path.exists(existing_final):
                    target_final_path = existing_final
                else:
                    target_final_path = get_final_report_path(
                        work_dir,
                        new_no,
                        new_name,
                        existing.get("pathology_diag") or existing.get("clinical_diagnosis") or ""
                    )

                final_doc_data = dict(existing)
                final_doc_data.update(body)
                final_doc_data["selected_images"] = selected_images
                generate_report_file(templates_dir, target_final_path, final_doc_data)
                sync_updated_report_to_special_archives(work_dir, final_doc_data, target_final_path)
                body["final_report_path"] = target_final_path
            else:
                sync_updated_report_to_special_archives(work_dir, draft_doc_data, rename_res["word_path"])

            # Determine whether to mark 需要重新归档 and prompt user to re-archive
            # ONLY prompt when the patient's report had already been archived (was_archived).
            rev_st = str(existing.get("review_status", ""))
            is_approved = (rev_st == "完成审核")
            is_still_pending = (not is_approved) and (
                any(k in rev_st for k in ("待审核", "重切"))
                or int(existing.get("is_pending_review", 0)) == 1
                or int(existing.get("is_needs_recut", 0)) == 1
            )
            ask_rearchive = bool(was_archived and not is_still_pending)
            archive_source = "review" if is_approved else "maker"

            if was_archived:
                body["archived"] = 2

            # Update database record
            updated = update_patient(db_path, old_no, body)

            # Sync independent IHC report paths/demographics if IHC report exists
            ihc_existing = get_ihc_report(db_path, new_no)
            if ihc_existing and ihc_existing.get("ihc_created"):
                renamed_ihc = rename_res.get("renamed_ihc") or {}
                new_ihc_dir = renamed_ihc.get("ihc_dir") or ihc_existing.get("ihc_dir", "")
                new_ihc_doc = renamed_ihc.get("ihc_doc_path") or ihc_existing.get("ihc_doc_path", "")
                new_ihc_photo = renamed_ihc.get("ihc_photo_dir") or ihc_existing.get("ihc_photo_dir", "")
                old_ihc_photo = ihc_existing.get("ihc_photo_dir", "")
                ihc_imgs = _remap_selected_images_dir(ihc_existing.get("selected_images", []), old_ihc_photo, new_ihc_photo)
                update_ihc_paths(
                    db_path, new_no, new_ihc_dir, new_ihc_doc, new_ihc_photo,
                    selected_images=ihc_imgs, mark_rearchive=(ihc_existing.get("archived") in (1, 2))
                )
                if new_ihc_doc and os.path.exists(new_ihc_doc):
                    ihc_doc_data = dict(updated)
                    for k in ("report_doctor", "review_doctor", "report_date", "report_date_raw",
                              "microscopic_desc", "pathology_diag", "image_count", "biopsy_sites"):
                        if ihc_existing.get(k) not in (None, ""):
                            ihc_doc_data[k] = ihc_existing[k]
                    ihc_doc_data["selected_images"] = ihc_imgs
                    generate_report_file(templates_dir, new_ihc_doc, ihc_doc_data)

            self.send_json({
                "success": True,
                "message": f"患者 [{new_name}] 信息及Word文件已成功更新！",
                "ask_rearchive": ask_rearchive,
                "archive_source": archive_source,
                "patient": updated
            })
            return

        self.send_error_json("未知的PUT API接口", status=404)

def run_server(port=8765):
    """Start HTTP server."""
    server_address = ("127.0.0.1", port)
    httpd = ThreadedHTTPServer(server_address, PathologyRequestHandler)
    print(f"[Server] 皮肤病理报告系统已在 http://127.0.0.1:{port} 启动")
    return httpd
