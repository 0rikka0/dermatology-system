# -*- coding: utf-8 -*-
"""
Dermatology Pathology Reporting System - Web Server & RESTful API
Uses standard library ThreadingHTTPServer.
"""
import os
import sys
import json
import urllib.parse
import mimetypes
import subprocess
from datetime import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler
from socketserver import ThreadingMixIn

from app.config import load_config, save_config, update_config
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
    archive_patient_report,
    save_special_draft_queue,
    sync_review_overwrites,
    locate_patient_all_paths,
    delete_patient_all_files,
    sync_status_file_operations,
    find_patient_archive_dir,
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
    get_output_regular_patients,
    get_output_special_approved_patients,
    revoke_patient_archive_db,
    revoke_patient_approval_db,
    get_patients_for_signing_sheet
)
from app.template_manager import (
    load_text_templates,
    search_templates,
    recommend_templates,
    parse_photo_filename,
    format_pathology_diagnosis
)
from app.doc_generator import generate_report_file, update_report_file
from app.doc_preview import render_preview_html
from app.signing_sheet_generator import generate_signing_sheet_excel, render_signing_sheet_html

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WEB_DIR = os.path.join(BASE_DIR, "web")

class ThreadedHTTPServer(ThreadingMixIn, HTTPServer):
    daemon_threads = True
    allow_reuse_address = True

class PathologyRequestHandler(SimpleHTTPRequestHandler):
    """Handles HTTP requests for Static files and REST API."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=WEB_DIR, **kwargs)

    def log_message(self, format, *args):
        # Concise logging
        sys.stdout.write(f"[{datetime.now().strftime('%H:%M:%S')}] {args[0]} - {args[1]}\n")

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
        query = urllib.parse.parse_qs(parsed.query)

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
                                f'id="wizard-work-dir" class="form-control" value="{work_dir}"',
                                html_content
                            )
                        if rep_doc:
                            html_content = re.sub(
                                r'id="wizard-report-doc" class="form-control" value="[^"]*"',
                                f'id="wizard-report-doc" class="form-control" value="{rep_doc}"',
                                html_content
                            )
                        if rev_doc:
                            html_content = re.sub(
                                r'id="wizard-review-doc" class="form-control" value="[^"]*"',
                                f'id="wizard-review-doc" class="form-control" value="{rev_doc}"',
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
            mime, _ = mimetypes.guess_type(p_file)
            mime = mime or "application/octet-stream"
            try:
                with open(p_file, "rb") as f:
                    content = f.read()
                self.send_response(200)
                self.send_header("Content-Type", mime)
                self.send_header("Content-Length", str(len(content)))
                self.send_header("Cache-Control", "max-age=3600")
                self.end_headers()
                self.wfile.write(content)
            except Exception as e:
                self.send_error(500, f"Error reading image: {e}")
            return

        # Text templates search
        if path == "/api/templates":
            cat = query.get("category", ["全部"])[0]
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

        # Today's patients
        if path == "/api/patients/today":
            if not cfg.get("initialized"):
                self.send_json({"success": True, "patients": []})
                return
            patients = get_today_patients(db_path)
            self.send_json({"success": True, "patients": patients, "count": len(patients)})
            return

        # Patients list for Report Making workspace
        if path == "/api/patients/report_making":
            if not cfg.get("initialized"):
                self.send_json({"success": True, "patients": []})
                return
            d_raw = query.get("date", [None])[0]
            kw = query.get("keyword", [None])[0]
            patients = get_patients_for_report_making(db_path, date_raw=d_raw, keyword=kw)
            self.send_json({"success": True, "patients": patients, "count": len(patients)})
            return

        # Patients list for Review workspace
        if path == "/api/reviews/list":
            if not cfg.get("initialized"):
                self.send_json({"success": True, "patients": []})
                return
            inc = query.get("include_approved", ["0"])[0]
            include_approved = inc in ["1", "true", "True"]
            patients = get_patients_for_review(db_path, include_approved=include_approved)
            self.send_json({"success": True, "patients": patients, "count": len(patients)})
            return

        # Search patients
        if path == "/api/patients/search":
            if not cfg.get("initialized"):
                self.send_json({"success": True, "patients": []})
                return
            query_params = {
                "start_date": query.get("start_date", [""])[0],
                "end_date": query.get("end_date", [""])[0],
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

        # Report Output Workspace Patients (3.1)
        if path == "/api/output/patients":
            if not cfg.get("initialized"):
                self.send_json({"success": True, "patients": [], "count": 0})
                return
            out_type = query.get("type", ["regular"])[0]
            prefix_val = query.get("prefix", ["ALL"])[0]
            today_only_val = query.get("today_only", ["0"])[0] in ["1", "true", "True"]

            if out_type == "special":
                patients = get_output_special_approved_patients(db_path, prefix=prefix_val, today_only=today_only_val)
            else:
                patients = get_output_regular_patients(db_path, prefix=prefix_val, today_only=today_only_val)
            self.send_json({"success": True, "patients": patients, "count": len(patients)})
            return

        # Signing Sheet Workspace Candidates (3.2)
        if path == "/api/signing/patients":
            if not cfg.get("initialized"):
                self.send_json({"success": True, "patients": [], "count": 0})
                return
            s_date = query.get("start_date", [None])[0]
            e_date = query.get("end_date", [None])[0]
            pfx_str = query.get("prefixes", ["ALL"])[0]

            prefixes_list = None
            if pfx_str and pfx_str not in ("ALL", "所有"):
                prefixes_list = [p.strip() for p in pfx_str.split(",") if p.strip() or p == ""]

            patients = get_patients_for_signing_sheet(db_path, start_date=s_date, end_date=e_date, prefixes=prefixes_list)
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
                    photos = []
                    parsed_items = []
                    diagnoses_found = []

                    if os.path.exists(photo_dir):
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

                    auto_diag = format_pathology_diagnosis(parsed_items)
                    recs = recommend_templates(diagnoses_found)

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
                    word_path = patient.get("final_report_path")
                    if not word_path or not os.path.exists(word_path):
                        word_path = patient.get("word_path") or get_patient_word_path(work_dir, pathology_no, patient["name"])
                    preview_html = render_preview_html(patient, doc_path=word_path)
                    self.send_json({"success": True, "html": preview_html, "patient": patient})
                    return

                # All related paths for patient (Requirement 3)
                if len(parts) == 4 and parts[3] == "paths":
                    paths_list = locate_patient_all_paths(work_dir, patient)
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
                if clin_res["found"]:
                    self.send_json({
                        "success": True,
                        "found": True,
                        "path": target_path,
                        "message": f"已在资源管理器中打开患者 [{p_name}] 临床照片文件夹"
                    })
                else:
                    self.send_json({
                        "success": False,
                        "found": False,
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
            generate_report_file(templates_dir, final_report_path, data_for_doc)

            # Update database
            update_patient_report_making(db_path, p_no, {
                "pathology_diag": path_diag,
                "microscopic_desc": micro_desc,
                "selected_images": selected_images,
                "final_report_path": final_report_path,
                "is_pending_review": 1 if body.get("is_pending_review") else 0,
                "is_needs_ihc": 1 if body.get("is_needs_ihc") else 0,
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
            # Synchronize special patient draft queue if marked for review or IHC (Requirement 3)
            save_special_draft_queue(work_dir, updated_p, final_report_path)

            self.send_json({
                "success": True,
                "message": f"报告制作完成并已输出！文件: {os.path.basename(final_report_path)}",
                "final_report_path": final_report_path,
                "patient": updated_p
            })
            return

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

            # Check: prohibited from maker if pending review / IHC (Requirement 3)
            is_pending = int(patient.get("is_pending_review", 0)) == 1 or patient.get("review_status") in ("待审核", "待审核/免疫组化")
            is_ihc = int(patient.get("is_needs_ihc", 0)) == 1 or patient.get("review_status") in ("免疫组化", "待审核/免疫组化")
            if source == "maker" and (is_pending or is_ihc):
                self.send_error_json("标记为“待审核”或“免疫组化”的报告禁止在报告制作工作区直接归档！已在保存时自动放入特殊患者待审队列，请移步【报告审核】工作区完成专家审核后再进行归档。", status=400)
                return

            report_file = patient.get("final_report_path")
            if not report_file or not os.path.exists(report_file):
                report_file = patient.get("word_path")

            if not report_file or not os.path.exists(report_file):
                self.send_error_json("未找到该患者生成的Word报告，请先制作并保存报告", status=400)
                return

            arch_res = archive_patient_report(work_dir, patient, report_file, is_from_review=(source == "review"))
            if not arch_res.get("success"):
                self.send_error_json(arch_res.get("error", "归档失败"), status=400)
                return

            mark_patient_archived(db_path, p_no, arch_res["archive_dir"])
            updated_p = get_patient(db_path, p_no)

            self.send_json({
                "success": True,
                "message": f"患者 [{patient['name']}] 档案已成功打包归档！",
                "archive_dir": arch_res["archive_dir"],
                "special_paths": arch_res.get("special_paths", []),
                "patient": updated_p
            })
            return

        # 9. Save & Sync Review Overwrite (2.9 & 3)
        if path == "/api/reviews/save":
            if not cfg.get("initialized"):
                self.send_error_json("系统尚未初始化工作目录")
                return

            p_no = body.get("pathology_no", "").strip().upper()
            patient = get_patient(db_path, p_no)
            if not patient:
                self.send_error_json(f"未找到病理号为 [{p_no}] 的患者", status=404)
                return

            path_diag = body.get("pathology_diag", patient.get("pathology_diag", "")).strip()
            micro_desc = body.get("microscopic_desc", patient.get("microscopic_desc", "")).strip()
            selected_images = body.get("selected_images") or patient.get("selected_images", [])
            image_count = int(body.get("image_count") or patient.get("image_count", 1))
            biopsy_sites = body.get("biopsy_sites") or patient.get("biopsy_sites", [])

            rep_doc = body.get("report_doctor") or patient.get("report_doctor") or "张医生"
            rev_doc = body.get("review_doctor") or patient.get("review_doctor") or "李主任"

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
                "report_date": body.get("report_date") or patient.get("report_date")
            })

            templates_dir = cfg.get("templates_dir")
            generate_report_file(templates_dir, final_report_path, data_for_doc)

            # Update database
            update_patient_report_making(db_path, p_no, {
                "pathology_diag": path_diag,
                "microscopic_desc": micro_desc,
                "selected_images": selected_images,
                "final_report_path": final_report_path,
                "is_pending_review": 1 if body.get("is_pending_review") else 0,
                "is_needs_ihc": 1 if body.get("is_needs_ihc") else 0,
                "report_doctor": rep_doc,
                "review_doctor": rev_doc
            })

            # Overwrite across 统一归档/病理报告输出, 按患者归档, 特殊患者归档
            sync_res = sync_review_overwrites(work_dir, patient, final_report_path)
            updated_p = get_patient(db_path, p_no)
            save_special_draft_queue(work_dir, updated_p, final_report_path)

            self.send_json({
                "success": True,
                "message": "审核内容已保存并同步覆盖各归档目录！",
                "overwritten": sync_res.get("overwritten", []),
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

            approved_p = mark_patient_approved(db_path, p_no)
            self.send_json({
                "success": True,
                "message": f"患者 [{patient['name']}] (病理号: {p_no}) 已标注为【完成审核】！",
                "patient": approved_p
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
            generate_report_file(templates_dir, final_report_path, data_for_doc)

            # Update DB (and mark archived=2 if already archived)
            update_patient_report_making(db_path, p_no, {
                "pathology_diag": path_diag,
                "microscopic_desc": micro_desc,
                "selected_images": selected_images,
                "final_report_path": final_report_path,
                "is_pending_review": 1 if body.get("is_pending_review") else 0,
                "is_needs_ihc": 1 if body.get("is_needs_ihc") else 0,
                "report_doctor": data_for_doc["report_doctor"],
                "review_doctor": data_for_doc["review_doctor"]
            })

            # Attempt silent hardware printing of Page 1
            print_ok = False
            try:
                cmd = f'powershell -NoProfile -Command "$w = New-Object -ComObject Word.Application; $w.Visible = $false; try {{ $d = $w.Documents.Open(\'{os.path.normpath(final_report_path)}\'); $d.PrintOut([ref]$false, [ref]$false, [ref]4, [ref]$null, [ref]\'1\', [ref]\'1\'); $d.Close([ref]$false); }} finally {{ $w.Quit(); }}"'
                res = subprocess.run(cmd, shell=True, capture_output=True, timeout=8)
                print_ok = (res.returncode == 0)
            except Exception:
                print_ok = False

            updated_p = get_patient(db_path, p_no)
            self.send_json({
                "success": True,
                "message": "报告已保存！" + ("已发送第一页至默认打印机" if print_ok else "已打开打印预览窗口"),
                "direct_printed": print_ok,
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
            patient = get_patient(db_path, p_no)
            if not patient:
                self.send_error_json(f"未找到病理号为 [{p_no}] 的患者", status=404)
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
            patient = get_patient(db_path, p_no)
            if not patient:
                self.send_error_json(f"未找到病理号为 [{p_no}] 的患者", status=404)
                return

            arch_dir = find_patient_archive_dir(work_dir, p_no, patient["name"])
            if arch_dir and os.path.exists(arch_dir):
                subprocess.Popen(["explorer.exe", os.path.normpath(arch_dir)])
                self.send_json({
                    "success": True,
                    "message": f"已在资源管理器中打开归档目录: {os.path.basename(arch_dir)}",
                    "path": arch_dir
                })
            else:
                fallback = os.path.join(work_dir, "按患者归档")
                os.makedirs(fallback, exist_ok=True)
                subprocess.Popen(["explorer.exe", os.path.normpath(fallback)])
                self.send_json({
                    "success": True,
                    "message": "该患者归档目录尚未生成，已打开“按患者归档”根目录",
                    "path": fallback
                })
            return

        # 17. Batch Print Page 1 of Reports in Output List (Requirement 3.1)
        if path == "/api/output/batch_print":
            p_nos = body.get("pathology_nos", [])
            if not p_nos:
                self.send_error_json("未指定需要打印的报告病理号")
                return

            printed_list = []
            failed_list = []

            for p_no in p_nos:
                patient = get_patient(db_path, p_no)
                if not patient:
                    continue
                word_p = patient.get("final_report_path")
                if not word_p or not os.path.exists(word_p):
                    diag = patient.get("pathology_diag") or ""
                    word_p = get_final_report_path(work_dir, p_no, patient["name"], diag)

                if word_p and os.path.exists(word_p):
                    cmd = f'powershell -NoProfile -Command "$w = New-Object -ComObject Word.Application; $w.Visible = $false; try {{ $d = $w.Documents.Open(\'{os.path.normpath(word_p)}\'); $d.PrintOut([ref]$false, [ref]$false, [ref]4, [ref]$null, [ref]\'1\', [ref]\'1\'); $d.Close([ref]$false); }} finally {{ $w.Quit(); }}"'
                    try:
                        res = subprocess.run(cmd, shell=True, capture_output=True, timeout=10)
                        if res.returncode == 0:
                            printed_list.append(p_no)
                        else:
                            failed_list.append(p_no)
                    except Exception:
                        failed_list.append(p_no)
                else:
                    failed_list.append(p_no)

            direct_ok = (len(printed_list) > 0 and len(failed_list) == 0)
            self.send_json({
                "success": True,
                "total": len(p_nos),
                "printed_count": len(printed_list),
                "failed_count": len(failed_list),
                "direct_printed": direct_ok,
                "printed": printed_list,
                "failed": failed_list,
                "message": f"批量打印执行完成：已成功提交 {len(printed_list)} 份报告第1页！"
            })
            return

        # 18. Generate Signing Sheet (Requirement 3.2)
        if path == "/api/signing/generate":
            p_nos = body.get("pathology_nos", [])
            start_date = body.get("start_date")
            end_date = body.get("end_date")
            prefixes = body.get("prefixes")
            custom_filename = body.get("filename")

            if p_nos:
                patients = []
                for p_no in p_nos:
                    p = get_patient(db_path, p_no)
                    if p:
                        patients.append(p)
            else:
                patients = get_patients_for_signing_sheet(db_path, start_date=start_date, end_date=end_date, prefixes=prefixes)

            if not patients:
                self.send_error_json("所选筛选范围内无匹配的患者数据，无法生成签名单")
                return

            gen_res = generate_signing_sheet_excel(
                work_dir,
                patients,
                start_date=start_date,
                end_date=end_date,
                prefixes=prefixes,
                custom_filename=custom_filename
            )

            date_range_str = f"{start_date} 至 {end_date}" if (start_date and end_date and start_date != end_date) else (start_date or end_date or datetime.now().strftime("%Y-%m-%d"))
            html_prev = render_signing_sheet_html(patients, date_range_str=date_range_str)

            self.send_json({
                "success": True,
                "message": f"签名单已生成，包含 {len(patients)} 位患者！",
                "file_path": gen_res["file_path"],
                "filename": gen_res["filename"],
                "count": len(patients),
                "html": html_prev
            })
            return

        # 19. Print Signing Sheet (Requirement 3.2)
        if path == "/api/signing/print":
            f_path = body.get("file_path", "").strip()
            direct_printed = False
            if f_path and os.path.exists(f_path):
                cmd = f'powershell -NoProfile -Command "$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; try {{ $wb = $xl.Workbooks.Open(\'{os.path.normpath(f_path)}\'); $wb.ActiveSheet.PrintOut(); $wb.Close($false); }} finally {{ $xl.Quit(); }}"'
                try:
                    res = subprocess.run(cmd, shell=True, capture_output=True, timeout=10)
                    direct_printed = (res.returncode == 0)
                except Exception:
                    direct_printed = False

            self.send_json({
                "success": True,
                "direct_printed": direct_printed,
                "message": "签名单已发送至默认打印机！" if direct_printed else "已调起打印窗口"
            })
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

            # Handle rename of files and folders if no or name changed
            rename_res = rename_patient_files_and_folders(
                work_dir,
                old_no,
                existing["name"],
                new_no,
                new_name
            )
            body["word_path"] = rename_res["word_path"]
            body["photo_dir"] = rename_res["photo_dir"]

            # Synchronize changes into Word document
            templates_dir = cfg.get("templates_dir")
            update_report_file(templates_dir, rename_res["word_path"], body)

            # Update database record
            updated = update_patient(db_path, old_no, body)

            self.send_json({
                "success": True,
                "message": f"患者 [{new_name}] 信息及Word文件已成功更新！",
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
