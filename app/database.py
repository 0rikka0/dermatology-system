# -*- coding: utf-8 -*-
"""
Dermatology Pathology Reporting System - SQLite Database Module
Supports patient registration, report making state, reviews, and archiving.
"""
import os
import sqlite3
import json
import re
from datetime import datetime

CREATE_TABLE_SQL = """
CREATE TABLE IF NOT EXISTS patients (
    pathology_no TEXT PRIMARY KEY,
    prefix TEXT DEFAULT '',
    name TEXT NOT NULL,
    gender TEXT,
    age TEXT,
    inpatient_no TEXT DEFAULT '',
    outpatient_no TEXT DEFAULT '',
    referring_doctor TEXT DEFAULT '',
    clinical_diagnosis TEXT DEFAULT '',
    image_count INTEGER DEFAULT 1,
    biopsy_sites TEXT DEFAULT '[]',
    medical_history TEXT DEFAULT '',
    submission_date TEXT DEFAULT '',
    submission_date_raw TEXT DEFAULT '',
    report_date TEXT DEFAULT '',
    report_date_raw TEXT DEFAULT '',
    report_doctor TEXT DEFAULT '',
    review_doctor TEXT DEFAULT '',
    word_path TEXT DEFAULT '',
    photo_dir TEXT DEFAULT '',
    pathology_diag TEXT DEFAULT '',
    microscopic_desc TEXT DEFAULT '',
    selected_images TEXT DEFAULT '[]',
    final_report_path TEXT DEFAULT '',
    is_pending_review INTEGER DEFAULT 0,
    is_needs_ihc INTEGER DEFAULT 0,
    is_needs_recut INTEGER DEFAULT 0,
    send_unit TEXT DEFAULT '本院',
    department TEXT DEFAULT '皮肤科',
    review_status TEXT DEFAULT '未审核',
    archived INTEGER DEFAULT 0,
    archive_dir TEXT DEFAULT '',
    clinical_photo_dir TEXT DEFAULT '',
    created_at TEXT DEFAULT '',
    updated_at TEXT DEFAULT ''
);
"""

CREATE_INDEX_SQL = """
CREATE INDEX IF NOT EXISTS idx_patients_name ON patients(name);
CREATE INDEX IF NOT EXISTS idx_patients_inpatient ON patients(inpatient_no);
CREATE INDEX IF NOT EXISTS idx_patients_outpatient ON patients(outpatient_no);
CREATE INDEX IF NOT EXISTS idx_patients_created ON patients(created_at);
CREATE INDEX IF NOT EXISTS idx_patients_rep_date ON patients(report_date_raw);
CREATE INDEX IF NOT EXISTS idx_patients_review ON patients(review_status);
CREATE INDEX IF NOT EXISTS idx_patients_archived ON patients(archived);
"""

CREATE_IHC_TABLE_SQL = """
CREATE TABLE IF NOT EXISTS ihc_reports (
    pathology_no TEXT PRIMARY KEY,
    ihc_created INTEGER DEFAULT 0,
    ihc_dir TEXT DEFAULT '',
    ihc_doc_path TEXT DEFAULT '',
    ihc_photo_dir TEXT DEFAULT '',
    source_archive_path TEXT DEFAULT '',
    source_archive_filename TEXT DEFAULT '',
    report_doctor TEXT DEFAULT '',
    review_doctor TEXT DEFAULT '',
    report_date TEXT DEFAULT '',
    report_date_raw TEXT DEFAULT '',
    microscopic_desc TEXT DEFAULT '',
    pathology_diag TEXT DEFAULT '',
    selected_images TEXT DEFAULT '[]',
    image_count INTEGER DEFAULT 1,
    biopsy_sites TEXT DEFAULT '[]',
    archived INTEGER DEFAULT 0,
    archive_file TEXT DEFAULT '',
    created_at TEXT DEFAULT '',
    updated_at TEXT DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_ihc_rep_date ON ihc_reports(report_date_raw);
CREATE INDEX IF NOT EXISTS idx_ihc_archived ON ihc_reports(archived);
"""

def get_connection(db_path, timeout=15.0):
    """Get SQLite database connection with row factory and busy timeout."""
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    conn = sqlite3.connect(db_path, timeout=timeout)
    conn.row_factory = sqlite3.Row
    try:
        conn.execute("PRAGMA busy_timeout = 15000;")
    except Exception:
        pass
    return conn

def init_db(db_path):
    """Initialize database tables, indexes and migrate new columns."""
    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        try:
            cursor.execute("PRAGMA journal_mode = WAL;")
            cursor.execute("PRAGMA synchronous = NORMAL;")
        except Exception:
            pass
        cursor.execute(CREATE_TABLE_SQL)
        cursor.executescript(CREATE_INDEX_SQL)
        cursor.executescript(CREATE_IHC_TABLE_SQL)
        
        # Check and migrate columns if upgrading existing db
        cursor.execute("PRAGMA table_info(patients)")
        existing_cols = {row["name"] for row in cursor.fetchall()}
        
        new_columns = [
            ("pathology_diag", "TEXT DEFAULT ''"),
            ("microscopic_desc", "TEXT DEFAULT ''"),
            ("selected_images", "TEXT DEFAULT '[]'"),
            ("final_report_path", "TEXT DEFAULT ''"),
            ("is_pending_review", "INTEGER DEFAULT 0"),
            ("is_needs_ihc", "INTEGER DEFAULT 0"),
            ("is_needs_recut", "INTEGER DEFAULT 0"),
            ("send_unit", "TEXT DEFAULT '本院'"),
            ("department", "TEXT DEFAULT '皮肤科'"),
            ("review_status", "TEXT DEFAULT '未审核'"),
            ("archived", "INTEGER DEFAULT 0"),
            ("archive_dir", "TEXT DEFAULT ''"),
            ("clinical_photo_dir", "TEXT DEFAULT ''")
        ]
        
        for col_name, col_type in new_columns:
            if col_name not in existing_cols:
                cursor.execute(f"ALTER TABLE patients ADD COLUMN {col_name} {col_type}")
                
        conn.commit()

def build_review_status(is_pending=False, is_ihc=False, is_recut=False, fallback="未审核"):
    """Build combined review_status string from special flags."""
    parts = []
    if is_pending:
        parts.append("待审核")
    if is_ihc:
        parts.append("免疫组化")
    if is_recut:
        parts.append("重切")
    return "/".join(parts) if parts else fallback

def patient_row_to_dict(row):
    """Convert SQLite row to dictionary, parsing JSON fields."""
    if not row:
        return None
    d = dict(row)
    if isinstance(d.get("biopsy_sites"), str):
        try:
            d["biopsy_sites"] = json.loads(d["biopsy_sites"])
        except Exception:
            d["biopsy_sites"] = [d["biopsy_sites"]]
            
    if isinstance(d.get("selected_images"), str):
        try:
            d["selected_images"] = json.loads(d["selected_images"])
        except Exception:
            d["selected_images"] = []
    return d

def get_patient(db_path, pathology_no):
    """Retrieve patient by pathology number."""
    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM patients WHERE pathology_no = ?", (pathology_no.strip(),))
        row = cursor.fetchone()
        return patient_row_to_dict(row)

def check_duplicate(db_path, pathology_no):
    """Check if a pathology number already exists."""
    return get_patient(db_path, pathology_no)

def parse_pathology_no(pathology_no):
    """Parses pathology number format: [Prefix]YYYYMMDDNN."""
    p = str(pathology_no).strip().upper()
    match = re.match(r"^([BLHT]?)(\d{4})(\d{2})(\d{2})(\d+)$", p)
    if not match:
        return {
            "valid": False,
            "prefix": "",
            "year": "",
            "month": "",
            "day": "",
            "seq": "",
            "submission_date": "",
            "submission_date_raw": ""
        }
    prefix, y, m, d, seq = match.groups()
    return {
        "valid": True,
        "prefix": prefix,
        "year": y,
        "month": m,
        "day": d,
        "seq": seq,
        "submission_date": f"{y}年{int(m)}月{int(d)}日",
        "submission_date_raw": f"{y}-{m}-{d}"
    }

def get_next_pathology_no(db_path, prefix="", date_str=None):
    """Calculate the next available pathology number for a given prefix and date."""
    prefix = (prefix or "").strip().upper()
    now = datetime.now()
    if date_str:
        clean_date = date_str.replace("-", "").replace("/", "")[:8]
    else:
        clean_date = now.strftime("%Y%m%d")

    pattern = f"{prefix}{clean_date}%"
    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT pathology_no FROM patients WHERE pathology_no LIKE ? ORDER BY pathology_no",
            (pattern,)
        )
        rows = cursor.fetchall()

    max_seq = 0
    prefix_len = len(prefix) + len(clean_date)
    for row in rows:
        no = row["pathology_no"].strip().upper()
        suffix = no[prefix_len:]
        if suffix.isdigit():
            val = int(suffix)
            if val > max_seq:
                max_seq = val

    next_seq = max_seq + 1
    seq_str = f"{next_seq:02d}"
    return f"{prefix}{clean_date}{seq_str}"

def insert_patient(db_path, data):
    """Insert a new patient in entry workspace."""
    pathology_no = data.get("pathology_no", "").strip().upper()
    if not pathology_no:
        raise ValueError("病理号不能为空")

    existing = get_patient(db_path, pathology_no)
    if existing:
        raise ValueError(f"病理号 [{pathology_no}] 已存在！患者姓名：{existing['name']}")

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    today_raw = datetime.now().strftime("%Y-%m-%d")
    today_cn = datetime.now().strftime("%Y年%m月%d日")

    parsed = parse_pathology_no(pathology_no)
    sub_date = data.get("submission_date") or parsed["submission_date"] or today_cn
    sub_date_raw = data.get("submission_date_raw") or parsed["submission_date_raw"] or today_raw
    rep_date = data.get("report_date") or today_cn
    rep_date_raw = data.get("report_date_raw") or today_raw

    biopsy_sites = data.get("biopsy_sites", [])
    if not isinstance(biopsy_sites, list):
        if isinstance(biopsy_sites, str):
            biopsy_sites = [s.strip() for s in biopsy_sites.split(";") if s.strip()]
        else:
            biopsy_sites = []

    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO patients (
                pathology_no, prefix, name, gender, age,
                inpatient_no, outpatient_no, referring_doctor, clinical_diagnosis,
                image_count, biopsy_sites, medical_history,
                submission_date, submission_date_raw, report_date, report_date_raw,
                report_doctor, review_doctor, word_path, photo_dir,
                pathology_diag, microscopic_desc, selected_images, final_report_path,
                is_pending_review, is_needs_ihc, is_needs_recut, send_unit, department,
                review_status, archived, archive_dir,
                clinical_photo_dir, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            pathology_no,
            parsed["prefix"],
            data.get("name", "").strip(),
            data.get("gender", "女").strip(),
            str(data.get("age", "")).strip(),
            data.get("inpatient_no", "").strip(),
            data.get("outpatient_no", "").strip(),
            data.get("referring_doctor", "").strip(),
            data.get("clinical_diagnosis", "").strip(),
            int(data.get("image_count", 1)),
            json.dumps(biopsy_sites, ensure_ascii=False),
            data.get("medical_history", "").strip(),
            sub_date,
            sub_date_raw,
            rep_date,
            rep_date_raw,
            data.get("report_doctor", "").strip(),
            data.get("review_doctor", "").strip(),
            data.get("word_path", "").strip(),
            data.get("photo_dir", "").strip(),
            data.get("pathology_diag", "").strip(),
            data.get("microscopic_desc", "").strip(),
            json.dumps(data.get("selected_images", []), ensure_ascii=False),
            data.get("final_report_path", "").strip(),
            int(data.get("is_pending_review", 0)),
            int(data.get("is_needs_ihc", 0)),
            int(data.get("is_needs_recut", 0)),
            (data.get("send_unit") or "本院").strip() or "本院",
            (data.get("department") or "皮肤科").strip() or "皮肤科",
            data.get("review_status", "未审核"),
            int(data.get("archived", 0)),
            data.get("archive_dir", "").strip(),
            data.get("clinical_photo_dir", "").strip(),
            now_str,
            now_str
        ))
        conn.commit()
    return get_patient(db_path, pathology_no)

def update_patient(db_path, old_pathology_no, data):
    """Update basic patient information from entry workspace."""
    old_p = old_pathology_no.strip().upper()
    new_p = data.get("pathology_no", old_p).strip().upper()

    if new_p != old_p:
        existing = get_patient(db_path, new_p)
        if existing:
            raise ValueError(f"无法修改病理号：目标病理号 [{new_p}] 已被患者 [{existing['name']}] 使用")

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    parsed = parse_pathology_no(new_p)

    biopsy_sites = data.get("biopsy_sites", [])
    if not isinstance(biopsy_sites, list):
        if isinstance(biopsy_sites, str):
            biopsy_sites = [s.strip() for s in biopsy_sites.split(";") if s.strip()]
        else:
            biopsy_sites = []

    existing_old = get_patient(db_path, old_p)
    selected_images = data.get("selected_images", existing_old.get("selected_images", []) if existing_old else [])
    if isinstance(selected_images, str):
        try:
            selected_images = json.loads(selected_images)
        except Exception:
            selected_images = []

    final_report_path = data.get("final_report_path")
    if final_report_path is None and existing_old:
        final_report_path = existing_old.get("final_report_path", "")

    archived_val = data.get("archived")
    if archived_val is None:
        archived_val = existing_old.get("archived", 0) if existing_old else 0

    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE patients SET
                pathology_no = ?,
                prefix = ?,
                name = ?,
                gender = ?,
                age = ?,
                inpatient_no = ?,
                outpatient_no = ?,
                referring_doctor = ?,
                clinical_diagnosis = ?,
                image_count = ?,
                biopsy_sites = ?,
                medical_history = ?,
                submission_date = ?,
                submission_date_raw = ?,
                send_unit = ?,
                department = ?,
                report_doctor = ?,
                review_doctor = ?,
                word_path = ?,
                photo_dir = ?,
                final_report_path = ?,
                selected_images = ?,
                archived = ?,
                updated_at = ?
            WHERE pathology_no = ?
        """, (
            new_p,
            parsed["prefix"],
            data.get("name", "").strip(),
            data.get("gender", "女").strip(),
            str(data.get("age", "")).strip(),
            data.get("inpatient_no", "").strip(),
            data.get("outpatient_no", "").strip(),
            data.get("referring_doctor", "").strip(),
            data.get("clinical_diagnosis", "").strip(),
            int(data.get("image_count", 1)),
            json.dumps(biopsy_sites, ensure_ascii=False),
            data.get("medical_history", "").strip(),
            data.get("submission_date") or parsed["submission_date"],
            data.get("submission_date_raw") or parsed["submission_date_raw"],
            (data.get("send_unit") or "本院").strip() or "本院",
            (data.get("department") or "皮肤科").strip() or "皮肤科",
            data.get("report_doctor", "").strip(),
            data.get("review_doctor", "").strip(),
            data.get("word_path", "").strip(),
            data.get("photo_dir", "").strip(),
            str(final_report_path or "").strip(),
            json.dumps(selected_images, ensure_ascii=False),
            int(archived_val),
            now_str,
            old_p
        ))
        if new_p != old_p:
            cursor.execute(
                "UPDATE ihc_reports SET pathology_no = ?, updated_at = ? WHERE pathology_no = ?",
                (new_p, now_str, old_p)
            )
        conn.commit()
    return get_patient(db_path, new_p)

def update_patient_report_making(db_path, pathology_no, data):
    """
    Updates patient data from the Report Making workspace:
    microscopic description, pathological diagnosis, selected photos,
    pending review, IHC, final report path.
    If the patient was already archived (archived == 1), updates status to 2 (需要重新归档).
    """
    p_no = pathology_no.strip().upper()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    existing = get_patient(db_path, p_no)
    archived_val = existing.get("archived", 0) if existing else 0
    # If already archived, mark as needing re-archive (2)
    if archived_val in [1, 2]:
        archived_val = 2

    is_pending = 1 if data.get("is_pending_review") else 0
    is_ihc = 1 if data.get("is_needs_ihc") else 0
    is_recut = 1 if data.get("is_needs_recut") else 0

    # Calculate review_status
    if existing and existing.get("review_status") == "完成审核" and not (is_pending or is_recut):
        rev_status = "完成审核"
    else:
        fallback_status = data.get("review_status", existing.get("review_status", "已制作") if existing else "已制作")
        if not (is_pending or is_ihc or is_recut) and existing and existing.get("review_status") in ("待审核", "免疫组化", "重切", "待审核/免疫组化", "待审核/重切", "免疫组化/重切", "待审核/免疫组化/重切"):
            fallback_status = "未审核"
        rev_status = build_review_status(is_pending, is_ihc, is_recut, fallback=fallback_status)

    selected_images = data.get("selected_images", [])
    if isinstance(selected_images, str):
        try:
            selected_images = json.loads(selected_images)
        except Exception:
            selected_images = []

    rep_date = str(data.get("report_date") or "").strip()
    rep_date_raw = str(data.get("report_date_raw") or "").strip()

    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE patients SET
                pathology_diag = ?,
                microscopic_desc = ?,
                selected_images = ?,
                final_report_path = ?,
                is_pending_review = ?,
                is_needs_ihc = ?,
                is_needs_recut = ?,
                review_status = ?,
                archived = ?,
                report_doctor = coalesce(nullif(?, ''), report_doctor),
                review_doctor = coalesce(nullif(?, ''), review_doctor),
                report_date = coalesce(nullif(?, ''), report_date),
                report_date_raw = coalesce(nullif(?, ''), report_date_raw),
                updated_at = ?
            WHERE pathology_no = ?
        """, (
            data.get("pathology_diag", "").strip(),
            data.get("microscopic_desc", "").strip(),
            json.dumps(selected_images, ensure_ascii=False),
            data.get("final_report_path", "").strip(),
            is_pending,
            is_ihc,
            is_recut,
            rev_status,
            archived_val,
            data.get("report_doctor", "").strip(),
            data.get("review_doctor", "").strip(),
            rep_date,
            rep_date_raw,
            now_str,
            p_no
        ))
        conn.commit()
    return get_patient(db_path, p_no)

def update_patient_status(db_path, pathology_no, review_status=None, archived=None):
    """
    Manually update review status and/or archived status from management console.
    """
    p_no = pathology_no.strip().upper()
    existing = get_patient(db_path, p_no)
    if not existing:
        return None

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_rev = review_status if review_status is not None else existing.get("review_status", "未审核")
    new_arch = int(archived) if archived is not None else existing.get("archived", 0)

    is_pending = 1 if "待审核" in new_rev else 0
    is_ihc = 1 if "免疫组化" in new_rev else 0
    is_recut = 1 if "重切" in new_rev else 0

    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE patients SET
                review_status = ?,
                archived = ?,
                is_pending_review = ?,
                is_needs_ihc = ?,
                is_needs_recut = ?,
                updated_at = ?
            WHERE pathology_no = ?
        """, (new_rev, new_arch, is_pending, is_ihc, is_recut, now_str, p_no))
        conn.commit()
    return get_patient(db_path, p_no)

def delete_patient(db_path, pathology_no):
    """Delete a patient record from SQLite database."""
    p_no = pathology_no.strip().upper()
    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM ihc_reports WHERE pathology_no = ?", (p_no,))
        cursor.execute("DELETE FROM patients WHERE pathology_no = ?", (p_no,))
        conn.commit()
    return True

def mark_patient_archived(db_path, pathology_no, archive_dir):
    """Mark a patient as archived with the archive folder path."""
    p_no = pathology_no.strip().upper()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE patients SET
                archived = 1,
                archive_dir = ?,
                updated_at = ?
            WHERE pathology_no = ?
        """, (archive_dir, now_str, p_no))
        conn.commit()
    return get_patient(db_path, p_no)

def mark_patient_approved(db_path, pathology_no):
    """Mark patient review as complete (完成审核)."""
    p_no = pathology_no.strip().upper()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE patients SET
                review_status = '完成审核',
                is_pending_review = 0,
                updated_at = ?
            WHERE pathology_no = ?
        """, (now_str, p_no))
        conn.commit()
    return get_patient(db_path, p_no)

SUB_DATE_EXPR = "COALESCE(NULLIF(submission_date_raw, ''), NULLIF(report_date_raw, ''), substr(created_at, 1, 10))"
REP_DATE_EXPR = "COALESCE(NULLIF(report_date_raw, ''), substr(created_at, 1, 10))"


def _append_date_filter(conditions, values, date_expr, single_date=None, start_date=None, end_date=None):
    """Append single-date or inclusive date-range SQL conditions for date_expr."""
    single = str(single_date or "").strip()
    if single:
        conditions.append(f"({date_expr} = ?)")
        values.append(single)
        return
    s_date = str(start_date or "").strip()
    e_date = str(end_date or "").strip()
    if s_date and e_date:
        conditions.append(f"({date_expr} BETWEEN ? AND ?)")
        values.extend([s_date, e_date])
    elif s_date:
        conditions.append(f"({date_expr} >= ?)")
        values.append(s_date)
    elif e_date:
        conditions.append(f"({date_expr} <= ?)")
        values.append(e_date)


def _append_prefix_filter(conditions, values, prefix=None, prefixes=None):
    """Append single-prefix or multi-prefix SQL conditions."""
    if prefixes is not None:
        if "ALL" in prefixes or "所有" in prefixes:
            return
        prefix_conds = []
        for pfx in prefixes:
            pfx_clean = str(pfx).strip().upper()
            if pfx_clean in ("", "NONE", "普通", "无前缀"):
                prefix_conds.append("(prefix = '' OR prefix IS NULL)")
            else:
                prefix_conds.append("UPPER(prefix) = ?")
                values.append(pfx_clean)
        if prefix_conds:
            conditions.append("(" + " OR ".join(prefix_conds) + ")")
        else:
            # Empty non-None list explicitly checked -> match nothing
            conditions.append("1 = 0")
        return

    pfx_val = str(prefix or "").strip().upper()
    if pfx_val and pfx_val not in ("ALL", "所有"):
        if pfx_val in ("NONE", "普通", "无前缀"):
            conditions.append("(prefix = '' OR prefix IS NULL)")
        else:
            conditions.append("UPPER(prefix) = ?")
            values.append(pfx_val)


def _append_keyword_filter(conditions, values, keyword=None, full_fields=False):
    """Append keyword LIKE filter across patient name, pathology_no, and key fields."""
    kw = str(keyword or "").strip()
    if not kw:
        return
    kw_like = f"%{kw}%"
    if full_fields:
        conditions.append(
            "(name LIKE ? OR pathology_no LIKE ? OR inpatient_no LIKE ? OR outpatient_no LIKE ? "
            "OR clinical_diagnosis LIKE ? OR biopsy_sites LIKE ? OR referring_doctor LIKE ? OR pathology_diag LIKE ?)"
        )
        values.extend([kw_like] * 8)
    else:
        conditions.append(
            "(name LIKE ? OR pathology_no LIKE ? OR inpatient_no LIKE ? OR outpatient_no LIKE ? "
            "OR clinical_diagnosis LIKE ? OR pathology_diag LIKE ?)"
        )
        values.extend([kw_like] * 6)


def get_today_patients(db_path, today_raw=None, report_date=None, submission_date=None, keyword=None):
    """
    Retrieve patients for the Entry workspace patient list.
    When called with no explicit filter parameters, defaults to today's report date.
    Passing report_date="" explicitly disables the report date filter (all dates).
    """
    conditions = []
    values = []

    if report_date is None and submission_date is None and keyword is None:
        effective_rep = today_raw or datetime.now().strftime("%Y-%m-%d")
        _append_date_filter(conditions, values, REP_DATE_EXPR, single_date=effective_rep)
    else:
        _append_date_filter(conditions, values, REP_DATE_EXPR, single_date=report_date)
        _append_date_filter(conditions, values, SUB_DATE_EXPR, single_date=submission_date)
        _append_keyword_filter(conditions, values, keyword=keyword)

    sql = "SELECT * FROM patients"
    if conditions:
        sql += " WHERE " + " AND ".join(conditions)
    sql += " ORDER BY pathology_no ASC LIMIT 200"

    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute(sql, values)
        rows = cursor.fetchall()
        return [patient_row_to_dict(r) for r in rows]


def get_patients_for_report_making(db_path, date_raw=None, keyword=None, report_date=None, submission_date=None):
    """
    Retrieve patients for Report Making workspace.
    Supports filtering by report_date, submission_date, and search keyword.
    """
    conditions = []
    values = []

    if report_date is not None or submission_date is not None:
        _append_date_filter(conditions, values, REP_DATE_EXPR, single_date=report_date)
        _append_date_filter(conditions, values, SUB_DATE_EXPR, single_date=submission_date)
    elif date_raw:
        conditions.append("(report_date_raw = ? OR submission_date_raw = ? OR created_at LIKE ?)")
        values.extend([date_raw, date_raw, f"{date_raw}%"])

    _append_keyword_filter(conditions, values, keyword=keyword)

    sql = "SELECT * FROM patients"
    if conditions:
        sql += " WHERE " + " AND ".join(conditions)
    sql += " ORDER BY pathology_no ASC LIMIT 200"

    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute(sql, values)
        rows = cursor.fetchall()
        return [patient_row_to_dict(r) for r in rows]


def get_patients_for_review(db_path, include_approved=False, report_date=None, submission_date=None, keyword=None):
    """
    Retrieve patients for Report Review workspace.
    Includes pending review and recut patients (excluding pure IHC patients),
    filterable by report/submission dates and keyword.
    """
    if include_approved:
        conditions = [
            "(is_pending_review = 1 OR is_needs_recut = 1 "
            "OR review_status LIKE '%待审核%' OR review_status LIKE '%重切%' "
            "OR review_status = '完成审核')"
        ]
    else:
        conditions = [
            "(is_pending_review = 1 OR is_needs_recut = 1 "
            "OR review_status LIKE '%待审核%' OR review_status LIKE '%重切%')",
            "review_status != '完成审核'"
        ]
    values = []

    _append_date_filter(conditions, values, REP_DATE_EXPR, single_date=report_date)
    _append_date_filter(conditions, values, SUB_DATE_EXPR, single_date=submission_date)
    _append_keyword_filter(conditions, values, keyword=keyword)

    sql = "SELECT * FROM patients WHERE " + " AND ".join(conditions) + " ORDER BY pathology_no ASC"
    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute(sql, values)
        rows = cursor.fetchall()
        return [patient_row_to_dict(r) for r in rows]


def ihc_row_to_dict(row):
    """Convert SQLite ihc_reports row to dictionary, parsing JSON fields."""
    if not row:
        return None
    d = dict(row)
    if isinstance(d.get("biopsy_sites"), str):
        try:
            d["biopsy_sites"] = json.loads(d["biopsy_sites"])
        except Exception:
            d["biopsy_sites"] = [d["biopsy_sites"]] if d["biopsy_sites"] else []
    if isinstance(d.get("selected_images"), str):
        try:
            d["selected_images"] = json.loads(d["selected_images"])
        except Exception:
            d["selected_images"] = []
    return d


def get_ihc_report(db_path, pathology_no):
    """Retrieve the independent IHC report record for a patient."""
    p_no = str(pathology_no or "").strip().upper()
    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM ihc_reports WHERE pathology_no = ?", (p_no,))
        return ihc_row_to_dict(cursor.fetchone())


def upsert_ihc_report(db_path, pathology_no, data, mode="create"):
    """Create or update an independent IHC report record without altering patients table."""
    p_no = str(pathology_no or "").strip().upper()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    existing = get_ihc_report(db_path, p_no)

    biopsy_sites = data.get("biopsy_sites")
    if biopsy_sites is None:
        biopsy_sites = existing.get("biopsy_sites", []) if existing else []
    elif isinstance(biopsy_sites, str):
        try:
            biopsy_sites = json.loads(biopsy_sites)
        except Exception:
            biopsy_sites = [s.strip() for s in biopsy_sites.split(";") if s.strip()]

    selected_images = data.get("selected_images")
    if selected_images is None:
        selected_images = existing.get("selected_images", []) if existing else []
    elif isinstance(selected_images, str):
        try:
            selected_images = json.loads(selected_images)
        except Exception:
            selected_images = []

    if mode == "save" and existing:
        arch_val = 2 if existing.get("archived", 0) in (1, 2) else 0
        arch_file = existing.get("archive_file", "")
    else:
        arch_val = int(data.get("archived", 0))
        arch_file = str(data.get("archive_file", "")).strip()

    ihc_created = int(data.get("ihc_created", existing.get("ihc_created", 1) if existing else 1))
    ihc_dir = str(data.get("ihc_dir", existing.get("ihc_dir", "") if existing else "")).strip()
    ihc_doc_path = str(data.get("ihc_doc_path", existing.get("ihc_doc_path", "") if existing else "")).strip()
    ihc_photo_dir = str(data.get("ihc_photo_dir", existing.get("ihc_photo_dir", "") if existing else "")).strip()
    src_arch_path = str(data.get("source_archive_path", existing.get("source_archive_path", "") if existing else "")).strip()
    src_arch_name = str(data.get("source_archive_filename", existing.get("source_archive_filename", "") if existing else "")).strip()

    rep_doc = str(data.get("report_doctor", existing.get("report_doctor", "") if existing else "")).strip()
    rev_doc = str(data.get("review_doctor", existing.get("review_doctor", "") if existing else "")).strip()
    rep_date = str(data.get("report_date", existing.get("report_date", "") if existing else "")).strip()
    rep_date_raw = str(data.get("report_date_raw", existing.get("report_date_raw", "") if existing else "")).strip()
    micro_desc = str(data.get("microscopic_desc", existing.get("microscopic_desc", "") if existing else "")).strip()
    patho_diag = str(data.get("pathology_diag", existing.get("pathology_diag", "") if existing else "")).strip()
    img_cnt = int(data.get("image_count", existing.get("image_count", 1) if existing else 1) or 1)
    created_at = existing.get("created_at", now_str) if existing else now_str

    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO ihc_reports (
                pathology_no, ihc_created, ihc_dir, ihc_doc_path, ihc_photo_dir,
                source_archive_path, source_archive_filename,
                report_doctor, review_doctor, report_date, report_date_raw,
                microscopic_desc, pathology_diag, selected_images, image_count, biopsy_sites,
                archived, archive_file, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(pathology_no) DO UPDATE SET
                ihc_created = excluded.ihc_created,
                ihc_dir = excluded.ihc_dir,
                ihc_doc_path = excluded.ihc_doc_path,
                ihc_photo_dir = excluded.ihc_photo_dir,
                source_archive_path = excluded.source_archive_path,
                source_archive_filename = excluded.source_archive_filename,
                report_doctor = excluded.report_doctor,
                review_doctor = excluded.review_doctor,
                report_date = excluded.report_date,
                report_date_raw = excluded.report_date_raw,
                microscopic_desc = excluded.microscopic_desc,
                pathology_diag = excluded.pathology_diag,
                selected_images = excluded.selected_images,
                image_count = excluded.image_count,
                biopsy_sites = excluded.biopsy_sites,
                archived = excluded.archived,
                archive_file = excluded.archive_file,
                updated_at = excluded.updated_at
        """, (
            p_no, ihc_created, ihc_dir, ihc_doc_path, ihc_photo_dir,
            src_arch_path, src_arch_name,
            rep_doc, rev_doc, rep_date, rep_date_raw,
            micro_desc, patho_diag,
            json.dumps(selected_images, ensure_ascii=False),
            img_cnt,
            json.dumps(biopsy_sites, ensure_ascii=False),
            arch_val, arch_file, created_at, now_str
        ))
        conn.commit()
    return get_ihc_report(db_path, p_no)


def mark_ihc_archived(db_path, pathology_no, archive_file):
    """Mark an IHC report as archived in ihc_reports table."""
    p_no = str(pathology_no or "").strip().upper()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE ihc_reports SET
                archived = 1,
                archive_file = ?,
                updated_at = ?
            WHERE pathology_no = ?
        """, (archive_file, now_str, p_no))
        conn.commit()
    return get_ihc_report(db_path, p_no)


def revoke_ihc_archive_db(db_path, pathology_no):
    """Revoke archive status of an IHC report in ihc_reports table."""
    p_no = str(pathology_no or "").strip().upper()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE ihc_reports SET
                archived = 0,
                archive_file = '',
                updated_at = ?
            WHERE pathology_no = ?
        """, (now_str, p_no))
        conn.commit()
    return get_ihc_report(db_path, p_no)


def update_ihc_paths(db_path, pathology_no, ihc_dir, ihc_doc_path, ihc_photo_dir,
                     selected_images=None, mark_rearchive=False):
    """Update IHC directory/document/photo paths after patient rename."""
    p_no = str(pathology_no or "").strip().upper()
    existing = get_ihc_report(db_path, p_no)
    if not existing:
        return None
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    arch_val = existing.get("archived", 0)
    if mark_rearchive and arch_val in (1, 2):
        arch_val = 2
    sel_imgs = selected_images if selected_images is not None else existing.get("selected_images", [])
    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE ihc_reports SET
                ihc_dir = ?,
                ihc_doc_path = ?,
                ihc_photo_dir = ?,
                selected_images = ?,
                archived = ?,
                updated_at = ?
            WHERE pathology_no = ?
        """, (
            ihc_dir,
            ihc_doc_path,
            ihc_photo_dir,
            json.dumps(sel_imgs, ensure_ascii=False),
            int(arch_val),
            now_str,
            p_no
        ))
        conn.commit()
    return get_ihc_report(db_path, p_no)


def delete_ihc_report_and_clear_status(db_path, pathology_no):
    """
    Delete the independent IHC report row from ihc_reports and completely clear the
    patient's IHC flag/status in patients table so it becomes an ordinary report.
    """
    p_no = str(pathology_no or "").strip().upper()
    existing = get_patient(db_path, p_no)
    if not existing:
        return None

    old_rev = str(existing.get("review_status", "") or "").strip()
    is_pending = 1 if (existing.get("is_pending_review") == 1 or "待审核" in old_rev) else 0
    is_recut = 1 if (existing.get("is_needs_recut") == 1 or "重切" in old_rev) else 0

    if old_rev == "完成审核" and not is_pending and not is_recut:
        new_rev = "完成审核"
    elif is_pending or is_recut:
        new_rev = build_review_status(bool(is_pending), False, bool(is_recut))
    else:
        new_rev = "已制作" if str(existing.get("final_report_path") or "").strip() else "未审核"

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM ihc_reports WHERE pathology_no = ?", (p_no,))
        cursor.execute("""
            UPDATE patients SET
                is_needs_ihc = 0,
                review_status = ?,
                updated_at = ?
            WHERE pathology_no = ?
        """, (new_rev, now_str, p_no))
        conn.commit()
    return get_patient(db_path, p_no)


def _extract_joined_ihc_dict(row_dict):
    """Pop ihc_* aliased columns from joined query row into a nested ihc_report dict."""
    ihc_keys = {
        "ihc_created": "ihc_created",
        "ihc_dir": "ihc_dir",
        "ihc_doc_path": "ihc_doc_path",
        "ihc_photo_dir": "ihc_photo_dir",
        "ihc_source_archive_path": "source_archive_path",
        "ihc_source_archive_filename": "source_archive_filename",
        "ihc_report_doctor": "report_doctor",
        "ihc_review_doctor": "review_doctor",
        "ihc_report_date": "report_date",
        "ihc_report_date_raw": "report_date_raw",
        "ihc_microscopic_desc": "microscopic_desc",
        "ihc_pathology_diag": "pathology_diag",
        "ihc_selected_images": "selected_images",
        "ihc_image_count": "image_count",
        "ihc_biopsy_sites": "biopsy_sites",
        "ihc_archived": "archived",
        "ihc_archive_file": "archive_file"
    }
    ihc = {}
    for alias, target in ihc_keys.items():
        if alias in row_dict:
            ihc[target] = row_dict.pop(alias)
    ihc["pathology_no"] = row_dict.get("pathology_no", "")
    ihc["ihc_created"] = int(ihc.get("ihc_created") or 0)
    ihc["archived"] = int(ihc.get("archived") or 0)
    ihc["image_count"] = int(ihc.get("image_count") or row_dict.get("image_count") or 1)
    for json_field in ("selected_images", "biopsy_sites"):
        val = ihc.get(json_field)
        if isinstance(val, str):
            try:
                ihc[json_field] = json.loads(val)
            except Exception:
                ihc[json_field] = []
        elif val is None:
            ihc[json_field] = []
    row_dict["ihc_report"] = ihc
    return row_dict


def get_patients_for_ihc(db_path, report_date=None, submission_date=None, keyword=None):
    """
    Retrieve patients tagged with 免疫组化 (or having created IHC reports) for the IHC workspace.
    """
    conditions = [
        "(p.is_needs_ihc = 1 OR p.review_status LIKE '%免疫组化%' OR COALESCE(ihc.ihc_created, 0) = 1)"
    ]
    values = []

    ihc_rep_expr = "COALESCE(NULLIF(ihc.report_date_raw, ''), NULLIF(p.report_date_raw, ''), substr(p.created_at, 1, 10))"
    p_sub_expr = "COALESCE(NULLIF(p.submission_date_raw, ''), NULLIF(p.report_date_raw, ''), substr(p.created_at, 1, 10))"
    _append_date_filter(conditions, values, ihc_rep_expr, single_date=report_date)
    _append_date_filter(conditions, values, p_sub_expr, single_date=submission_date)

    kw = str(keyword or "").strip()
    if kw:
        kw_like = f"%{kw}%"
        conditions.append(
            "(p.name LIKE ? OR p.pathology_no LIKE ? OR p.inpatient_no LIKE ? OR p.outpatient_no LIKE ? "
            "OR p.clinical_diagnosis LIKE ? OR p.pathology_diag LIKE ? OR COALESCE(ihc.pathology_diag, '') LIKE ?)"
        )
        values.extend([kw_like] * 7)

    sql = """
        SELECT p.*,
               ihc.ihc_created AS ihc_created,
               ihc.ihc_dir AS ihc_dir,
               ihc.ihc_doc_path AS ihc_doc_path,
               ihc.ihc_photo_dir AS ihc_photo_dir,
               ihc.source_archive_path AS ihc_source_archive_path,
               ihc.source_archive_filename AS ihc_source_archive_filename,
               ihc.report_doctor AS ihc_report_doctor,
               ihc.review_doctor AS ihc_review_doctor,
               ihc.report_date AS ihc_report_date,
               ihc.report_date_raw AS ihc_report_date_raw,
               ihc.microscopic_desc AS ihc_microscopic_desc,
               ihc.pathology_diag AS ihc_pathology_diag,
               ihc.selected_images AS ihc_selected_images,
               ihc.image_count AS ihc_image_count,
               ihc.biopsy_sites AS ihc_biopsy_sites,
               ihc.archived AS ihc_archived,
               ihc.archive_file AS ihc_archive_file
        FROM patients p
        LEFT JOIN ihc_reports ihc ON p.pathology_no = ihc.pathology_no
        WHERE """ + " AND ".join(conditions) + " ORDER BY p.pathology_no ASC"

    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute(sql, values)
        rows = cursor.fetchall()
        return [_extract_joined_ihc_dict(patient_row_to_dict(r)) for r in rows]


def search_patients(db_path, query_params):
    """Search patients across all criteria including dates, prefix, archived and review status."""
    conditions = []
    values = []

    # 1. Date range filter (by submission date)
    _append_date_filter(
        conditions,
        values,
        SUB_DATE_EXPR,
        start_date=query_params.get("start_date", ""),
        end_date=query_params.get("end_date", "")
    )

    # 1b. Report date range filter
    _append_date_filter(
        conditions,
        values,
        REP_DATE_EXPR,
        start_date=query_params.get("report_start_date", ""),
        end_date=query_params.get("report_end_date", "")
    )

    # 2. Prefix filter
    _append_prefix_filter(conditions, values, prefix=query_params.get("prefix", ""))

    # 3. Archive status filter
    archived_val = str(query_params.get("archived", "")).strip().upper()
    if archived_val != "" and archived_val not in ("ALL", "所有"):
        try:
            conditions.append("archived = ?")
            values.append(int(archived_val))
        except ValueError:
            pass

    # 4. Review status filter
    rev_status = str(query_params.get("review_status", "")).strip()
    if rev_status and rev_status not in ("ALL", "所有"):
        if rev_status in ("普通", "无需审核"):
            conditions.append("(review_status = '无需审核' OR review_status = '' OR review_status IS NULL)")
        elif rev_status == "待审核":
            conditions.append("(review_status LIKE '%待审核%')")
        elif rev_status == "免疫组化":
            conditions.append("(review_status LIKE '%免疫组化%' OR is_needs_ihc = 1)")
        elif rev_status == "重切":
            conditions.append("(review_status LIKE '%重切%')")
        elif rev_status == "完成审核":
            conditions.append("review_status = '完成审核'")
        else:
            conditions.append("review_status = ?")
            values.append(rev_status)

    _append_keyword_filter(conditions, values, keyword=query_params.get("keyword", ""), full_fields=True)

    name = query_params.get("name", "").strip()
    if name:
        conditions.append("name LIKE ?")
        values.append(f"%{name}%")

    pathology_no = query_params.get("pathology_no", "").strip()
    if pathology_no:
        conditions.append("pathology_no LIKE ?")
        values.append(f"%{pathology_no}%")

    inpatient_no = query_params.get("inpatient_no", "").strip()
    if inpatient_no:
        conditions.append("inpatient_no LIKE ?")
        values.append(f"%{inpatient_no}%")

    outpatient_no = query_params.get("outpatient_no", "").strip()
    if outpatient_no:
        conditions.append("outpatient_no LIKE ?")
        values.append(f"%{outpatient_no}%")

    sql = "SELECT * FROM patients"
    if conditions:
        sql += " WHERE " + " AND ".join(conditions)
    sql += " ORDER BY created_at DESC, pathology_no DESC LIMIT 300"

    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute(sql, values)
        rows = cursor.fetchall()
        return [patient_row_to_dict(r) for r in rows]


def _query_output_patients(db_path, base_conditions, prefix=None, today_only=False, today_str=None,
                           prefixes=None, report_start_date=None, report_end_date=None,
                           sub_start_date=None, sub_end_date=None):
    """Shared helper for querying archived output lists with date-range and prefix filters."""
    if not today_str:
        today_str = datetime.now().strftime("%Y-%m-%d")

    conditions = list(base_conditions)
    values = []

    _append_prefix_filter(conditions, values, prefix=prefix, prefixes=prefixes)

    if today_only:
        _append_date_filter(conditions, values, REP_DATE_EXPR, single_date=today_str)
    else:
        _append_date_filter(conditions, values, REP_DATE_EXPR, start_date=report_start_date, end_date=report_end_date)
        _append_date_filter(conditions, values, SUB_DATE_EXPR, start_date=sub_start_date, end_date=sub_end_date)

    sql = "SELECT * FROM patients WHERE " + " AND ".join(conditions) + " ORDER BY pathology_no DESC"
    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute(sql, values)
        return [patient_row_to_dict(r) for r in cursor.fetchall()]


def get_output_regular_patients(db_path, prefix=None, today_only=False, today_str=None,
                                prefixes=None, report_start_date=None, report_end_date=None,
                                sub_start_date=None, sub_end_date=None):
    """
    Retrieve archived regular patients (List 1 in Report Output).
    archived in (1, 2) AND not pending review / not recut (IHC-only patients follow regular flow).
    """
    base_conditions = [
        "archived IN (1, 2)",
        "is_pending_review = 0",
        "is_needs_recut = 0",
        "(review_status IS NULL OR review_status IN ('', '无', '未审核', '已制作', '免疫组化'))"
    ]
    return _query_output_patients(
        db_path, base_conditions, prefix=prefix, today_only=today_only, today_str=today_str,
        prefixes=prefixes, report_start_date=report_start_date, report_end_date=report_end_date,
        sub_start_date=sub_start_date, sub_end_date=sub_end_date
    )


def get_output_special_approved_patients(db_path, prefix=None, today_only=False, today_str=None,
                                         prefixes=None, report_start_date=None, report_end_date=None,
                                         sub_start_date=None, sub_end_date=None):
    """
    Retrieve archived approved special patients (List 2 in Report Output).
    archived in (1, 2) AND review_status == '完成审核'.
    """
    base_conditions = [
        "archived IN (1, 2)",
        "review_status = '完成审核'"
    ]
    return _query_output_patients(
        db_path, base_conditions, prefix=prefix, today_only=today_only, today_str=today_str,
        prefixes=prefixes, report_start_date=report_start_date, report_end_date=report_end_date,
        sub_start_date=sub_start_date, sub_end_date=sub_end_date
    )


def get_output_ihc_patients(db_path, prefix=None, today_only=False, today_str=None,
                            prefixes=None, report_start_date=None, report_end_date=None,
                            sub_start_date=None, sub_end_date=None):
    """
    Retrieve archived independent IHC reports (List 3 in Report Output).
    """
    if not today_str:
        today_str = datetime.now().strftime("%Y-%m-%d")

    conditions = ["ihc.archived IN (1, 2)"]
    values = []

    if prefixes is not None:
        if "ALL" not in prefixes and "所有" not in prefixes:
            prefix_conds = []
            for pfx in prefixes:
                pfx_clean = str(pfx).strip().upper()
                if pfx_clean in ("", "NONE", "普通", "无前缀"):
                    prefix_conds.append("(p.prefix = '' OR p.prefix IS NULL)")
                else:
                    prefix_conds.append("UPPER(p.prefix) = ?")
                    values.append(pfx_clean)
            if prefix_conds:
                conditions.append("(" + " OR ".join(prefix_conds) + ")")
            else:
                conditions.append("1 = 0")
    else:
        pfx_val = str(prefix or "").strip().upper()
        if pfx_val and pfx_val not in ("ALL", "所有"):
            if pfx_val in ("NONE", "普通", "无前缀"):
                conditions.append("(p.prefix = '' OR p.prefix IS NULL)")
            else:
                conditions.append("UPPER(p.prefix) = ?")
                values.append(pfx_val)

    ihc_rep_expr = "COALESCE(NULLIF(ihc.report_date_raw, ''), NULLIF(p.report_date_raw, ''), substr(p.created_at, 1, 10))"
    p_sub_expr = "COALESCE(NULLIF(p.submission_date_raw, ''), NULLIF(p.report_date_raw, ''), substr(p.created_at, 1, 10))"
    if today_only:
        _append_date_filter(conditions, values, ihc_rep_expr, single_date=today_str)
    else:
        _append_date_filter(conditions, values, ihc_rep_expr, start_date=report_start_date, end_date=report_end_date)
        _append_date_filter(conditions, values, p_sub_expr, start_date=sub_start_date, end_date=sub_end_date)

    sql = """
        SELECT p.*,
               ihc.ihc_created AS ihc_created,
               ihc.ihc_dir AS ihc_dir,
               ihc.ihc_doc_path AS ihc_doc_path,
               ihc.ihc_photo_dir AS ihc_photo_dir,
               ihc.source_archive_path AS ihc_source_archive_path,
               ihc.source_archive_filename AS ihc_source_archive_filename,
               ihc.report_doctor AS ihc_report_doctor,
               ihc.review_doctor AS ihc_review_doctor,
               ihc.report_date AS ihc_report_date,
               ihc.report_date_raw AS ihc_report_date_raw,
               ihc.microscopic_desc AS ihc_microscopic_desc,
               ihc.pathology_diag AS ihc_pathology_diag,
               ihc.selected_images AS ihc_selected_images,
               ihc.image_count AS ihc_image_count,
               ihc.biopsy_sites AS ihc_biopsy_sites,
               ihc.archived AS ihc_archived,
               ihc.archive_file AS ihc_archive_file
        FROM patients p
        INNER JOIN ihc_reports ihc ON p.pathology_no = ihc.pathology_no
        WHERE """ + " AND ".join(conditions) + " ORDER BY p.pathology_no DESC"

    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute(sql, values)
        rows = cursor.fetchall()
        results = []
        for r in rows:
            d = _extract_joined_ihc_dict(patient_row_to_dict(r))
            ihc = d["ihc_report"]
            d["report_kind"] = "ihc"
            d["report_doctor"] = ihc.get("report_doctor") or d.get("report_doctor", "")
            d["review_doctor"] = ihc.get("review_doctor") or d.get("review_doctor", "")
            d["report_date"] = ihc.get("report_date") or d.get("report_date", "")
            d["report_date_raw"] = ihc.get("report_date_raw") or d.get("report_date_raw", "")
            d["microscopic_desc"] = ihc.get("microscopic_desc", "")
            d["pathology_diag"] = ihc.get("pathology_diag") or d.get("pathology_diag", "")
            d["final_report_path"] = ihc.get("ihc_doc_path", "")
            d["archive_dir"] = ihc.get("archive_file", "")
            d["archived"] = ihc.get("archived", 0)
            results.append(d)
        return results

def revoke_patient_archive_db(db_path, pathology_no):
    """Revoke archive status for a patient in database."""
    p_no = pathology_no.strip().upper()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE patients SET
                archived = 0,
                archive_dir = '',
                updated_at = ?
            WHERE pathology_no = ?
        """, (now_str, p_no))
        conn.commit()
    return get_patient(db_path, p_no)

def revoke_patient_approval_db(db_path, pathology_no):
    """Revoke approval status for a special patient in database, returning to review queue."""
    p_no = pathology_no.strip().upper()
    existing = get_patient(db_path, p_no)
    if not existing:
        return None

    is_ihc = (existing.get("is_needs_ihc") == 1)
    is_recut = (existing.get("is_needs_recut") == 1)
    is_pending = 1
    new_status = build_review_status(True, is_ihc, is_recut, fallback="待审核")

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE patients SET
                review_status = ?,
                is_pending_review = ?,
                updated_at = ?
            WHERE pathology_no = ?
        """, (new_status, is_pending, now_str, p_no))
        conn.commit()
    return get_patient(db_path, p_no)

def get_patients_for_signing_sheet(db_path, start_date=None, end_date=None, prefixes=None,
                                   report_start_date=None, report_end_date=None):
    """
    Retrieve patients for signing sheet based on submission/report date ranges and prefixes.
    Does not depend on archive or review status.
    """
    conditions = []
    values = []

    _append_date_filter(conditions, values, SUB_DATE_EXPR, start_date=start_date, end_date=end_date)
    _append_date_filter(conditions, values, REP_DATE_EXPR, start_date=report_start_date, end_date=report_end_date)
    _append_prefix_filter(conditions, values, prefixes=prefixes)

    sql = "SELECT * FROM patients"
    if conditions:
        sql += " WHERE " + " AND ".join(conditions)
    sql += " ORDER BY pathology_no ASC"

    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute(sql, values)
        return [patient_row_to_dict(r) for r in cursor.fetchall()]


def get_today_batch_archivable_patients(db_path, today_str=None):
    """
    Retrieve patients with today's report date that have a saved report not yet archived
    (includes regular reports, IHC-only reports, and approved reports; excludes pending review / recut).
    """
    if not today_str:
        today_str = datetime.now().strftime("%Y-%m-%d")

    conditions = [
        "archived IN (0, 2)",
        "final_report_path IS NOT NULL",
        "TRIM(final_report_path) != ''",
        f"({REP_DATE_EXPR} = ?)",
        "("
        "review_status = '完成审核' "
        "OR ("
        "is_pending_review = 0 AND is_needs_recut = 0 "
        "AND (review_status IS NULL OR (review_status NOT LIKE '%待审核%' AND review_status NOT LIKE '%重切%'))"
        ")"
        ")"
    ]
    sql = "SELECT * FROM patients WHERE " + " AND ".join(conditions) + " ORDER BY pathology_no ASC"

    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute(sql, [today_str])
        return [patient_row_to_dict(r) for r in cursor.fetchall()]
