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

def get_today_patients(db_path, today_raw=None):
    """Retrieve patients created or reported today."""
    if not today_raw:
        today_raw = datetime.now().strftime("%Y-%m-%d")

    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT * FROM patients
            WHERE report_date_raw = ? OR created_at LIKE ?
            ORDER BY created_at DESC, pathology_no DESC
        """, (today_raw, f"{today_raw}%"))
        rows = cursor.fetchall()
        return [patient_row_to_dict(r) for r in rows]

def get_patients_for_report_making(db_path, date_raw=None, keyword=None):
    """
    Retrieve patients for Report Making workspace.
    Filterable by date and search keyword.
    """
    conditions = []
    values = []

    if date_raw:
        conditions.append("(report_date_raw = ? OR submission_date_raw = ? OR created_at LIKE ?)")
        values.extend([date_raw, date_raw, f"{date_raw}%"])

    if keyword:
        kw_like = f"%{keyword.strip()}%"
        conditions.append("(name LIKE ? OR pathology_no LIKE ? OR inpatient_no LIKE ? OR outpatient_no LIKE ? OR clinical_diagnosis LIKE ? OR pathology_diag LIKE ?)")
        values.extend([kw_like, kw_like, kw_like, kw_like, kw_like, kw_like])

    sql = "SELECT * FROM patients"
    if conditions:
        sql += " WHERE " + " AND ".join(conditions)
    sql += " ORDER BY pathology_no ASC LIMIT 200"

    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute(sql, values)
        rows = cursor.fetchall()
        return [patient_row_to_dict(r) for r in rows]

def get_patients_for_review(db_path, include_approved=False):
    """
    Retrieve patients for Report Review workspace.
    Includes pending review and IHC patients.
    """
    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        if include_approved:
            cursor.execute("""
                SELECT * FROM patients
                WHERE is_pending_review = 1 OR is_needs_ihc = 1 OR is_needs_recut = 1
                   OR review_status LIKE '%待审核%' OR review_status LIKE '%免疫组化%' OR review_status LIKE '%重切%'
                   OR review_status = '完成审核'
                ORDER BY updated_at DESC, pathology_no DESC
            """)
        else:
            cursor.execute("""
                SELECT * FROM patients
                WHERE (is_pending_review = 1 OR is_needs_ihc = 1 OR is_needs_recut = 1
                   OR review_status LIKE '%待审核%' OR review_status LIKE '%免疫组化%' OR review_status LIKE '%重切%')
                  AND review_status != '完成审核'
                ORDER BY updated_at DESC, pathology_no DESC
            """)
        rows = cursor.fetchall()
        return [patient_row_to_dict(r) for r in rows]

def search_patients(db_path, query_params):
    """Search patients across all criteria including dates, prefix, archived and review status."""
    conditions = []
    values = []

    # 1. Date range filter (by submission date)
    start_date = query_params.get("start_date", "").strip()
    end_date = query_params.get("end_date", "").strip()
    date_field = "COALESCE(NULLIF(submission_date_raw, ''), NULLIF(report_date_raw, ''), substr(created_at, 1, 10))"
    if start_date and end_date:
        conditions.append(f"({date_field} BETWEEN ? AND ?)")
        values.extend([start_date, end_date])
    elif start_date:
        conditions.append(f"({date_field} >= ?)")
        values.append(start_date)
    elif end_date:
        conditions.append(f"({date_field} <= ?)")
        values.append(end_date)

    # 2. Prefix filter
    prefix = query_params.get("prefix", "").strip().upper()
    if prefix and prefix not in ("ALL", "所有"):
        if prefix in ("NONE", "普通", "无前缀"):
            conditions.append("(prefix = '' OR prefix IS NULL)")
        else:
            conditions.append("UPPER(prefix) = ?")
            values.append(prefix)

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
            conditions.append("(review_status LIKE '%免疫组化%')")
        elif rev_status == "重切":
            conditions.append("(review_status LIKE '%重切%')")
        elif rev_status == "完成审核":
            conditions.append("review_status = '完成审核'")
        else:
            conditions.append("review_status = ?")
            values.append(rev_status)

    keyword = query_params.get("keyword", "").strip()
    if keyword:
        kw_like = f"%{keyword}%"
        conditions.append("(name LIKE ? OR pathology_no LIKE ? OR inpatient_no LIKE ? OR outpatient_no LIKE ? OR clinical_diagnosis LIKE ? OR biopsy_sites LIKE ? OR referring_doctor LIKE ? OR pathology_diag LIKE ?)")
        values.extend([kw_like, kw_like, kw_like, kw_like, kw_like, kw_like, kw_like, kw_like])

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

def get_output_regular_patients(db_path, prefix=None, today_only=False, today_str=None):
    """
    Retrieve archived regular patients (List 1 in Report Output).
    archived in (1, 2) AND not pending review / not IHC.
    """
    if not today_str:
        today_str = datetime.now().strftime("%Y-%m-%d")

    conditions = [
        "archived IN (1, 2)",
        "is_pending_review = 0",
        "is_needs_ihc = 0",
        "is_needs_recut = 0",
        "(review_status IS NULL OR review_status IN ('', '无', '未审核', '已制作'))"
    ]
    values = []

    if prefix is not None and prefix != "ALL" and prefix != "所有":
        p_clean = str(prefix).strip().upper()
        if p_clean in ("", "NONE", "普通"):
            conditions.append("(prefix = '' OR prefix IS NULL)")
        else:
            conditions.append("prefix = ?")
            values.append(p_clean)

    if today_only:
        conditions.append("(report_date_raw = ? OR submission_date_raw = ? OR created_at LIKE ?)")
        values.extend([today_str, today_str, f"{today_str}%"])

    sql = "SELECT * FROM patients WHERE " + " AND ".join(conditions) + " ORDER BY pathology_no DESC"
    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute(sql, values)
        return [patient_row_to_dict(r) for r in cursor.fetchall()]

def get_output_special_approved_patients(db_path, prefix=None, today_only=False, today_str=None):
    """
    Retrieve archived approved special patients (List 2 in Report Output).
    archived in (1, 2) AND review_status == '完成审核'.
    """
    if not today_str:
        today_str = datetime.now().strftime("%Y-%m-%d")

    conditions = [
        "archived IN (1, 2)",
        "review_status = '完成审核'"
    ]
    values = []

    if prefix is not None and prefix != "ALL" and prefix != "所有":
        p_clean = str(prefix).strip().upper()
        if p_clean in ("", "NONE", "普通"):
            conditions.append("(prefix = '' OR prefix IS NULL)")
        else:
            conditions.append("prefix = ?")
            values.append(p_clean)

    if today_only:
        conditions.append("(report_date_raw = ? OR submission_date_raw = ? OR created_at LIKE ?)")
        values.extend([today_str, today_str, f"{today_str}%"])

    sql = "SELECT * FROM patients WHERE " + " AND ".join(conditions) + " ORDER BY pathology_no DESC"
    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute(sql, values)
        return [patient_row_to_dict(r) for r in cursor.fetchall()]

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

def get_patients_for_signing_sheet(db_path, start_date=None, end_date=None, prefixes=None):
    """
    Retrieve patients for signing sheet based on date range and prefixes.
    Does not depend on archive or review status.
    """
    conditions = []
    values = []

    if start_date and end_date:
        conditions.append("(COALESCE(NULLIF(submission_date_raw, ''), NULLIF(report_date_raw, ''), substr(created_at, 1, 10)) BETWEEN ? AND ?)")
        values.extend([start_date, end_date])
    elif start_date:
        conditions.append("(COALESCE(NULLIF(submission_date_raw, ''), NULLIF(report_date_raw, ''), substr(created_at, 1, 10)) >= ?)")
        values.append(start_date)
    elif end_date:
        conditions.append("(COALESCE(NULLIF(submission_date_raw, ''), NULLIF(report_date_raw, ''), substr(created_at, 1, 10)) <= ?)")
        values.append(end_date)

    if prefixes is not None:
        if "ALL" not in prefixes and "所有" not in prefixes:
            prefix_conds = []
            for pfx in prefixes:
                pfx_clean = str(pfx).strip().upper()
                if pfx_clean in ("", "NONE", "普通"):
                    prefix_conds.append("(prefix = '' OR prefix IS NULL)")
                else:
                    prefix_conds.append("prefix = ?")
                    values.append(pfx_clean)
            if prefix_conds:
                conditions.append("(" + " OR ".join(prefix_conds) + ")")

    sql = "SELECT * FROM patients"
    if conditions:
        sql += " WHERE " + " AND ".join(conditions)
    sql += " ORDER BY pathology_no ASC"

    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute(sql, values)
        return [patient_row_to_dict(r) for r in cursor.fetchall()]
