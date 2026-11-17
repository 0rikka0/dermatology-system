# -*- coding: utf-8 -*-
"""
Dermatology Pathology Reporting System - Path & Directory Management
"""
import os
import re
import shutil
import subprocess

SUBDIRECTORIES = [
    "已归档报告",
    "数据库",
    os.path.join("特殊患者归档", "待审核"),
    os.path.join("特殊患者归档", "免疫组化"),
    os.path.join("特殊患者归档", "重切"),
    os.path.join("特殊患者归档", "完成审核"),
    os.path.join("未归档报告", "临床照片"),
    os.path.join("未归档报告", "病理报告录入"),
    os.path.join("未归档报告", "病理报告输出"),
    os.path.join("未归档报告", "病理照片"),
    "领取表"
]

def safe_delete_path(target_path):
    """
    Safely delete a file or directory using Windows Recycle Bin (soft delete).
    Falls back to permanent delete only if shell operations fail.
    """
    if not target_path or not os.path.exists(target_path):
        return {"success": True}
    norm = os.path.normpath(target_path)
    try:
        from send2trash import send2trash
        send2trash(norm)
        return {"success": not os.path.exists(norm)}
    except Exception:
        pass

    try:
        norm_esc = norm.replace("'", "''")
        if os.path.isdir(norm):
            cmd = f'powershell -NoProfile -Command "Add-Type -AssemblyName Microsoft.VisualBasic; [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory(\'{norm_esc}\', \'OnlyErrorDialogs\', \'SendToRecycleBin\')"'
        else:
            cmd = f'powershell -NoProfile -Command "Add-Type -AssemblyName Microsoft.VisualBasic; [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(\'{norm_esc}\', \'OnlyErrorDialogs\', \'SendToRecycleBin\')"'
        res = subprocess.run(cmd, shell=True, capture_output=True, timeout=5)
        if res.returncode == 0 and not os.path.exists(norm):
            return {"success": True}
    except Exception:
        pass

    # Fallback to permanent delete if recycling fails
    if os.path.isdir(norm):
        shutil.rmtree(norm, ignore_errors=True)
    else:
        try:
            os.remove(norm)
        except Exception:
            pass
    return {"success": not os.path.exists(norm)}

def ensure_workspace_structure(work_dir):
    """
    Ensure the standard directory structure exists inside the working directory.
    """
    if not os.path.exists(work_dir):
        os.makedirs(work_dir, exist_ok=True)

    created = []
    for sub in SUBDIRECTORIES:
        full_path = os.path.join(work_dir, sub)
        if not os.path.exists(full_path):
            os.makedirs(full_path, exist_ok=True)
            created.append(sub)

    return {
        "success": True,
        "work_dir": work_dir,
        "created_dirs": created
    }

def get_database_path(work_dir):
    """Returns absolute path to SQLite database."""
    return os.path.join(work_dir, "数据库", "pathology.db")

def get_report_dir(work_dir):
    """Directory for draft word reports (病理报告录入)."""
    d = os.path.join(work_dir, "未归档报告", "病理报告录入")
    if not os.path.exists(d):
        legacy = os.path.join(work_dir, "统一归档", "病理报告录入")
        if os.path.exists(legacy):
            return legacy
    os.makedirs(d, exist_ok=True)
    return d

def get_output_report_dir(work_dir):
    """Directory for finalized word reports (病理报告输出)."""
    d = os.path.join(work_dir, "未归档报告", "病理报告输出")
    if not os.path.exists(d):
        legacy = os.path.join(work_dir, "统一归档", "病理报告输出")
        if os.path.exists(legacy):
            return legacy
    os.makedirs(d, exist_ok=True)
    return d

def get_photo_root_dir(work_dir):
    """Root directory for pathology photos (病理照片)."""
    d = os.path.join(work_dir, "未归档报告", "病理照片")
    if not os.path.exists(d):
        legacy = os.path.join(work_dir, "统一归档", "病理照片")
        if os.path.exists(legacy):
            return legacy
    os.makedirs(d, exist_ok=True)
    return d

def get_clinical_photo_root_dir(work_dir):
    """Root directory for clinical photos (临床照片)."""
    d = os.path.join(work_dir, "未归档报告", "临床照片")
    if not os.path.exists(d):
        legacy = os.path.join(work_dir, "统一归档", "临床照片")
        if os.path.exists(legacy):
            return legacy
    os.makedirs(d, exist_ok=True)
    return d

def get_archive_root_dir(work_dir):
    """Root directory for finalized archives (已归档报告)."""
    d = os.path.join(work_dir, "已归档报告")
    os.makedirs(d, exist_ok=True)
    return d

IHC_ARCHIVE_SUFFIX = "_免疫组化报告.docx"


def is_ihc_archived_filename(filename):
    """Returns True if filename is an archived IHC report docx."""
    return str(filename or "").endswith(IHC_ARCHIVE_SUFFIX)


def sanitize_filename(name):
    """Remove characters forbidden in Windows filenames except valid dots/semicolons."""
    # Forbidden: \ / : * ? " < > | '
    s = re.sub(r'[\\/:*?"<>|\']', '_', str(name).strip())
    return s.strip()

def get_patient_word_filename(pathology_no, name):
    """Draft filename: 病理号_姓名.docx"""
    safe_no = sanitize_filename(pathology_no)
    safe_name = sanitize_filename(name)
    return f"{safe_no}_{safe_name}.docx"

def get_patient_folder_name(pathology_no, name):
    """Patient folder name: 病理号_姓名"""
    safe_no = sanitize_filename(pathology_no)
    safe_name = sanitize_filename(name)
    return f"{safe_no}_{safe_name}"

def get_final_report_filename(pathology_no, name, diagnosis=""):
    """
    Final output report filename: 病理号_姓名_病理诊断.docx
    Matches hospital format: e.g. 2026102105_张建国_1.2.交界痣；3.皮内痣.docx
    """
    safe_no = sanitize_filename(pathology_no)
    safe_name = sanitize_filename(name)
    diag = sanitize_filename(diagnosis) if diagnosis else "待诊断"
    return f"{safe_no}_{safe_name}_{diag}.docx"

def get_archived_report_filename(pathology_no, name, diagnosis=""):
    """
    Standard archived report filename in 已归档报告/:
    病理号_姓名_病理诊断.docx (No '[病理报告]' prefix!)
    """
    safe_no = sanitize_filename(pathology_no)
    safe_name = sanitize_filename(name)
    diag = sanitize_filename(diagnosis) if diagnosis else "已诊断"
    return f"{safe_no}_{safe_name}_{diag}.docx"

def get_patient_word_path(work_dir, pathology_no, name):
    """Path to draft word report in 病理报告录入."""
    return os.path.join(get_report_dir(work_dir), get_patient_word_filename(pathology_no, name))

def get_final_report_path(work_dir, pathology_no, name, diagnosis=""):
    """Path to final word report in 病理报告输出."""
    return os.path.join(get_output_report_dir(work_dir), get_final_report_filename(pathology_no, name, diagnosis))

def get_archived_report_path(work_dir, pathology_no, name, diagnosis=""):
    """Path to flat archived docx in 已归档报告/."""
    return os.path.join(get_archive_root_dir(work_dir), get_archived_report_filename(pathology_no, name, diagnosis))

def get_patient_photo_dir(work_dir, pathology_no, name):
    """Path to patient's photo directory in 病理照片."""
    return os.path.join(get_photo_root_dir(work_dir), get_patient_folder_name(pathology_no, name))

def ensure_patient_photo_dir(work_dir, pathology_no, name):
    """Creates the patient's photo folder if it does not exist."""
    path = get_patient_photo_dir(work_dir, pathology_no, name)
    os.makedirs(path, exist_ok=True)
    return path

def find_clinical_photo_folder(work_dir, patient_name, pathology_no=""):
    """
    Looks for a folder matching patient_name in 未归档报告/临床照片/.
    Folder names in 未归档报告/临床照片/ are patient names.
    Rules:
    - If exactly 1 matching folder found: found=True, multiple=False, path=folder_path.
    - If multiple matching folders found (重名/多个匹配): found=False, multiple=True, matches=[...], path=root.
    - If 0 matching folders found: found=False, multiple=False, matches=[], path=root.
    """
    root = get_clinical_photo_root_dir(work_dir)
    if not os.path.exists(root):
        os.makedirs(root, exist_ok=True)
        return {"found": False, "multiple": False, "matches": [], "path": root, "root": root}

    clean_name = str(patient_name).strip()
    clean_no = str(pathology_no).strip()

    if not clean_name:
        return {"found": False, "multiple": False, "matches": [], "path": root, "root": root}

    try:
        entries = [e for e in os.listdir(root) if os.path.isdir(os.path.join(root, e))]
    except Exception:
        entries = []

    # 1. Exact match candidates:
    exact_matches = [e for e in entries if e == clean_name or e == f"{clean_no}_{clean_name}" or e == f"{clean_name}_{clean_no}"]

    # 2. Substring match candidates:
    fuzzy_matches = [e for e in entries if clean_name in e]

    # Combine candidates
    candidate_set = set(exact_matches)
    for fm in fuzzy_matches:
        candidate_set.add(fm)
    all_candidates = sorted(list(candidate_set))

    if len(all_candidates) == 1:
        match_folder = all_candidates[0]
        full_p = os.path.join(root, match_folder)
        return {
            "found": True,
            "multiple": False,
            "matches": all_candidates,
            "path": full_p,
            "folder_name": match_folder,
            "root": root
        }
    elif len(all_candidates) > 1:
        # Multiple matches (重名/多个匹配项)!
        return {
            "found": False,
            "multiple": True,
            "matches": all_candidates,
            "path": root,
            "root": root,
            "message": f"检测到存在多个匹配的临床照片文件夹 ({', '.join(all_candidates)})，请手动确认！"
        }
    else:
        return {
            "found": False,
            "multiple": False,
            "matches": [],
            "path": root,
            "root": root,
            "message": f"未找到患者 [{clean_name}] 专属临床照片文件夹"
        }

def rename_patient_files_and_folders(work_dir, old_no, old_name, new_no, new_name):
    """
    Renames Word report file, output report file, photo folder, and special archive folders
    if pathology number or patient name changed.
    Does NOT touch 已归档报告/ (archived reports are only updated on explicit re-archive).
    """
    old_word = get_patient_word_path(work_dir, old_no, old_name)
    new_word = get_patient_word_path(work_dir, new_no, new_name)

    old_photo = get_patient_photo_dir(work_dir, old_no, old_name)
    new_photo = get_patient_photo_dir(work_dir, new_no, new_name)

    old_prefix_base = f"{sanitize_filename(old_no)}_{sanitize_filename(old_name)}"
    new_prefix_base = f"{sanitize_filename(new_no)}_{sanitize_filename(new_name)}"
    old_prefix = f"{old_prefix_base}_"
    new_prefix = f"{new_prefix_base}_"

    # 1. Rename draft Word file
    if old_word != new_word and os.path.exists(old_word):
        if not os.path.exists(new_word):
            shutil.move(old_word, new_word)
        else:
            try:
                safe_delete_path(old_word)
            except Exception:
                pass

    # 2. Rename final output report file if exists
    new_final_path = ""
    out_dir = get_output_report_dir(work_dir)
    if os.path.exists(out_dir):
        for f in os.listdir(out_dir):
            if f.startswith(old_prefix) and f.endswith(".docx"):
                old_f = os.path.join(out_dir, f)
                diag_part = f[len(old_prefix):]
                new_f = os.path.join(out_dir, f"{new_prefix}{diag_part}")
                if old_f != new_f:
                    try:
                        shutil.move(old_f, new_f)
                    except Exception:
                        pass
                new_final_path = new_f if os.path.exists(new_f) else old_f
                break

    # 3. Rename special archive folders & internal report docx (excluding 已归档报告/)
    renamed_ihc = {}
    if old_prefix_base != new_prefix_base:
        for sub in ["待审核", "免疫组化", "重切", "完成审核"]:
            spec_root = os.path.join(work_dir, "特殊患者归档", sub)
            if not os.path.exists(spec_root):
                continue
            for entry in list(os.listdir(spec_root)):
                if entry.startswith(old_prefix) or entry == old_prefix_base:
                    old_dir = os.path.join(spec_root, entry)
                    if not os.path.isdir(old_dir):
                        continue
                    suffix = entry[len(old_prefix_base):]
                    new_entry = f"{new_prefix_base}{suffix}"
                    new_dir = os.path.join(spec_root, new_entry)
                    if old_dir != new_dir and not os.path.exists(new_dir):
                        try:
                            shutil.move(old_dir, new_dir)
                        except Exception:
                            new_dir = old_dir
                    if os.path.isdir(new_dir):
                        old_doc = os.path.join(new_dir, f"[病理报告]{entry}.docx")
                        new_doc = os.path.join(new_dir, f"[病理报告]{new_entry}.docx")
                        if os.path.exists(old_doc) and old_doc != new_doc:
                            try:
                                shutil.move(old_doc, new_doc)
                            except Exception:
                                pass
                        if sub == "免疫组化":
                            renamed_ihc = {
                                "old_dir": old_dir,
                                "ihc_dir": new_dir,
                                "ihc_doc_path": new_doc if os.path.exists(new_doc) else old_doc,
                                "ihc_photo_dir": os.path.join(new_dir, "病理照片")
                            }

    # 4. Rename pathology photo folder
    if old_photo != new_photo and os.path.exists(old_photo):
        if not os.path.exists(new_photo):
            shutil.move(old_photo, new_photo)
        else:
            for item in os.listdir(old_photo):
                s = os.path.join(old_photo, item)
                d = os.path.join(new_photo, item)
                if not os.path.exists(d):
                    shutil.move(s, d)
            try:
                os.rmdir(old_photo)
            except Exception:
                pass

    ensure_patient_photo_dir(work_dir, new_no, new_name)
    return {
        "word_path": new_word,
        "photo_dir": new_photo,
        "old_photo_dir": old_photo,
        "final_report_path": new_final_path,
        "renamed_ihc": renamed_ihc
    }


def _ensure_writable(file_path):
    """Verify target file is not locked by Word/WPS before overwriting."""
    if file_path and os.path.exists(file_path) and os.path.isfile(file_path):
        try:
            with open(file_path, "a+b"):
                pass
        except (PermissionError, IOError) as pe:
            raise PermissionError(
                f"报告文件被其他软件（如 Word 或 WPS）打开占用，无法覆盖！请先关闭该文件后重试。路径: {os.path.basename(file_path)}"
            ) from pe


def sync_photo_files(src_dir, dst_dir):
    """Copy missing files from src_dir to dst_dir without overwriting existing files."""
    if not src_dir or not os.path.exists(src_dir) or not os.path.isdir(src_dir):
        return []
    os.makedirs(dst_dir, exist_ok=True)
    copied = []
    for item in os.listdir(src_dir):
        s = os.path.join(src_dir, item)
        d = os.path.join(dst_dir, item)
        if os.path.isfile(s) and not os.path.exists(d):
            shutil.copy2(s, d)
            copied.append(d)
    return copied


def cleanup_stale_output_reports(work_dir, pathology_no, patient_name, keep_path=""):
    """Remove obsolete diagnosis filenames in 未归档报告/病理报告输出/ for this patient."""
    out_dir = get_output_report_dir(work_dir)
    if not os.path.exists(out_dir):
        return []
    p_no_clean = sanitize_filename(str(pathology_no or "").strip())
    p_name_clean = sanitize_filename(str(patient_name or "").strip())
    prefix_base = f"{p_no_clean}_{p_name_clean}"
    keep_norm = os.path.normcase(os.path.normpath(keep_path)) if keep_path else ""
    removed = []
    for f in os.listdir(out_dir):
        if (f.startswith(f"{prefix_base}_") or f.startswith(f"{prefix_base}.")) and f.lower().endswith(".docx"):
            fp = os.path.join(out_dir, f)
            if keep_norm and os.path.normcase(os.path.normpath(fp)) == keep_norm:
                continue
            safe_delete_path(fp)
            removed.append(fp)
    return removed


def _find_special_cat_folder(work_dir, cat_name, pathology_no, patient_name):
    """Find an existing subfolder for patient under 工作目录/特殊患者归档/<cat_name>/."""
    p_no_clean = sanitize_filename(str(pathology_no or "").strip())
    p_name_clean = sanitize_filename(str(patient_name or "").strip())
    prefix_base = f"{p_no_clean}_{p_name_clean}"
    cat_root = os.path.join(work_dir, "特殊患者归档", cat_name)
    if os.path.exists(cat_root):
        try:
            for entry in os.listdir(cat_root):
                full = os.path.join(cat_root, entry)
                if os.path.isdir(full) and (entry.startswith(f"{prefix_base}_") or entry == prefix_base):
                    return full
        except Exception:
            pass
    return ""


def find_patient_review_folder(work_dir, pathology_no, patient_name, review_status="", is_pending=None, is_recut=None):
    """
    Resolve the patient's copied special review directory and its 病理照片 subfolder
    under 工作目录/特殊患者归档/(待审核 | 重切 | 完成审核).
    """
    rev_st = str(review_status or "")
    has_pending = bool(is_pending) or ("待审核" in rev_st)
    has_recut = bool(is_recut) or ("重切" in rev_st)

    if rev_st == "完成审核":
        order = ["完成审核", "待审核", "重切"]
    elif has_recut and not has_pending:
        order = ["重切", "待审核", "完成审核"]
    else:
        order = ["待审核", "重切", "完成审核"]

    for cat in order:
        fdir = _find_special_cat_folder(work_dir, cat, pathology_no, patient_name)
        if fdir:
            return {
                "category": cat,
                "folder_dir": fdir,
                "photo_dir": os.path.join(fdir, "病理照片")
            }
    return {}


def sync_review_photos_to_making_and_siblings(work_dir, patient_data, primary_photo_dir=""):
    """
    Ensure any newly added pathology photos in copied review directories (待审核/重切/完成审核)
    are copied into 未归档报告/病理照片/ and sibling review directories so Maker & Review never lose images.
    """
    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    making_photo_dir = get_patient_photo_dir(work_dir, p_no, p_name)

    review_photo_dirs = []
    if primary_photo_dir and os.path.isdir(primary_photo_dir):
        review_photo_dirs.append(primary_photo_dir)
    for cat in ("待审核", "重切", "完成审核"):
        fdir = _find_special_cat_folder(work_dir, cat, p_no, p_name)
        if fdir:
            pdir = os.path.join(fdir, "病理照片")
            if os.path.isdir(pdir) and pdir not in review_photo_dirs:
                review_photo_dirs.append(pdir)

    copied = []
    # Copy any additions from copied review directories -> 未归档报告/病理照片
    for src_dir in review_photo_dirs:
        copied.extend(sync_photo_files(src_dir, making_photo_dir))
    # Propagate across sibling review photo directories (e.g., if both 待审核 and 重切 exist)
    if len(review_photo_dirs) > 1:
        for src_dir in review_photo_dirs:
            for dst_dir in review_photo_dirs:
                if src_dir != dst_dir:
                    copied.extend(sync_photo_files(src_dir, dst_dir))
    return copied


def _sync_folder_to_special_category(work_dir, cat_name, patient_data, report_file_path):
    """
    Syncs the patient's report docx and pathology photos (NO clinical photos) into
    工作目录/特殊患者归档/<cat_name>/<no>_<name>_<diag>/ cleanly, migrating photos and removing
    any stale old-diagnosis folder for the same patient inside that category.
    """
    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    p_diag = patient_data.get("pathology_diag") or patient_data.get("clinical_diagnosis") or "已诊断"
    p_no_clean = sanitize_filename(p_no)
    p_name_clean = sanitize_filename(p_name)
    safe_folder_name = f"{p_no_clean}_{p_name_clean}_{sanitize_filename(p_diag)}"
    prefix_base = f"{p_no_clean}_{p_name_clean}"
    prefix = f"{prefix_base}_"

    cat_root = os.path.join(work_dir, "特殊患者归档", cat_name)
    os.makedirs(cat_root, exist_ok=True)
    target_dir = os.path.join(cat_root, safe_folder_name)
    photo_target = os.path.join(target_dir, "病理照片")
    target_doc_name = f"[病理报告]{safe_folder_name}.docx"
    target_doc_path = os.path.join(target_dir, target_doc_name)

    _ensure_writable(target_doc_path)

    os.makedirs(target_dir, exist_ok=True)
    os.makedirs(photo_target, exist_ok=True)

    # Migrate photos and remove any stale old-diagnosis folder in this category
    for entry in list(os.listdir(cat_root)):
        if entry != safe_folder_name and (entry.startswith(prefix) or entry == prefix_base):
            old_dir = os.path.join(cat_root, entry)
            if os.path.isdir(old_dir):
                old_doc = os.path.join(old_dir, f"[病理报告]{entry}.docx")
                _ensure_writable(old_doc)
                sync_photo_files(os.path.join(old_dir, "病理照片"), photo_target)
                safe_delete_path(old_dir)

    # Sync pathology photos from 未归档报告/病理照片/ (NO clinical photos copied!)
    src_photo_dir = get_patient_photo_dir(work_dir, p_no, p_name)
    sync_photo_files(src_photo_dir, photo_target)

    # Remove any differently-named [病理报告]*.docx inside target_dir
    for fn in list(os.listdir(target_dir)):
        if fn.lower().endswith(".docx") and fn != target_doc_name:
            old_inner_doc = os.path.join(target_dir, fn)
            _ensure_writable(old_inner_doc)
            safe_delete_path(old_inner_doc)

    if report_file_path and os.path.exists(report_file_path):
        shutil.copy2(report_file_path, target_doc_path)

    return target_dir, target_doc_path


def sync_updated_report_to_special_archives(work_dir, patient_data, report_file_path):
    """
    Overwrites report docx and syncs folder names inside any existing 特殊患者归档/ review subfolders
    (待审核, 重切, 完成审核) for this patient, without touching 已归档报告/ or 特殊患者归档/免疫组化/.
    """
    if not report_file_path or not os.path.exists(report_file_path):
        return []
    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    updated = []
    for sub in ["待审核", "重切", "完成审核"]:
        existing_dir = _find_special_cat_folder(work_dir, sub, p_no, p_name)
        if existing_dir:
            _, dst_doc = _sync_folder_to_special_category(work_dir, sub, patient_data, report_file_path)
            updated.append(dst_doc)
    return updated

def save_special_draft_queue(work_dir, patient_data, report_file_path):
    """
    When saving a report marked as 待审核 or 重切:
    Packages the patient's pathology photos (NO clinical photos) and report into
    '特殊患者归档/待审核' and/or '特殊患者归档/重切'.
    If a tag was unchecked, removes it from that category.
    Note: 免疫组化 is managed independently in the IHC workspace and is NOT auto-created here.
    """
    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    prefix_base = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}"
    prefix = f"{prefix_base}_"

    is_pending = int(patient_data.get("is_pending_review", 0)) == 1 or "待审核" in str(patient_data.get("review_status", ""))
    is_recut = int(patient_data.get("is_needs_recut", 0)) == 1 or "重切" in str(patient_data.get("review_status", ""))

    saved_dirs = []
    for cat_name, enabled in [("待审核", is_pending), ("重切", is_recut)]:
        cat_root = os.path.join(work_dir, "特殊患者归档", cat_name)
        if enabled:
            t_dir, _ = _sync_folder_to_special_category(work_dir, cat_name, patient_data, report_file_path)
            saved_dirs.append(t_dir)
        else:
            if os.path.exists(cat_root):
                for e in list(os.listdir(cat_root)):
                    if e.startswith(prefix) or e == prefix_base:
                        safe_delete_path(os.path.join(cat_root, e))

    return {
        "success": True,
        "saved_dirs": saved_dirs,
        "duplicate_clinical_photos": False,
        "duplicate_matches": [],
        "clinical_root": get_clinical_photo_root_dir(work_dir),
        "target_dirs": []
    }

def archive_patient_report(work_dir, patient_data, report_file_path, is_from_review=False):
    """
    Archives a patient's original report:
      1. Regular archives (including IHC-only patients):
         Stored directly into: 工作目录/已归档报告/病理号_姓名_诊断.docx (Flat file, NO subfolders, NO '[病理报告]' prefix!)
         Photos are NOT copied to regular archive (solves file duplication / bloat).
         Does NOT overwrite or delete *_免疫组化报告.docx or 特殊患者归档/免疫组化/.
      2. Review special patients archives (待审核 / 重切):
         Moves package (pathology photos + [病理报告]...docx, NO clinical photos) from 待审核 / 重切 into 特殊患者归档/完成审核/.
    """
    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    p_diag = patient_data.get("pathology_diag") or patient_data.get("clinical_diagnosis") or "已诊断"
    safe_folder_name = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}_{sanitize_filename(p_diag)}"
    prefix = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}_"
    prefix_base = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}"

    rev_st = str(patient_data.get("review_status", ""))
    is_still_pending = (rev_st != "完成审核") and (
        any(k in rev_st for k in ("待审核", "重切"))
        or int(patient_data.get("is_pending_review", 0)) == 1
        or int(patient_data.get("is_needs_recut", 0)) == 1
    )

    if not is_from_review and is_still_pending:
        return {
            "success": False,
            "error": "标记为“待审核”或“重切”的报告禁止在报告制作工作区直接归档！已在保存时自动放入特殊患者待审队列，请移步【报告审核】工作区完成专家审核后再进行归档。"
        }

    # 1. Main archive in '已归档报告' (Flat docx report file, NO subfolders, NO '[病理报告]' prefix!)
    archive_root = get_archive_root_dir(work_dir)
    target_report_name = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}_{sanitize_filename(p_diag)}.docx"
    target_archive_file = os.path.join(archive_root, target_report_name)
    _ensure_writable(target_archive_file)

    # Clean up any legacy or duplicate non-IHC archived files for this patient in 已归档报告/
    if os.path.exists(archive_root):
        for item in list(os.listdir(archive_root)):
            if (
                (item.startswith(prefix) or item.startswith(f"{prefix_base}."))
                and item.endswith(".docx")
                and not is_ihc_archived_filename(item)
            ):
                old_f = os.path.join(archive_root, item)
                if old_f != target_archive_file:
                    safe_delete_path(old_f)

    # Copy final report to 已归档报告/
    if report_file_path and os.path.exists(report_file_path):
        shutil.copy2(report_file_path, target_archive_file)

    # 2. Special archive handling (only for 待审核 / 重切 / 完成审核, never 免疫组化, NO clinical photos):
    special_paths = []
    is_approved_special = (patient_data.get("review_status") == "完成审核") or is_from_review
    is_special_patient = (
        int(patient_data.get("is_pending_review", 0)) == 1 or
        int(patient_data.get("is_needs_recut", 0)) == 1 or
        any(k in str(patient_data.get("review_status", "")) for k in ("待审核", "重切", "完成审核"))
    )

    if is_special_patient and is_approved_special:
        # Preserve any photos added in 待审核/重切 before removing them
        making_photo_dir = get_patient_photo_dir(work_dir, p_no, p_name)
        for sub in ["待审核", "重切"]:
            sub_dir = _find_special_cat_folder(work_dir, sub, p_no, p_name)
            if sub_dir:
                sync_photo_files(os.path.join(sub_dir, "病理照片"), making_photo_dir)

        approved_dir, _ = _sync_folder_to_special_category(work_dir, "完成审核", patient_data, report_file_path)

        # Clean remaining in 待审核 and 重切 (never touch 免疫组化)
        for sub in ["待审核", "重切"]:
            sub_root = os.path.join(work_dir, "特殊患者归档", sub)
            if os.path.exists(sub_root):
                for e in list(os.listdir(sub_root)):
                    if e.startswith(prefix) or e == prefix_base:
                        safe_delete_path(os.path.join(sub_root, e))

        special_paths.append(approved_dir)

    return {
        "success": True,
        "archive_dir": target_archive_file,
        "archive_file": target_archive_file,
        "special_paths": special_paths,
        "duplicate_clinical_photos": False,
        "duplicate_matches": [],
        "clinical_root": get_clinical_photo_root_dir(work_dir),
        "target_dirs": []
    }

def sync_review_overwrites(work_dir, patient_data, new_report_path):
    """
    Synchronizes updated draft docx across existing 特殊患者归档/<cat>/<patient_folder>
    (待审核, 重切, 完成审核; excluding 免疫组化).
    Does NOT overwrite 已归档报告/ directly — re-archiving is triggered explicitly in Report Making.
    """
    if not new_report_path or not os.path.exists(new_report_path):
        return {"success": False, "error": "Report file does not exist"}

    overwritten = sync_updated_report_to_special_archives(work_dir, patient_data, new_report_path)
    return {"success": True, "overwritten": overwritten}

def find_patient_ihc_folder(work_dir, pathology_no, patient_name):
    """Finds existing folder under 特殊患者归档/免疫组化/ for this patient."""
    p_no_clean = sanitize_filename(str(pathology_no).strip())
    p_name_clean = sanitize_filename(str(patient_name).strip())
    prefix = f"{p_no_clean}_{p_name_clean}"
    ihc_root = os.path.join(work_dir, "特殊患者归档", "免疫组化")
    if os.path.exists(ihc_root):
        try:
            for entry in os.listdir(ihc_root):
                full = os.path.join(ihc_root, entry)
                if os.path.isdir(full) and (entry.startswith(f"{prefix}_") or entry == prefix):
                    return full
        except Exception:
            pass
    return ""


def find_patient_ihc_archive_path(work_dir, pathology_no, patient_name):
    """Finds the patient's archived IHC report (*_免疫组化报告.docx) in 工作目录/已归档报告/."""
    p_no_clean = sanitize_filename(str(pathology_no).strip())
    p_name_clean = sanitize_filename(str(patient_name).strip())
    prefix = f"{p_no_clean}_{p_name_clean}"
    archive_root = get_archive_root_dir(work_dir)
    if os.path.exists(archive_root):
        try:
            for entry in os.listdir(archive_root):
                if (entry.startswith(f"{prefix}_") or entry.startswith(f"{prefix}.")) and is_ihc_archived_filename(entry):
                    return os.path.join(archive_root, entry)
        except Exception:
            pass
    return ""


def create_ihc_report_package(work_dir, patient_data, source_archive_file):
    """
    Create an independent IHC report package in 工作目录/特殊患者归档/免疫组化/<folder>/
    by copying the patient's archived report and pathology photo directory (WITHOUT clinical photos).
    """
    if not source_archive_file or not os.path.exists(source_archive_file) or not os.path.isfile(source_archive_file):
        raise FileNotFoundError("未找到患者的已归档原报告文件，请先将原报告归档后再创建免疫组化报告")

    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    source_filename = os.path.basename(source_archive_file)
    source_stem = os.path.splitext(source_filename)[0]

    ihc_root = os.path.join(work_dir, "特殊患者归档", "免疫组化")
    os.makedirs(ihc_root, exist_ok=True)
    ihc_dir = os.path.join(ihc_root, source_stem)
    ihc_photo_dir = os.path.join(ihc_dir, "病理照片")
    os.makedirs(ihc_dir, exist_ok=True)
    os.makedirs(ihc_photo_dir, exist_ok=True)

    # Preserve any photos if an older-named IHC folder existed for this patient
    prefix = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}"
    for entry in list(os.listdir(ihc_root)):
        if entry != source_stem and (entry.startswith(f"{prefix}_") or entry == prefix):
            old_dir = os.path.join(ihc_root, entry)
            if os.path.isdir(old_dir):
                old_photos = os.path.join(old_dir, "病理照片")
                if os.path.isdir(old_photos):
                    for item in os.listdir(old_photos):
                        s = os.path.join(old_photos, item)
                        d = os.path.join(ihc_photo_dir, item)
                        if os.path.isfile(s) and not os.path.exists(d):
                            shutil.copy2(s, d)
                safe_delete_path(old_dir)

    # Copy pathology photos (NO clinical photos copied!)
    src_photo_dir = patient_data.get("photo_dir") or get_patient_photo_dir(work_dir, p_no, p_name)
    if src_photo_dir and os.path.exists(src_photo_dir):
        for item in os.listdir(src_photo_dir):
            s = os.path.join(src_photo_dir, item)
            d = os.path.join(ihc_photo_dir, item)
            if os.path.isfile(s) and not os.path.exists(d):
                shutil.copy2(s, d)

    # Copy archived report docx into IHC directory
    ihc_doc_path = os.path.join(ihc_dir, f"[病理报告]{source_stem}.docx")
    if os.path.exists(ihc_doc_path):
        try:
            with open(ihc_doc_path, "a+b"):
                pass
        except (PermissionError, IOError) as pe:
            raise PermissionError(
                f"免疫组化报告文件被其他软件（如 Word 或 WPS）打开占用，无法覆盖！请先关闭该文件后重试。路径: {os.path.basename(ihc_doc_path)}"
            ) from pe
    shutil.copy2(source_archive_file, ihc_doc_path)

    # Remap selected_images paths to the copied IHC photo directory
    orig_selected = patient_data.get("selected_images") or []
    remapped_images = []
    if isinstance(orig_selected, list):
        for idx, item in enumerate(orig_selected):
            if isinstance(item, dict):
                item_copy = dict(item)
                p = item_copy.get("path")
                if p:
                    candidate = os.path.join(ihc_photo_dir, os.path.basename(p))
                    item_copy["path"] = candidate if os.path.exists(candidate) else None
                remapped_images.append(item_copy)
            elif isinstance(item, str):
                candidate = os.path.join(ihc_photo_dir, os.path.basename(item))
                remapped_images.append({"slot": idx + 1, "path": candidate if os.path.exists(candidate) else None, "site": ""})

    return {
        "ihc_dir": ihc_dir,
        "ihc_doc_path": ihc_doc_path,
        "ihc_photo_dir": ihc_photo_dir,
        "source_archive_filename": source_filename,
        "selected_images": remapped_images
    }


def archive_ihc_report(work_dir, patient_data, ihc_data):
    """
    Archive the saved IHC report document (document only, no photos, no move to 完成审核)
    into 工作目录/已归档报告/<原已归档文件名>_免疫组化报告.docx without touching the original report.
    """
    ihc_doc_path = (ihc_data or {}).get("ihc_doc_path", "")
    if not ihc_doc_path or not os.path.exists(ihc_doc_path):
        return {"success": False, "error": "未找到已保存的免疫组化报告文件，请先创建并保存免疫组化报告"}

    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    prefix = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}_"

    src_name = (ihc_data or {}).get("source_archive_filename", "")
    if not src_name:
        orig_arch = find_patient_archive_path(work_dir, p_no, p_name)
        if orig_arch:
            src_name = os.path.basename(orig_arch)

    if src_name:
        base_stem = os.path.splitext(src_name)[0]
    else:
        p_diag = (ihc_data or {}).get("pathology_diag") or patient_data.get("pathology_diag") or patient_data.get("clinical_diagnosis") or "已诊断"
        base_stem = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}_{sanitize_filename(p_diag)}"

    if base_stem.endswith("_免疫组化报告"):
        base_stem = base_stem[:-len("_免疫组化报告")]

    target_filename = f"{base_stem}{IHC_ARCHIVE_SUFFIX}"
    archive_root = get_archive_root_dir(work_dir)
    target_archive_file = os.path.join(archive_root, target_filename)

    # Remove only stale *_免疫组化报告.docx for this patient, NEVER touch the original archived report
    if os.path.exists(archive_root):
        for item in os.listdir(archive_root):
            if (item.startswith(prefix) or item.startswith(f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}.")) and is_ihc_archived_filename(item):
                old_f = os.path.join(archive_root, item)
                if old_f != target_archive_file:
                    safe_delete_path(old_f)

    if os.path.exists(target_archive_file):
        try:
            with open(target_archive_file, "a+b"):
                pass
        except (PermissionError, IOError) as pe:
            raise PermissionError(
                f"已归档免疫组化报告文件被其他软件（如 Word 或 WPS）打开占用，无法覆盖！请先关闭该文件后重试。路径: {target_filename}"
            ) from pe

    shutil.copy2(ihc_doc_path, target_archive_file)
    return {
        "success": True,
        "archive_file": target_archive_file,
        "filename": target_filename
    }


def revoke_ihc_archive_files(work_dir, patient_data, ihc_data=None):
    """
    Removes ONLY the archived IHC report (*_免疫组化报告.docx) from 已归档报告/,
    preserving the original archived report and 特殊患者归档/免疫组化/.
    """
    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    prefix = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}"
    removed = []

    arch_file = (ihc_data or {}).get("archive_file", "")
    if arch_file and os.path.exists(arch_file) and is_ihc_archived_filename(arch_file):
        safe_delete_path(arch_file)
        removed.append(arch_file)

    archive_root = get_archive_root_dir(work_dir)
    if os.path.exists(archive_root):
        for entry in os.listdir(archive_root):
            if (entry.startswith(f"{prefix}_") or entry.startswith(f"{prefix}.")) and is_ihc_archived_filename(entry):
                full = os.path.join(archive_root, entry)
                safe_delete_path(full)
                removed.append(full)

    return {"success": True, "removed": removed}


def locate_patient_all_paths(work_dir, patient_data, ihc_data=None):
    """
    Returns a list of all existing files and directories associated with a patient.
    """
    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    p_diag = patient_data.get("pathology_diag") or patient_data.get("clinical_diagnosis") or ""
    safe_folder_name = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}_{sanitize_filename(p_diag)}" if p_diag else ""

    paths = []

    # 1. Draft Word Report
    draft_p = patient_data.get("word_path") or get_patient_word_path(work_dir, p_no, p_name)
    paths.append({
        "key": "word_draft",
        "name": "病理报告录入 (草稿)",
        "path": draft_p,
        "is_file": True,
        "exists": os.path.exists(draft_p)
    })

    # 2. Final Output Word Report
    out_p = patient_data.get("final_report_path")
    if not out_p and p_diag:
        out_p = get_final_report_path(work_dir, p_no, p_name, p_diag)
    paths.append({
        "key": "word_final",
        "name": "病理报告输出 (成品)",
        "path": out_p or "",
        "is_file": True,
        "exists": bool(out_p and os.path.exists(out_p))
    })

    # 3. Pathology Photos Dir
    photo_p = patient_data.get("photo_dir") or get_patient_photo_dir(work_dir, p_no, p_name)
    paths.append({
        "key": "photo_dir",
        "name": "病理照片文件夹",
        "path": photo_p,
        "is_file": False,
        "exists": os.path.exists(photo_p)
    })

    # 4. Clinical Photos Dir
    clin_lookup = find_clinical_photo_folder(work_dir, p_name, p_no)
    paths.append({
        "key": "clinical_photo_dir",
        "name": "临床照片文件夹",
        "path": clin_lookup["path"],
        "is_file": False,
        "exists": clin_lookup["found"]
    })

    # 5. Archive Report in 已归档报告
    arch_file = patient_data.get("archive_dir")
    if not arch_file or not os.path.exists(arch_file):
        arch_file = find_patient_archive_path(work_dir, p_no, p_name)
    paths.append({
        "key": "archive_dir",
        "name": "已归档报告 (成品文件)",
        "path": arch_file or "",
        "is_file": True if arch_file and os.path.isfile(arch_file) else False,
        "exists": bool(arch_file and os.path.exists(arch_file))
    })

    # 6. Special Archive (待审核)
    spec_pending = _find_special_cat_folder(work_dir, "待审核", p_no, p_name) or (
        os.path.join(work_dir, "特殊患者归档", "待审核", safe_folder_name) if safe_folder_name else ""
    )
    paths.append({
        "key": "special_pending",
        "name": "特殊患者归档 (待审核)",
        "path": spec_pending,
        "is_file": False,
        "exists": bool(spec_pending and os.path.exists(spec_pending))
    })

    # 7. Special Archive (免疫组化)
    spec_ihc = (ihc_data or {}).get("ihc_dir") or find_patient_ihc_folder(work_dir, p_no, p_name)
    if not spec_ihc and safe_folder_name:
        spec_ihc = os.path.join(work_dir, "特殊患者归档", "免疫组化", safe_folder_name)
    paths.append({
        "key": "special_ihc",
        "name": "特殊患者归档 (免疫组化)",
        "path": spec_ihc or "",
        "is_file": False,
        "exists": bool(spec_ihc and os.path.exists(spec_ihc))
    })

    # 8. Archived IHC Report in 已归档报告
    ihc_arch_file = (ihc_data or {}).get("archive_file")
    if not ihc_arch_file or not os.path.exists(ihc_arch_file):
        ihc_arch_file = find_patient_ihc_archive_path(work_dir, p_no, p_name)
    paths.append({
        "key": "ihc_archive_file",
        "name": "已归档免疫组化报告 (成品文件)",
        "path": ihc_arch_file or "",
        "is_file": bool(ihc_arch_file and os.path.isfile(ihc_arch_file)),
        "exists": bool(ihc_arch_file and os.path.exists(ihc_arch_file))
    })

    # 9. Special Archive (重切)
    spec_recut = _find_special_cat_folder(work_dir, "重切", p_no, p_name) or (
        os.path.join(work_dir, "特殊患者归档", "重切", safe_folder_name) if safe_folder_name else ""
    )
    paths.append({
        "key": "special_recut",
        "name": "特殊患者归档 (重切)",
        "path": spec_recut,
        "is_file": False,
        "exists": bool(spec_recut and os.path.exists(spec_recut))
    })

    # 10. Special Archive (完成审核)
    spec_approved = _find_special_cat_folder(work_dir, "完成审核", p_no, p_name) or (
        os.path.join(work_dir, "特殊患者归档", "完成审核", safe_folder_name) if safe_folder_name else ""
    )
    paths.append({
        "key": "special_approved",
        "name": "特殊患者归档 (完成审核)",
        "path": spec_approved,
        "is_file": False,
        "exists": bool(spec_approved and os.path.exists(spec_approved))
    })

    return paths

def delete_patient_all_files(work_dir, patient_data):
    """
    Deletes all files and directories related to the patient into Windows Recycle Bin:
    - Draft word report
    - Output word report
    - Pathology photo folder
    - Archived report file in 已归档报告
    - Special archive folders
    """
    deleted_items = []
    errors = []

    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    prefix = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}"

    # 1. Draft Word
    draft_p = patient_data.get("word_path") or get_patient_word_path(work_dir, p_no, p_name)
    if draft_p and os.path.exists(draft_p):
        try:
            safe_delete_path(draft_p)
            deleted_items.append(draft_p)
        except Exception as e:
            errors.append(f"删除草稿文件失败: {e}")

    # 2. Output Word
    out_dir = get_output_report_dir(work_dir)
    if os.path.exists(out_dir):
        for f in os.listdir(out_dir):
            if (f.startswith(f"{prefix}_") or f.startswith(f"{prefix}.")) and f.endswith(".docx"):
                fp = os.path.join(out_dir, f)
                try:
                    safe_delete_path(fp)
                    deleted_items.append(fp)
                except Exception:
                    pass

    # 3. Pathology photos folder
    photo_p = patient_data.get("photo_dir") or get_patient_photo_dir(work_dir, p_no, p_name)
    if photo_p and os.path.exists(photo_p):
        try:
            safe_delete_path(photo_p)
            deleted_items.append(photo_p)
        except Exception as e:
            errors.append(f"删除病理照片目录失败: {e}")

    # 4. Archived report in 已归档报告/ (and legacy 按患者归档/)
    for arch_name in ["已归档报告", "按患者归档"]:
        arch_root = os.path.join(work_dir, arch_name)
        if os.path.exists(arch_root):
            for entry in os.listdir(arch_root):
                if entry.startswith(f"{prefix}_") or entry.startswith(f"{prefix}."):
                    target = os.path.join(arch_root, entry)
                    try:
                        safe_delete_path(target)
                        deleted_items.append(target)
                    except Exception as e:
                        errors.append(f"删除归档报告失败: {e}")

    # 5. Special archives
    for sub in ["待审核", "免疫组化", "重切", "完成审核"]:
        spec_root = os.path.join(work_dir, "特殊患者归档", sub)
        if os.path.exists(spec_root):
            for entry in os.listdir(spec_root):
                if entry.startswith(f"{prefix}_") or entry == prefix:
                    target = os.path.join(spec_root, entry)
                    try:
                        safe_delete_path(target)
                        deleted_items.append(target)
                    except Exception as e:
                        errors.append(f"删除特殊归档失败: {e}")

    return {
        "success": len(errors) == 0,
        "deleted_items": deleted_items,
        "errors": errors
    }

def sync_status_file_operations(work_dir, patient_data, new_review_status, new_archived_status):
    """
    Synchronizes files according to manual status change from management interface:
    - If new_archived_status == 0: removes archive from 已归档报告 and 特殊患者归档
    - If new_archived_status == 1: packages archive if final report exists
    - Handles special archive copying / removing based on review status
    """
    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    p_diag = patient_data.get("pathology_diag") or patient_data.get("clinical_diagnosis") or "已诊断"
    prefix = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}"

    # 1. If un-archiving (new_archived_status == 0)
    if int(new_archived_status) == 0:
        revoke_patient_archive_files(work_dir, patient_data)

    # 2. If archiving (new_archived_status == 1)
    elif int(new_archived_status) == 1:
        report_f = patient_data.get("final_report_path") or patient_data.get("word_path")
        if report_f and os.path.exists(report_f):
            pat_copy = dict(patient_data)
            pat_copy["is_pending_review"] = 1 if "待审核" in new_review_status else 0
            pat_copy["is_needs_ihc"] = 1 if "免疫组化" in new_review_status else 0
            pat_copy["is_needs_recut"] = 1 if "重切" in new_review_status else 0
            pat_copy["review_status"] = new_review_status
            archive_patient_report(work_dir, pat_copy, report_f, is_from_review=(new_review_status == "完成审核"))

    # 3. Synchronize special archives based on review status
    if int(new_archived_status) in [1, 2]:
        report_f = patient_data.get("final_report_path") or patient_data.get("word_path")
        save_special_draft_queue(work_dir, patient_data, report_f)

    return {"success": True}

def get_signing_sheet_dir(work_dir):
    """Returns the path to '工作目录/领取表'."""
    p = os.path.join(work_dir, "领取表")
    os.makedirs(p, exist_ok=True)
    return p

def find_patient_archive_path(work_dir, pathology_no, patient_name):
    """
    Finds the patient's original (non-IHC) archived report in '工作目录/已归档报告/'.
    Matches prefix '病理号_姓名' and excludes '*_免疫组化报告.docx'. Returns file path if found.
    Also supports legacy subdirectories under 已归档报告/ or 按患者归档/ for backwards compatibility.
    """
    p_no_clean = sanitize_filename(str(pathology_no).strip())
    p_name_clean = sanitize_filename(str(patient_name).strip())
    prefix = f"{p_no_clean}_{p_name_clean}"

    # 1. Search in 已归档报告/
    archive_root = get_archive_root_dir(work_dir)
    if os.path.exists(archive_root):
        try:
            for entry in os.listdir(archive_root):
                if (entry.startswith(f"{prefix}_") or entry.startswith(f"{prefix}.")) and not is_ihc_archived_filename(entry):
                    return os.path.join(archive_root, entry)
        except Exception:
            pass

    # 2. Fallback to legacy 按患者归档/ if exists
    legacy_root = os.path.join(work_dir, "按患者归档")
    if os.path.exists(legacy_root):
        try:
            for entry in os.listdir(legacy_root):
                if (entry.startswith(f"{prefix}_") or entry.startswith(f"{prefix}.")) and not is_ihc_archived_filename(entry):
                    return os.path.join(legacy_root, entry)
        except Exception:
            pass

    return ""

def find_patient_archive_dir(work_dir, pathology_no, patient_name):
    """Alias for backwards compatibility."""
    return find_patient_archive_path(work_dir, pathology_no, patient_name)

def revoke_patient_archive_files(work_dir, patient_data):
    """
    Removes the patient's original archived report from 已归档报告/ and review 特殊患者归档 folders.
    Preserves *_免疫组化报告.docx and 特殊患者归档/免疫组化/.
    """
    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    p_no_clean = sanitize_filename(str(p_no).strip())
    p_name_clean = sanitize_filename(str(p_name).strip())
    prefix = f"{p_no_clean}_{p_name_clean}"

    removed = []

    # 1. 已归档报告 (and legacy 按患者归档), excluding *_免疫组化报告.docx
    for arch_name in ["已归档报告", "按患者归档"]:
        arch_root = os.path.join(work_dir, arch_name)
        if os.path.exists(arch_root):
            for entry in os.listdir(arch_root):
                full = os.path.join(arch_root, entry)
                if (entry.startswith(f"{prefix}_") or entry.startswith(f"{prefix}.")) and not is_ihc_archived_filename(entry):
                    safe_delete_path(full)
                    removed.append(full)

    # 2. 特殊患者归档 (待审核, 重切, 完成审核; excluding 免疫组化)
    for sub in ["待审核", "重切", "完成审核"]:
        spec_root = os.path.join(work_dir, "特殊患者归档", sub)
        if os.path.exists(spec_root):
            for entry in os.listdir(spec_root):
                full = os.path.join(spec_root, entry)
                if entry.startswith(f"{prefix}_") or entry == prefix:
                    safe_delete_path(full)
                    removed.append(full)

    return {"success": True, "removed": removed}

def revoke_patient_approval_files(work_dir, patient_data, new_status):
    """
    Copies patient archive from 特殊患者归档/完成审核 back to 特殊患者归档 (待审核/重切) when approval is revoked.
    Does not touch 特殊患者归档/免疫组化/.
    """
    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    p_no_clean = sanitize_filename(str(p_no).strip())
    p_name_clean = sanitize_filename(str(p_name).strip())
    prefix = f"{p_no_clean}_{p_name_clean}"

    approved_root = os.path.join(work_dir, "特殊患者归档", "完成审核")
    arch_dir = ""

    if os.path.exists(approved_root):
        for entry in os.listdir(approved_root):
            full = os.path.join(approved_root, entry)
            if os.path.isdir(full) and (entry.startswith(f"{prefix}_") or entry == prefix):
                arch_dir = full
                break

    if not arch_dir or not os.path.exists(arch_dir):
        return {"success": True, "message": "未找到归档源文件夹"}

    folder_name = os.path.basename(arch_dir)

    for sub in ("待审核", "重切"):
        if sub in new_status:
            spec_target = os.path.join(work_dir, "特殊患者归档", sub, folder_name)
            os.makedirs(os.path.join(work_dir, "特殊患者归档", sub), exist_ok=True)
            shutil.copytree(arch_dir, spec_target, dirs_exist_ok=True)

    safe_delete_path(arch_dir)

    return {"success": True}
