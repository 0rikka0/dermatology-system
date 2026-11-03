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
    Renames Word report file, output report file, and photo folder if pathology number or patient name changed.
    """
    old_word = get_patient_word_path(work_dir, old_no, old_name)
    new_word = get_patient_word_path(work_dir, new_no, new_name)

    old_photo = get_patient_photo_dir(work_dir, old_no, old_name)
    new_photo = get_patient_photo_dir(work_dir, new_no, new_name)

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
    out_dir = get_output_report_dir(work_dir)
    if os.path.exists(out_dir):
        old_prefix = f"{sanitize_filename(old_no)}_{sanitize_filename(old_name)}_"
        for f in os.listdir(out_dir):
            if f.startswith(old_prefix) and f.endswith(".docx"):
                old_f = os.path.join(out_dir, f)
                diag_part = f[len(old_prefix):]
                new_f = os.path.join(out_dir, f"{sanitize_filename(new_no)}_{sanitize_filename(new_name)}_{diag_part}")
                try:
                    shutil.move(old_f, new_f)
                except Exception:
                    pass

    # 3. Rename archived report file in 已归档报告 if exists
    arch_root = get_archive_root_dir(work_dir)
    if os.path.exists(arch_root):
        old_prefix = f"{sanitize_filename(old_no)}_{sanitize_filename(old_name)}_"
        for f in os.listdir(arch_root):
            if f.startswith(old_prefix) and f.endswith(".docx"):
                old_f = os.path.join(arch_root, f)
                diag_part = f[len(old_prefix):]
                new_f = os.path.join(arch_root, f"{sanitize_filename(new_no)}_{sanitize_filename(new_name)}_{diag_part}")
                try:
                    shutil.move(old_f, new_f)
                except Exception:
                    pass

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
    return {"word_path": new_word, "photo_dir": new_photo}

def save_special_draft_queue(work_dir, patient_data, report_file_path):
    """
    When saving a report marked as 待审核 or 免疫组化:
    Packages the patient's photos and report into '特殊患者归档/待审核' and/or '特殊患者归档/免疫组化'.
    If a tag was unchecked, removes it from that category.
    If clinical photos have multiple matches (重名):
    Flags duplicate_clinical_photos = True so server can pop up prompt and open both directories in Explorer!
    """
    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    p_diag = patient_data.get("pathology_diag") or patient_data.get("clinical_diagnosis") or "已诊断"
    safe_folder_name = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}_{sanitize_filename(p_diag)}"
    prefix = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}_"

    is_pending = int(patient_data.get("is_pending_review", 0)) == 1 or "待审核" in str(patient_data.get("review_status", ""))
    is_ihc = int(patient_data.get("is_needs_ihc", 0)) == 1 or "免疫组化" in str(patient_data.get("review_status", ""))

    saved_dirs = []
    duplicate_info = {
        "has_duplicate": False,
        "matches": [],
        "clinical_root": get_clinical_photo_root_dir(work_dir),
        "target_dirs": []
    }

    clin_lookup = find_clinical_photo_folder(work_dir, p_name, p_no)
    if clin_lookup.get("multiple"):
        duplicate_info["has_duplicate"] = True
        duplicate_info["matches"] = clin_lookup.get("matches", [])

    def package_to_dir(target_dir):
        os.makedirs(target_dir, exist_ok=True)
        clin_target = os.path.join(target_dir, "临床照片")
        os.makedirs(clin_target, exist_ok=True)

        if duplicate_info["has_duplicate"]:
            if target_dir not in duplicate_info["target_dirs"]:
                duplicate_info["target_dirs"].append(target_dir)
        elif clin_lookup.get("found"):
            src_clin = clin_lookup["path"]
            if os.path.exists(src_clin):
                for item in os.listdir(src_clin):
                    s = os.path.join(src_clin, item)
                    d = os.path.join(clin_target, item)
                    if os.path.isfile(s) and not os.path.exists(d):
                        shutil.copy2(s, d)

        # 2. Pathology photos
        photo_target = os.path.join(target_dir, "病理照片")
        os.makedirs(photo_target, exist_ok=True)
        src_photo_dir = get_patient_photo_dir(work_dir, p_no, p_name)
        if os.path.exists(src_photo_dir):
            for item in os.listdir(src_photo_dir):
                s = os.path.join(src_photo_dir, item)
                d = os.path.join(photo_target, item)
                if os.path.isfile(s) and not os.path.exists(d):
                    shutil.copy2(s, d)

        # 3. Report docx (keep [病理报告] prefix in 特殊患者归档 as original specification)
        if report_file_path and os.path.exists(report_file_path):
            archived_report_name = f"[病理报告]{safe_folder_name}.docx"
            archived_report_path = os.path.join(target_dir, archived_report_name)
            shutil.copy2(report_file_path, archived_report_path)

    # 1. 待审核
    pending_root = os.path.join(work_dir, "特殊患者归档", "待审核")
    if is_pending:
        t_dir = os.path.join(pending_root, safe_folder_name)
        package_to_dir(t_dir)
        saved_dirs.append(t_dir)
    else:
        if os.path.exists(pending_root):
            for e in os.listdir(pending_root):
                if e.startswith(prefix) or e == f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}":
                    safe_delete_path(os.path.join(pending_root, e))

    # 2. 免疫组化
    ihc_root = os.path.join(work_dir, "特殊患者归档", "免疫组化")
    if is_ihc:
        t_dir = os.path.join(ihc_root, safe_folder_name)
        package_to_dir(t_dir)
        saved_dirs.append(t_dir)
    else:
        if os.path.exists(ihc_root):
            for e in os.listdir(ihc_root):
                if e.startswith(prefix) or e == f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}":
                    safe_delete_path(os.path.join(ihc_root, e))

    return {
        "success": True,
        "saved_dirs": saved_dirs,
        "duplicate_clinical_photos": duplicate_info["has_duplicate"],
        "duplicate_matches": duplicate_info["matches"],
        "clinical_root": duplicate_info["clinical_root"],
        "target_dirs": duplicate_info["target_dirs"]
    }

def archive_patient_report(work_dir, patient_data, report_file_path, is_from_review=False):
    """
    Archives a patient's report:
      1. Regular archives:
         Stored directly into: 工作目录/已归档报告/病理号_姓名_诊断.docx (Flat file, NO subfolders, NO '[病理报告]' prefix!)
         Photos are NOT copied to regular archive (solves file duplication / bloat).
      2. Special patients archives:
         Retains full package structure in 特殊患者归档/ (待审核, 免疫组化, 完成审核).
         If clinical photos have multiple matches (重名), alerts operator and opens both directories in Explorer for manual copy.
         Upon review completion, moves package from 待审核 / 免疫组化 into 特殊患者归档/完成审核/.
    """
    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    p_diag = patient_data.get("pathology_diag") or patient_data.get("clinical_diagnosis") or "已诊断"
    safe_folder_name = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}_{sanitize_filename(p_diag)}"
    prefix = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}_"

    is_still_pending = patient_data.get("review_status") in ("待审核", "免疫组化", "待审核/免疫组化") or (
        (int(patient_data.get("is_pending_review", 0)) == 1 or int(patient_data.get("is_needs_ihc", 0)) == 1)
        and patient_data.get("review_status") != "完成审核"
    )

    if not is_from_review and is_still_pending:
        return {
            "success": False,
            "error": "标记为“待审核”或“免疫组化”的报告禁止在报告制作工作区直接归档！已在保存时自动放入特殊患者待审队列，请移步【报告审核】工作区完成专家审核后再进行归档。"
        }

    # 1. Main archive in '已归档报告' (Flat docx report file, NO subfolders, NO '[病理报告]' prefix!)
    archive_root = get_archive_root_dir(work_dir)
    target_report_name = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}_{sanitize_filename(p_diag)}.docx"
    target_archive_file = os.path.join(archive_root, target_report_name)

    # Clean up any legacy or duplicate archived files for this patient in 已归档报告/
    if os.path.exists(archive_root):
        for item in os.listdir(archive_root):
            if (item.startswith(prefix) or item.startswith(f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}.")) and item.endswith(".docx"):
                old_f = os.path.join(archive_root, item)
                if old_f != target_archive_file:
                    safe_delete_path(old_f)

    # Copy final report to 已归档报告/
    if report_file_path and os.path.exists(report_file_path):
        shutil.copy2(report_file_path, target_archive_file)

    # 2. Special archive handling:
    special_paths = []
    duplicate_info = {
        "has_duplicate": False,
        "matches": [],
        "clinical_root": get_clinical_photo_root_dir(work_dir),
        "target_dirs": []
    }

    is_approved_special = (patient_data.get("review_status") == "完成审核") or is_from_review
    is_special_patient = (
        int(patient_data.get("is_pending_review", 0)) == 1 or
        int(patient_data.get("is_needs_ihc", 0)) == 1 or
        patient_data.get("review_status") in ("待审核", "免疫组化", "待审核/免疫组化", "完成审核")
    )

    if is_special_patient and is_approved_special:
        approved_root = os.path.join(work_dir, "特殊患者归档", "完成审核")
        os.makedirs(approved_root, exist_ok=True)
        approved_dir = os.path.join(approved_root, safe_folder_name)

        # Check if an existing package was in 待审核 or 免疫组化
        source_dir = None
        for sub in ["待审核", "免疫组化"]:
            sub_root = os.path.join(work_dir, "特殊患者归档", sub)
            if os.path.exists(sub_root):
                for e in os.listdir(sub_root):
                    if e.startswith(prefix) or e == safe_folder_name:
                        source_dir = os.path.join(sub_root, e)
                        break
            if source_dir:
                break

        if source_dir and os.path.exists(source_dir):
            if source_dir != approved_dir:
                shutil.copytree(source_dir, approved_dir, dirs_exist_ok=True)
                safe_delete_path(source_dir)
            clin_lookup = find_clinical_photo_folder(work_dir, p_name, p_no)
            if clin_lookup.get("multiple"):
                duplicate_info["has_duplicate"] = True
                duplicate_info["matches"] = clin_lookup.get("matches", [])
                if approved_dir not in duplicate_info["target_dirs"]:
                    duplicate_info["target_dirs"].append(approved_dir)
        else:
            os.makedirs(approved_dir, exist_ok=True)
            clin_target = os.path.join(approved_dir, "临床照片")
            os.makedirs(clin_target, exist_ok=True)

            clin_lookup = find_clinical_photo_folder(work_dir, p_name, p_no)
            if clin_lookup.get("multiple"):
                duplicate_info["has_duplicate"] = True
                duplicate_info["matches"] = clin_lookup.get("matches", [])
                if approved_dir not in duplicate_info["target_dirs"]:
                    duplicate_info["target_dirs"].append(approved_dir)
            elif clin_lookup.get("found"):
                src_clin = clin_lookup["path"]
                if os.path.exists(src_clin):
                    for item in os.listdir(src_clin):
                        s = os.path.join(src_clin, item)
                        d = os.path.join(clin_target, item)
                        if os.path.isfile(s) and not os.path.exists(d):
                            shutil.copy2(s, d)

            photo_target = os.path.join(approved_dir, "病理照片")
            os.makedirs(photo_target, exist_ok=True)
            src_photo_dir = get_patient_photo_dir(work_dir, p_no, p_name)
            if os.path.exists(src_photo_dir):
                for item in os.listdir(src_photo_dir):
                    s = os.path.join(src_photo_dir, item)
                    d = os.path.join(photo_target, item)
                    if os.path.isfile(s) and not os.path.exists(d):
                        shutil.copy2(s, d)

            if report_file_path and os.path.exists(report_file_path):
                archived_report_name = f"[病理报告]{safe_folder_name}.docx"
                shutil.copy2(report_file_path, os.path.join(approved_dir, archived_report_name))

        # Clean remaining in 待审核 and 免疫组化
        for sub in ["待审核", "免疫组化"]:
            sub_root = os.path.join(work_dir, "特殊患者归档", sub)
            if os.path.exists(sub_root):
                for e in os.listdir(sub_root):
                    if e.startswith(prefix) or e == safe_folder_name:
                        safe_delete_path(os.path.join(sub_root, e))

        special_paths.append(approved_dir)

    return {
        "success": True,
        "archive_dir": target_archive_file,
        "archive_file": target_archive_file,
        "special_paths": special_paths,
        "duplicate_clinical_photos": duplicate_info["has_duplicate"],
        "duplicate_matches": duplicate_info["matches"],
        "clinical_root": duplicate_info["clinical_root"],
        "target_dirs": duplicate_info["target_dirs"]
    }

def sync_review_overwrites(work_dir, patient_data, new_report_path):
    """
    Overwrites report file across:
      1. 未归档报告/病理报告输出
      2. 已归档报告/<pathology_no>_<name>_<diag>.docx
      3. 特殊患者归档/<cat>/<patient_folder>
    """
    if not os.path.exists(new_report_path):
        return {"success": False, "error": "Report file does not exist"}

    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    p_diag = patient_data.get("pathology_diag") or patient_data.get("clinical_diagnosis") or "已诊断"
    safe_folder_name = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}_{sanitize_filename(p_diag)}"
    archived_filename = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}_{sanitize_filename(p_diag)}.docx"
    prefix = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}_"

    overwritten = []

    # 1. 已归档报告 (flat file)
    arch_root = get_archive_root_dir(work_dir)
    target_arch_file = os.path.join(arch_root, archived_filename)
    if os.path.exists(arch_root):
        for item in os.listdir(arch_root):
            if (item.startswith(prefix) or item.startswith(f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}.")) and item.endswith(".docx"):
                old_f = os.path.join(arch_root, item)
                if old_f != target_arch_file:
                    safe_delete_path(old_f)
        if patient_data.get("archived") in (1, 2) or os.path.exists(target_arch_file):
            shutil.copy2(new_report_path, target_arch_file)
            overwritten.append(target_arch_file)

    # 2. 特殊患者归档
    archived_special_doc_name = f"[病理报告]{safe_folder_name}.docx"
    for sub in ["待审核", "免疫组化", "完成审核"]:
        spec_patient_dir = os.path.join(work_dir, "特殊患者归档", sub, safe_folder_name)
        if os.path.exists(spec_patient_dir):
            dst = os.path.join(spec_patient_dir, archived_special_doc_name)
            shutil.copy2(new_report_path, dst)
            overwritten.append(dst)

    return {"success": True, "overwritten": overwritten}

def locate_patient_all_paths(work_dir, patient_data):
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
    spec_pending = os.path.join(work_dir, "特殊患者归档", "待审核", safe_folder_name) if safe_folder_name else ""
    paths.append({
        "key": "special_pending",
        "name": "特殊患者归档 (待审核)",
        "path": spec_pending,
        "is_file": False,
        "exists": bool(spec_pending and os.path.exists(spec_pending))
    })

    # 7. Special Archive (免疫组化)
    spec_ihc = os.path.join(work_dir, "特殊患者归档", "免疫组化", safe_folder_name) if safe_folder_name else ""
    paths.append({
        "key": "special_ihc",
        "name": "特殊患者归档 (免疫组化)",
        "path": spec_ihc,
        "is_file": False,
        "exists": bool(spec_ihc and os.path.exists(spec_ihc))
    })

    # 8. Special Archive (完成审核)
    spec_approved = os.path.join(work_dir, "特殊患者归档", "完成审核", safe_folder_name) if safe_folder_name else ""
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
    for sub in ["待审核", "免疫组化", "完成审核"]:
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
    Finds the patient's archived report in '工作目录/已归档报告/'.
    Matches prefix '病理号_姓名'. Returns file path if found.
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
                full = os.path.join(archive_root, entry)
                if entry.startswith(prefix):
                    return full
        except Exception:
            pass

    # 2. Fallback to legacy 按患者归档/ if exists
    legacy_root = os.path.join(work_dir, "按患者归档")
    if os.path.exists(legacy_root):
        try:
            for entry in os.listdir(legacy_root):
                full = os.path.join(legacy_root, entry)
                if entry.startswith(prefix):
                    return full
        except Exception:
            pass

    return ""

def find_patient_archive_dir(work_dir, pathology_no, patient_name):
    """Alias for backwards compatibility."""
    return find_patient_archive_path(work_dir, pathology_no, patient_name)

def revoke_patient_archive_files(work_dir, patient_data):
    """
    Removes the patient's archived report from 已归档报告/ and any 特殊患者归档 folders.
    """
    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    p_no_clean = sanitize_filename(str(p_no).strip())
    p_name_clean = sanitize_filename(str(p_name).strip())
    prefix = f"{p_no_clean}_{p_name_clean}"

    removed = []

    # 1. 已归档报告 (and legacy 按患者归档)
    for arch_name in ["已归档报告", "按患者归档"]:
        arch_root = os.path.join(work_dir, arch_name)
        if os.path.exists(arch_root):
            for entry in os.listdir(arch_root):
                full = os.path.join(arch_root, entry)
                if entry.startswith(f"{prefix}_") or entry.startswith(f"{prefix}."):
                    safe_delete_path(full)
                    removed.append(full)

    # 2. 特殊患者归档 (待审核, 免疫组化, 完成审核)
    for sub in ["待审核", "免疫组化", "完成审核"]:
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
    Copies patient archive from 特殊患者归档/完成审核 back to 特殊患者归档 when approval is revoked.
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

    if "待审核" in new_status:
        spec_target = os.path.join(work_dir, "特殊患者归档", "待审核", folder_name)
        os.makedirs(os.path.join(work_dir, "特殊患者归档", "待审核"), exist_ok=True)
        shutil.copytree(arch_dir, spec_target, dirs_exist_ok=True)

    if "免疫组化" in new_status:
        spec_target = os.path.join(work_dir, "特殊患者归档", "免疫组化", folder_name)
        os.makedirs(os.path.join(work_dir, "特殊患者归档", "免疫组化"), exist_ok=True)
        shutil.copytree(arch_dir, spec_target, dirs_exist_ok=True)

    safe_delete_path(arch_dir)

    return {"success": True}
