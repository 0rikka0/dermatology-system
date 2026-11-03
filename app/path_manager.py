# -*- coding: utf-8 -*-
"""
Dermatology Pathology Reporting System - Path & Directory Management
"""
import os
import re
import shutil

SUBDIRECTORIES = [
    "按患者归档",
    "数据库",
    os.path.join("特殊患者归档", "待审核"),
    os.path.join("特殊患者归档", "免疫组化"),
    os.path.join("特殊患者归档", "完成审核"),
    os.path.join("统一归档", "临床照片"),
    os.path.join("统一归档", "病理报告录入"),
    os.path.join("统一归档", "病理报告输出"),
    os.path.join("统一归档", "病理照片"),
    "领取表",
    "文档"
]

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
    d = os.path.join(work_dir, "统一归档", "病理报告录入")
    os.makedirs(d, exist_ok=True)
    return d

def get_output_report_dir(work_dir):
    """Directory for finalized word reports (病理报告输出)."""
    d = os.path.join(work_dir, "统一归档", "病理报告输出")
    os.makedirs(d, exist_ok=True)
    return d

def get_photo_root_dir(work_dir):
    """Root directory for pathology photos (病理照片)."""
    d = os.path.join(work_dir, "统一归档", "病理照片")
    os.makedirs(d, exist_ok=True)
    return d

def get_clinical_photo_root_dir(work_dir):
    """Root directory for clinical photos (临床照片)."""
    d = os.path.join(work_dir, "统一归档", "临床照片")
    os.makedirs(d, exist_ok=True)
    return d

def get_archive_root_dir(work_dir):
    """Root directory for patient archives (按患者归档)."""
    d = os.path.join(work_dir, "按患者归档")
    os.makedirs(d, exist_ok=True)
    return d

def sanitize_filename(name):
    """Remove characters forbidden in Windows filenames except valid dots/semicolons."""
    # Forbidden: \ / : * ? " < > |
    s = re.sub(r'[\\/:*?"<>|]', '_', str(name).strip())
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

def get_patient_word_path(work_dir, pathology_no, name):
    """Path to draft word report in 病理报告录入."""
    return os.path.join(get_report_dir(work_dir), get_patient_word_filename(pathology_no, name))

def get_final_report_path(work_dir, pathology_no, name, diagnosis=""):
    """Path to final word report in 病理报告输出."""
    return os.path.join(get_output_report_dir(work_dir), get_final_report_filename(pathology_no, name, diagnosis))

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
    Looks for a folder matching patient_name in 统一归档/临床照片/.
    Matches exact name or subfolder containing the patient name.
    """
    root = get_clinical_photo_root_dir(work_dir)
    if not os.path.exists(root):
        os.makedirs(root, exist_ok=True)
        return {"found": False, "path": root, "root": root}

    clean_name = str(patient_name).strip()
    clean_no = str(pathology_no).strip()

    # Scan directories inside clinical photos
    try:
        entries = os.listdir(root)
    except Exception:
        entries = []

    # 1. Exact match with name or pathology_no
    for entry in entries:
        full = os.path.join(root, entry)
        if os.path.isdir(full):
            if entry == clean_name or entry == f"{clean_no}_{clean_name}" or entry == f"{clean_name}_{clean_no}":
                return {"found": True, "path": full, "folder_name": entry, "root": root}

    # 2. Substring match
    if clean_name:
        for entry in entries:
            full = os.path.join(root, entry)
            if os.path.isdir(full):
                if clean_name in entry:
                    return {"found": True, "path": full, "folder_name": entry, "root": root}

    return {"found": False, "path": root, "root": root}

def rename_patient_files_and_folders(work_dir, old_no, old_name, new_no, new_name):
    """
    Renames Word report file and photo folder if pathology number or patient name changed.
    """
    old_word = get_patient_word_path(work_dir, old_no, old_name)
    new_word = get_patient_word_path(work_dir, new_no, new_name)
    
    old_photo = get_patient_photo_dir(work_dir, old_no, old_name)
    new_photo = get_patient_photo_dir(work_dir, new_no, new_name)

    if old_word != new_word and os.path.exists(old_word):
        if not os.path.exists(new_word):
            shutil.move(old_word, new_word)

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
    """
    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    p_diag = patient_data.get("pathology_diag") or patient_data.get("clinical_diagnosis") or "已诊断"
    safe_folder_name = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}_{sanitize_filename(p_diag)}"
    prefix = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}"

    is_pending = int(patient_data.get("is_pending_review", 0)) == 1 or "待审核" in str(patient_data.get("review_status", ""))
    is_ihc = int(patient_data.get("is_needs_ihc", 0)) == 1 or "免疫组化" in str(patient_data.get("review_status", ""))

    saved_dirs = []

    # Helper to package a single special directory
    def package_to_dir(target_dir):
        os.makedirs(target_dir, exist_ok=True)
        # 1. Clinical photos
        clin_target = os.path.join(target_dir, "临床照片")
        os.makedirs(clin_target, exist_ok=True)
        clin_lookup = find_clinical_photo_folder(work_dir, p_name, p_no)
        if clin_lookup["found"]:
            src_clin = clin_lookup["path"]
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

        # 3. Report docx
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
                if e.startswith(prefix):
                    shutil.rmtree(os.path.join(pending_root, e), ignore_errors=True)

    # 2. 免疫组化
    ihc_root = os.path.join(work_dir, "特殊患者归档", "免疫组化")
    if is_ihc:
        t_dir = os.path.join(ihc_root, safe_folder_name)
        package_to_dir(t_dir)
        saved_dirs.append(t_dir)
    else:
        if os.path.exists(ihc_root):
            for e in os.listdir(ihc_root):
                if e.startswith(prefix):
                    shutil.rmtree(os.path.join(ihc_root, e), ignore_errors=True)

    return {"success": True, "saved_dirs": saved_dirs}

def archive_patient_report(work_dir, patient_data, report_file_path, is_from_review=False):
    """
    Archives a patient's report, clinical photos and pathology photos into:
      工作目录/按患者归档/病理号_患者姓名_诊断/
    Special patients logic:
      - If marked 待审核 or 免疫组化, directly archiving from report maker is PROHIBITED.
      - Upon completing review, archive removes from 待审核 / 免疫组化 and places into:
        工作目录/特殊患者归档/完成审核/
      - Simultaneously, standard archive in 按患者归档 is created.
    """
    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    p_diag = patient_data.get("pathology_diag") or patient_data.get("clinical_diagnosis") or "已诊断"
    safe_folder_name = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}_{sanitize_filename(p_diag)}"
    prefix = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}"

    is_still_pending = patient_data.get("review_status") in ("待审核", "免疫组化", "待审核/免疫组化") or (
        (int(patient_data.get("is_pending_review", 0)) == 1 or int(patient_data.get("is_needs_ihc", 0)) == 1)
        and patient_data.get("review_status") != "完成审核"
    )

    # If triggered from report maker and patient is still pending review/IHC: PROHIBIT direct archive
    if not is_from_review and is_still_pending:
        return {
            "success": False,
            "error": "标记为“待审核”或“免疫组化”的报告禁止在报告制作工作区直接归档！已在保存时自动放入特殊患者待审队列，请移步【报告审核】工作区完成专家审核后再进行归档。"
        }

    # 1. Main archive folder under 按患者归档
    archive_root = get_archive_root_dir(work_dir)
    target_archive_dir = os.path.join(archive_root, safe_folder_name)
    os.makedirs(target_archive_dir, exist_ok=True)

    # Copy / package clinical photos
    clin_target = os.path.join(target_archive_dir, "临床照片")
    os.makedirs(clin_target, exist_ok=True)
    clin_lookup = find_clinical_photo_folder(work_dir, p_name, p_no)
    if clin_lookup["found"]:
        src_clin = clin_lookup["path"]
        for item in os.listdir(src_clin):
            s = os.path.join(src_clin, item)
            d = os.path.join(clin_target, item)
            if os.path.isfile(s) and not os.path.exists(d):
                shutil.copy2(s, d)

    # Copy / package pathology photos
    photo_target = os.path.join(target_archive_dir, "病理照片")
    os.makedirs(photo_target, exist_ok=True)
    src_photo_dir = get_patient_photo_dir(work_dir, p_no, p_name)
    if os.path.exists(src_photo_dir):
        for item in os.listdir(src_photo_dir):
            s = os.path.join(src_photo_dir, item)
            d = os.path.join(photo_target, item)
            if os.path.isfile(s) and not os.path.exists(d):
                shutil.copy2(s, d)

    # Copy report docx file
    if os.path.exists(report_file_path):
        archived_report_name = f"[病理报告]{sanitize_filename(p_no)}_{sanitize_filename(p_name)}_{sanitize_filename(p_diag)}.docx"
        archived_report_path = os.path.join(target_archive_dir, archived_report_name)
        shutil.copy2(report_file_path, archived_report_path)

    # 2. Special archive completion:
    # If this patient has completed review, take out from 待审核 / 免疫组化 and put into 完成审核
    special_paths = []
    is_approved_special = (patient_data.get("review_status") == "完成审核") or is_from_review
    if is_approved_special:
        # Take out from 待审核
        pending_root = os.path.join(work_dir, "特殊患者归档", "待审核")
        if os.path.exists(pending_root):
            for e in os.listdir(pending_root):
                if e.startswith(prefix):
                    shutil.rmtree(os.path.join(pending_root, e), ignore_errors=True)

        # Take out from 免疫组化
        ihc_root = os.path.join(work_dir, "特殊患者归档", "免疫组化")
        if os.path.exists(ihc_root):
            for e in os.listdir(ihc_root):
                if e.startswith(prefix):
                    shutil.rmtree(os.path.join(ihc_root, e), ignore_errors=True)

        # Put into 完成审核
        approved_root = os.path.join(work_dir, "特殊患者归档", "完成审核")
        os.makedirs(approved_root, exist_ok=True)
        approved_dir = os.path.join(approved_root, safe_folder_name)
        shutil.copytree(target_archive_dir, approved_dir, dirs_exist_ok=True)
        special_paths.append(approved_dir)

    return {
        "success": True,
        "archive_dir": target_archive_dir,
        "special_paths": special_paths
    }

def sync_review_overwrites(work_dir, patient_data, new_report_path):
    """
    Overwrites report file across:
      1. 统一归档/病理报告输出
      2. 按患者归档/<patient_folder>
      3. 特殊患者归档/<cat>/<patient_folder>
    """
    if not os.path.exists(new_report_path):
        return {"success": False, "error": "Report file does not exist"}

    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    p_diag = patient_data.get("pathology_diag") or patient_data.get("clinical_diagnosis") or "已诊断"
    safe_folder_name = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}_{sanitize_filename(p_diag)}"
    archived_report_name = f"[病理报告]{sanitize_filename(p_no)}_{sanitize_filename(p_name)}_{sanitize_filename(p_diag)}.docx"

    overwritten = []

    # 1. 按患者归档
    arch_root = get_archive_root_dir(work_dir)
    arch_patient_dir = os.path.join(arch_root, safe_folder_name)
    if os.path.exists(arch_patient_dir):
        dst = os.path.join(arch_patient_dir, archived_report_name)
        shutil.copy2(new_report_path, dst)
        overwritten.append(dst)

    # 2. 特殊患者归档
    for sub in ["待审核", "免疫组化", "完成审核"]:
        spec_patient_dir = os.path.join(work_dir, "特殊患者归档", sub, safe_folder_name)
        if os.path.exists(spec_patient_dir):
            dst = os.path.join(spec_patient_dir, archived_report_name)
            shutil.copy2(new_report_path, dst)
            overwritten.append(dst)

    return {"success": True, "overwritten": overwritten}

def locate_patient_all_paths(work_dir, patient_data):
    """
    Returns a dictionary of all existing files and directories associated with a patient.
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

    # 5. Archive Folder
    arch_dir = patient_data.get("archive_dir")
    if not arch_dir and safe_folder_name:
        arch_dir = os.path.join(get_archive_root_dir(work_dir), safe_folder_name)
    paths.append({
        "key": "archive_dir",
        "name": "按患者归档目录",
        "path": arch_dir or "",
        "is_file": False,
        "exists": bool(arch_dir and os.path.exists(arch_dir))
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
    Deletes all files and directories related to the patient:
    - Draft word report
    - Output word report
    - Pathology photo folder
    - Patient archive folder
    - Special archive folders
    """
    deleted_items = []
    errors = []

    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    p_diag = patient_data.get("pathology_diag") or patient_data.get("clinical_diagnosis") or ""

    # 1. Draft Word
    draft_p = patient_data.get("word_path") or get_patient_word_path(work_dir, p_no, p_name)
    if draft_p and os.path.exists(draft_p):
        try:
            os.remove(draft_p)
            deleted_items.append(draft_p)
        except Exception as e:
            errors.append(f"删除草稿文件失败: {e}")

    # 2. Output Word
    out_p = patient_data.get("final_report_path")
    if out_p and os.path.exists(out_p):
        try:
            os.remove(out_p)
            deleted_items.append(out_p)
        except Exception as e:
            errors.append(f"删除输出报告失败: {e}")

    # Also search output dir for any matching files starting with p_no_p_name
    out_dir = get_output_report_dir(work_dir)
    if os.path.exists(out_dir):
        prefix = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}"
        for f in os.listdir(out_dir):
            if f.startswith(prefix) and f.endswith(".docx"):
                fp = os.path.join(out_dir, f)
                try:
                    os.remove(fp)
                    deleted_items.append(fp)
                except Exception:
                    pass

    # 3. Pathology photos folder
    photo_p = patient_data.get("photo_dir") or get_patient_photo_dir(work_dir, p_no, p_name)
    if photo_p and os.path.exists(photo_p):
        try:
            shutil.rmtree(photo_p)
            deleted_items.append(photo_p)
        except Exception as e:
            errors.append(f"删除病理照片目录失败: {e}")

    # 4. Patient archive folder
    arch_root = get_archive_root_dir(work_dir)
    if os.path.exists(arch_root):
        prefix = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}"
        for entry in os.listdir(arch_root):
            if entry.startswith(prefix):
                target = os.path.join(arch_root, entry)
                try:
                    shutil.rmtree(target)
                    deleted_items.append(target)
                except Exception as e:
                    errors.append(f"删除归档目录失败: {e}")

    # 5. Special archives
    for sub in ["待审核", "免疫组化", "完成审核"]:
        spec_root = os.path.join(work_dir, "特殊患者归档", sub)
        if os.path.exists(spec_root):
            prefix = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}"
            for entry in os.listdir(spec_root):
                if entry.startswith(prefix):
                    target = os.path.join(spec_root, entry)
                    try:
                        shutil.rmtree(target)
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
    - If new_archived_status == 0: removes archive folder from 按患者归档 and 特殊患者归档
    - If new_archived_status == 1: packages archive if final report exists
    - Handles special archive copying / removing based on review status
    """
    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    p_diag = patient_data.get("pathology_diag") or patient_data.get("clinical_diagnosis") or "已诊断"
    prefix = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}"

    # 1. If un-archiving (new_archived_status == 0)
    if int(new_archived_status) == 0:
        arch_root = get_archive_root_dir(work_dir)
        if os.path.exists(arch_root):
            for entry in os.listdir(arch_root):
                if entry.startswith(prefix):
                    shutil.rmtree(os.path.join(arch_root, entry), ignore_errors=True)
        for sub in ["待审核", "免疫组化", "完成审核"]:
            spec_root = os.path.join(work_dir, "特殊患者归档", sub)
            if os.path.exists(spec_root):
                for entry in os.listdir(spec_root):
                    if entry.startswith(prefix):
                        shutil.rmtree(os.path.join(spec_root, entry), ignore_errors=True)

    # 2. If archiving (new_archived_status == 1)
    elif int(new_archived_status) == 1:
        report_f = patient_data.get("final_report_path") or patient_data.get("word_path")
        if report_f and os.path.exists(report_f):
            # Temporarily set flags for archive_patient_report
            pat_copy = dict(patient_data)
            pat_copy["is_pending_review"] = 1 if "待审核" in new_review_status else 0
            pat_copy["is_needs_ihc"] = 1 if "免疫组化" in new_review_status else 0
            pat_copy["review_status"] = new_review_status
            archive_patient_report(work_dir, pat_copy, report_f, is_from_review=(new_review_status == "完成审核"))

    # 3. Synchronize special archives based on review status
    if int(new_archived_status) in [1, 2]:
        safe_folder_name = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}_{sanitize_filename(p_diag)}"
        arch_dir = os.path.join(get_archive_root_dir(work_dir), safe_folder_name)

        # 待审核
        pending_dir = os.path.join(work_dir, "特殊患者归档", "待审核", safe_folder_name)
        if "待审核" in new_review_status:
            if os.path.exists(arch_dir) and not os.path.exists(pending_dir):
                os.makedirs(os.path.join(work_dir, "特殊患者归档", "待审核"), exist_ok=True)
                shutil.copytree(arch_dir, pending_dir, dirs_exist_ok=True)
        else:
            if os.path.exists(pending_dir):
                shutil.rmtree(pending_dir, ignore_errors=True)

        # 免疫组化
        ihc_dir = os.path.join(work_dir, "特殊患者归档", "免疫组化", safe_folder_name)
        if "免疫组化" in new_review_status:
            if os.path.exists(arch_dir) and not os.path.exists(ihc_dir):
                os.makedirs(os.path.join(work_dir, "特殊患者归档", "免疫组化"), exist_ok=True)
                shutil.copytree(arch_dir, ihc_dir, dirs_exist_ok=True)
        else:
            if os.path.exists(ihc_dir):
                shutil.rmtree(ihc_dir, ignore_errors=True)

        # 完成审核
        approved_dir = os.path.join(work_dir, "特殊患者归档", "完成审核", safe_folder_name)
        if new_review_status == "完成审核":
            if os.path.exists(arch_dir) and not os.path.exists(approved_dir):
                os.makedirs(os.path.join(work_dir, "特殊患者归档", "完成审核"), exist_ok=True)
                shutil.copytree(arch_dir, approved_dir, dirs_exist_ok=True)
        else:
            if os.path.exists(approved_dir):
                shutil.rmtree(approved_dir, ignore_errors=True)

    return {"success": True}

def get_signing_sheet_dir(work_dir):
    """Returns the path to '工作目录/领取表'."""
    p = os.path.join(work_dir, "领取表")
    os.makedirs(p, exist_ok=True)
    return p

def find_patient_archive_dir(work_dir, pathology_no, patient_name):
    """
    Finds the patient's archive directory in '工作目录/按患者归档/'.
    Matches prefix '病理号_姓名'.
    """
    archive_root = get_archive_root_dir(work_dir)
    if not os.path.exists(archive_root):
        return ""
    p_no_clean = sanitize_filename(str(pathology_no).strip())
    p_name_clean = sanitize_filename(str(patient_name).strip())
    prefix = f"{p_no_clean}_{p_name_clean}"

    try:
        entries = os.listdir(archive_root)
    except Exception:
        entries = []

    for entry in entries:
        full = os.path.join(archive_root, entry)
        if os.path.isdir(full) and entry.startswith(prefix):
            return full
    return ""

def revoke_patient_archive_files(work_dir, patient_data):
    """
    Removes the patient's archive folder from 按患者归档 and any 特殊患者归档 folders.
    """
    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    p_no_clean = sanitize_filename(str(p_no).strip())
    p_name_clean = sanitize_filename(str(p_name).strip())
    prefix = f"{p_no_clean}_{p_name_clean}"

    removed = []

    # 1. 按患者归档
    arch_root = get_archive_root_dir(work_dir)
    if os.path.exists(arch_root):
        for entry in os.listdir(arch_root):
            full = os.path.join(arch_root, entry)
            if os.path.isdir(full) and entry.startswith(prefix):
                shutil.rmtree(full, ignore_errors=True)
                removed.append(full)

    # 2. 特殊患者归档 (待审核, 免疫组化, 完成审核)
    for sub in ["待审核", "免疫组化", "完成审核"]:
        spec_root = os.path.join(work_dir, "特殊患者归档", sub)
        if os.path.exists(spec_root):
            for entry in os.listdir(spec_root):
                full = os.path.join(spec_root, entry)
                if os.path.isdir(full) and entry.startswith(prefix):
                    shutil.rmtree(full, ignore_errors=True)
                    removed.append(full)

    return {"success": True, "removed": removed}

def revoke_patient_approval_files(work_dir, patient_data, new_status):
    """
    Copies patient archive from 按患者归档 or 特殊患者归档/完成审核 back to 特殊患者归档 when approval is revoked.
    """
    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    p_no_clean = sanitize_filename(str(p_no).strip())
    p_name_clean = sanitize_filename(str(p_name).strip())
    prefix = f"{p_no_clean}_{p_name_clean}"

    # 1. Locate archive directory (either in 按患者归档 or 特殊患者归档/完成审核)
    arch_dir = find_patient_archive_dir(work_dir, p_no, p_name)
    approved_root = os.path.join(work_dir, "特殊患者归档", "完成审核")
    if not arch_dir or not os.path.exists(arch_dir):
        if os.path.exists(approved_root):
            for entry in os.listdir(approved_root):
                full = os.path.join(approved_root, entry)
                if os.path.isdir(full) and entry.startswith(prefix):
                    arch_dir = full
                    break

    # 2. Remove from 完成审核
    if os.path.exists(approved_root):
        for entry in os.listdir(approved_root):
            full = os.path.join(approved_root, entry)
            if os.path.isdir(full) and entry.startswith(prefix):
                shutil.rmtree(full, ignore_errors=True)

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

    return {"success": True}
