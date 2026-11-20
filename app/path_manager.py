# -*- coding: utf-8 -*-
"""
Dermatology Pathology Reporting System - Path & Directory Management
"""
import hashlib
import os
import re
import shutil
import subprocess
import threading
import uuid
from contextlib import contextmanager

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


class RecycleBinFailedError(RuntimeError):
    """Raised when both send2trash and PowerShell Recycle Bin fail and permanent delete is unconfirmed."""
    def __init__(self, target_path, message):
        super().__init__(message)
        self.target_path = target_path


class ConfirmRequiredError(RuntimeError):
    """Raised when an operation requires explicit user confirmation before overwriting an existing draft."""
    def __init__(self, target_path, code, message):
        super().__init__(message)
        self.target_path = target_path
        self.code = code


_PERM_DELETE_TLS = threading.local()


def is_permanent_delete_allowed():
    return bool(getattr(_PERM_DELETE_TLS, "allow_permanent", False))


@contextmanager
def permanent_delete_context(allowed=True):
    """Temporarily enable or disable explicit permanent-delete fallback in the current thread."""
    prev = getattr(_PERM_DELETE_TLS, "allow_permanent", False)
    _PERM_DELETE_TLS.allow_permanent = bool(allowed)
    try:
        yield
    finally:
        _PERM_DELETE_TLS.allow_permanent = prev


def safe_delete_path(target_path, allow_permanent=None):
    """
    Safely delete a file or directory using Windows Recycle Bin (soft delete).
    When both send2trash and PowerShell Recycle Bin fail, blocks silent permanent deletion
    by raising RecycleBinFailedError unless permanent fallback is explicitly allowed.
    """
    if not target_path or not os.path.exists(target_path):
        return {"success": True}
    norm = os.path.normpath(target_path)
    last_err = ""

    # 1. Primary: send2trash (native Windows shell / IFileOperation)
    try:
        from send2trash import send2trash
        send2trash(norm)
        if not os.path.exists(norm):
            return {"success": True}
    except Exception as exc:
        last_err = str(exc)

    # 2. Secondary: PowerShell Microsoft.VisualBasic FileSystem (argument list, no shell=True)
    if os.path.exists(norm):
        try:
            ps_script = (
                "$p = $env:DERM_TRASH_TARGET; "
                "Add-Type -AssemblyName Microsoft.VisualBasic; "
                "if (Test-Path -LiteralPath $p -PathType Container) { "
                "[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($p, 'OnlyErrorDialogs', 'SendToRecycleBin') "
                "} else { "
                "[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($p, 'OnlyErrorDialogs', 'SendToRecycleBin') "
                "}"
            )
            env = os.environ.copy()
            env["DERM_TRASH_TARGET"] = norm
            res = subprocess.run(
                ["powershell", "-NoProfile", "-Command", ps_script],
                shell=False,
                env=env,
                stdin=subprocess.DEVNULL,
                capture_output=True,
                timeout=20,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            if res.returncode == 0 and not os.path.exists(norm):
                return {"success": True}
            err_bytes = res.stderr or res.stdout or b""
            if err_bytes:
                try:
                    last_err = err_bytes.decode("utf-8", errors="ignore").strip() or last_err
                except Exception:
                    pass
            elif res.returncode != 0:
                last_err = f"PowerShell returncode={res.returncode}"
        except Exception as exc:
            last_err = str(exc)

    if not os.path.exists(norm):
        return {"success": True}

    # 3. Block silent permanent deletion unless explicitly confirmed by the user
    effective_allow = bool(allow_permanent) if allow_permanent is not None else is_permanent_delete_allowed()
    if not effective_allow:
        detail = f"（详情: {last_err}）" if last_err else ""
        raise RecycleBinFailedError(
            norm,
            f"无法将路径移入 Windows 回收站（可能位于网络共享路径、移动介质或权限受限）：{norm}{detail}\n"
            f"为防止数据不可恢复丢失，本次操作已暂停。若确认无需回收站备份，可选择【继续永久删除/直接覆盖】。"
        )

    # User explicitly confirmed permanent deletion fallback
    if os.path.isdir(norm):
        shutil.rmtree(norm, ignore_errors=True)
    else:
        try:
            os.remove(norm)
        except Exception:
            pass
    return {"success": not os.path.exists(norm), "permanent_deleted": True}

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


def _matches_patient_no(entry_name, pathology_no, allow_exact=False, allow_bracket_prefix=False):
    """
    Primary-key matcher: checks whether a filename or folder name belongs to `pathology_no`
    using boundary '_' or '.' (or exact folder match when allow_exact=True).
    Always ignores Word temporary lock files starting with '~$'.
    """
    if not entry_name:
        return False
    name = str(entry_name).strip()
    if name.startswith("~$"):
        return False
    p_no_clean = sanitize_filename(str(pathology_no or "").strip())
    if not p_no_clean:
        return False
    if allow_bracket_prefix and name.startswith("[病理报告]"):
        name = name[len("[病理报告]"):]
    if name.startswith(f"{p_no_clean}_") or name.startswith(f"{p_no_clean}."):
        return True
    if allow_exact and name == p_no_clean:
        return True
    return False


def _matches_any_patient_no(entry_name, pathology_nos, allow_exact=False, allow_bracket_prefix=False):
    """Checks whether entry_name matches any pathology_no in `pathology_nos`."""
    for pno in (pathology_nos or []):
        if pno and _matches_patient_no(
            entry_name,
            pno,
            allow_exact=allow_exact,
            allow_bracket_prefix=allow_bracket_prefix,
        ):
            return True
    return False


def _collect_patient_no_candidates(pathology_no, extra_old_nos=None, extra_old_paths=None):
    """Build ordered unique list of candidate pathology numbers (current + old/renamed)."""
    cands = []
    for raw in [pathology_no] + list(extra_old_nos or []):
        s = str(raw or "").strip()
        if s and s not in cands:
            cands.append(s)
    for op in (extra_old_paths or []):
        if not op:
            continue
        base = os.path.basename(str(op).strip())
        if base.startswith("[病理报告]"):
            base = base[len("[病理报告]"):]
        stem = os.path.splitext(base)[0]
        if "_" in stem:
            no_part = stem.split("_", 1)[0].strip()
            if no_part and no_part not in cands:
                cands.append(no_part)
    return cands


def sanitize_filename(name):
    """Remove characters forbidden in Windows filenames (including newlines/tabs) except valid dots/semicolons."""
    # Convert line breaks / tabs in multiline diagnoses into semicolons before scrubbing Windows-forbidden chars
    s = re.sub(r'[\r\n]+', '；', str(name or '').strip())
    s = re.sub(r'[\t\x00-\x1f]+', ' ', s)
    # Forbidden: \ / : * ? " < > | '
    s = re.sub(r'[\\/:*?"<>|\']', '_', s)
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

def _paths_same_entry(path_a, path_b):
    """Return True if two paths resolve to the exact same filesystem entry (case-insensitive on Windows)."""
    if not path_a or not path_b:
        return False
    return os.path.normcase(os.path.abspath(path_a)) == os.path.normcase(os.path.abspath(path_b))


def _ensure_writable(file_path, lock_label="报告文件"):
    """Verify target file is not locked by Word/Excel/WPS before overwriting or deleting."""
    if file_path and os.path.exists(file_path) and os.path.isfile(file_path):
        try:
            with open(file_path, "a+b"):
                pass
        except (PermissionError, IOError) as pe:
            app_hint = "Excel 或 WPS" if "Excel" in str(lock_label) else "Word 或 WPS"
            raise PermissionError(
                f"{lock_label}被其他软件（如 {app_hint}）打开占用，无法覆盖！请先关闭该文件后重试。路径: {os.path.basename(file_path)}"
            ) from pe


def _ensure_dir_docs_writable(dir_path):
    """Verify all Word/Excel documents inside dir_path are unlocked before moving/deleting directory."""
    if not dir_path or not os.path.isdir(dir_path):
        return
    for root_p, _, files in os.walk(dir_path):
        for fn in files:
            if fn.lower().endswith((".docx", ".doc", ".xlsx", ".xls")):
                _ensure_writable(os.path.join(root_p, fn))


def backup_before_overwrite(target_path, allow_permanent=None, lock_label="报告文件"):
    """
    Protect an existing file before overwriting it:
    1) Check _ensure_writable(target_path) to ensure not locked by Word/Excel/WPS.
    2) Move old version into Windows Recycle Bin via safe_delete_path(target_path).
    """
    if not target_path or not os.path.exists(target_path):
        return False
    if os.path.isdir(target_path):
        raise IsADirectoryError(f"目标路径是文件夹而非文件，无法直接覆写: {target_path}")
    _ensure_writable(target_path, lock_label=lock_label)
    res = safe_delete_path(target_path, allow_permanent=allow_permanent)
    if not res.get("success") or os.path.exists(target_path):
        raise OSError(f"无法将旧版文件移入回收站，覆写已中止: {os.path.basename(target_path)}")
    return True


def _move_path_safe(src_path, dst_path, is_dir=False, lock_label="报告文件"):
    """
    Move file or directory from src_path to dst_path safely without silent exception swallowing:
    - Verifies source files are unlocked.
    - Supports Windows case-only renames via temporary path.
    - If distinct destination already exists, recycles conflicting destination first before moving.
    """
    if not src_path or not os.path.exists(src_path) or src_path == dst_path:
        return dst_path

    if is_dir:
        _ensure_dir_docs_writable(src_path)
    else:
        _ensure_writable(src_path, lock_label=lock_label)

    dst_parent = os.path.dirname(dst_path)
    if dst_parent:
        os.makedirs(dst_parent, exist_ok=True)

    if _paths_same_entry(src_path, dst_path):
        tmp_path = f"{src_path}.__rename_{uuid.uuid4().hex[:8]}"
        shutil.move(src_path, tmp_path)
        shutil.move(tmp_path, dst_path)
        return dst_path

    if os.path.exists(dst_path):
        if is_dir:
            _ensure_dir_docs_writable(dst_path)
            res = safe_delete_path(dst_path)
            if not res.get("success") or os.path.exists(dst_path):
                raise FileExistsError(f"重命名冲突：目标文件夹已存在且无法清理至回收站: {dst_path}")
        else:
            backup_before_overwrite(dst_path, lock_label=lock_label)

    shutil.move(src_path, dst_path)
    return dst_path


def rename_patient_files_and_folders(work_dir, old_no, old_name, new_no, new_name):
    """
    Renames Word report file, output report file, photo folder, and special archive folders
    if pathology number or patient name changed.
    Does NOT touch 已归档报告/ (archived reports are only updated on explicit re-archive).
    Boundary 3: NEVER touch 未归档报告/临床照片/ (identified strictly by patient name folder).
    Recycles conflicting pre-existing destinations and propagates lock/conflict errors loudly.
    """
    old_word = get_patient_word_path(work_dir, old_no, old_name)
    new_word = get_patient_word_path(work_dir, new_no, new_name)

    old_photo = get_patient_photo_dir(work_dir, old_no, old_name)
    new_photo = get_patient_photo_dir(work_dir, new_no, new_name)

    old_no_clean = sanitize_filename(old_no)
    old_name_clean = sanitize_filename(old_name)
    new_no_clean = sanitize_filename(new_no)
    new_name_clean = sanitize_filename(new_name)
    old_prefix_base = f"{old_no_clean}_{old_name_clean}"
    new_prefix_base = f"{new_no_clean}_{new_name_clean}"
    old_prefix = f"{old_prefix_base}_"
    new_prefix = f"{new_prefix_base}_"

    renamed_pairs = []

    # 1. Rename draft Word file (病理报告录入)
    if old_word != new_word and os.path.exists(old_word):
        _move_path_safe(old_word, new_word, is_dir=False)
        renamed_pairs.append((old_word, new_word))
    elif old_prefix_base != new_prefix_base and os.path.exists(get_report_dir(work_dir)):
        rep_dir = get_report_dir(work_dir)
        for fn in sorted(os.listdir(rep_dir)):
            if fn.lower().endswith(".docx") and _matches_patient_no(fn, old_no):
                cand = os.path.join(rep_dir, fn)
                if not _paths_same_entry(cand, new_word) and os.path.isfile(cand):
                    _move_path_safe(cand, new_word, is_dir=False)
                    renamed_pairs.append((cand, new_word))
                break

    # 2. Rename final output report file if exists
    new_final_path = ""
    out_dir = get_output_report_dir(work_dir)
    if os.path.exists(out_dir):
        for f in sorted(os.listdir(out_dir)):
            if not f.startswith("~$") and f.lower().endswith(".docx") and (
                f.startswith(old_prefix) or _matches_patient_no(f, old_no)
            ):
                old_f = os.path.join(out_dir, f)
                if f.startswith(old_prefix):
                    diag_part = f[len(old_prefix):]
                else:
                    parts = os.path.splitext(f)[0].split("_", 2)
                    diag_part = f"{parts[2]}.docx" if len(parts) >= 3 else f"{os.path.splitext(f)[0]}.docx"
                new_f = os.path.join(out_dir, f"{new_prefix}{diag_part}")
                if old_f != new_f:
                    _move_path_safe(old_f, new_f, is_dir=False)
                    renamed_pairs.append((old_f, new_f))
                new_final_path = new_f if os.path.exists(new_f) else old_f
                break

    # 3. Rename special archive folders & internal report docx (excluding 已归档报告/ & 临床照片/)
    renamed_ihc = {}
    if old_prefix_base != new_prefix_base:
        for sub in ["待审核", "免疫组化", "重切", "完成审核"]:
            spec_root = os.path.join(work_dir, "特殊患者归档", sub)
            if not os.path.exists(spec_root):
                continue
            for entry in list(os.listdir(spec_root)):
                if (
                    entry.startswith(old_prefix)
                    or entry == old_prefix_base
                    or _matches_patient_no(entry, old_no, allow_exact=True)
                ):
                    old_dir = os.path.join(spec_root, entry)
                    if not os.path.isdir(old_dir):
                        continue
                    if entry.startswith(old_prefix_base):
                        suffix = entry[len(old_prefix_base):]
                    else:
                        parts = entry.split("_", 2)
                        suffix = f"_{parts[2]}" if len(parts) >= 3 else ""
                    new_entry = f"{new_prefix_base}{suffix}"
                    new_dir = os.path.join(spec_root, new_entry)
                    if old_dir != new_dir:
                        _move_path_safe(old_dir, new_dir, is_dir=True)
                        renamed_pairs.append((old_dir, new_dir))
                    if os.path.isdir(new_dir):
                        old_doc = os.path.join(new_dir, f"[病理报告]{entry}.docx")
                        new_doc = os.path.join(new_dir, f"[病理报告]{new_entry}.docx")
                        if os.path.exists(old_doc) and old_doc != new_doc:
                            _move_path_safe(old_doc, new_doc, is_dir=False)
                            renamed_pairs.append((os.path.join(old_dir, f"[病理报告]{entry}.docx"), new_doc))
                            renamed_pairs.append((old_doc, new_doc))
                        if sub == "免疫组化":
                            renamed_ihc = {
                                "old_dir": old_dir,
                                "ihc_dir": new_dir,
                                "ihc_doc_path": new_doc if os.path.exists(new_doc) else old_doc,
                                "ihc_photo_dir": os.path.join(new_dir, "病理照片")
                            }

    # 4. Rename pathology photo folder (Boundary 3: never touch 临床照片/)
    if old_photo != new_photo and os.path.exists(old_photo):
        if _paths_same_entry(old_photo, new_photo) or not os.path.exists(new_photo):
            _move_path_safe(old_photo, new_photo, is_dir=True)
        else:
            for item in list(os.listdir(old_photo)):
                s = os.path.join(old_photo, item)
                d = os.path.join(new_photo, item)
                _move_path_safe(s, d, is_dir=os.path.isdir(s))
            os.rmdir(old_photo)

    ensure_patient_photo_dir(work_dir, new_no, new_name)
    return {
        "word_path": new_word,
        "photo_dir": new_photo,
        "old_photo_dir": old_photo,
        "final_report_path": new_final_path,
        "renamed_ihc": renamed_ihc,
        "renamed_pairs": renamed_pairs,
    }


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


def cleanup_stale_output_reports(work_dir, pathology_no, patient_name="", keep_path="", extra_old_nos=None):
    """Remove obsolete diagnosis/name filenames in 未归档报告/病理报告输出/ for this patient (primary-key matched)."""
    out_dir = get_output_report_dir(work_dir)
    if not os.path.exists(out_dir):
        return []
    no_cands = _collect_patient_no_candidates(pathology_no, extra_old_nos=extra_old_nos)
    keep_norm = os.path.normcase(os.path.normpath(keep_path)) if keep_path else ""
    removed = []
    for f in os.listdir(out_dir):
        if not f.startswith("~$") and f.lower().endswith(".docx") and _matches_any_patient_no(f, no_cands):
            fp = os.path.join(out_dir, f)
            if keep_norm and os.path.normcase(os.path.normpath(fp)) == keep_norm:
                continue
            safe_delete_path(fp)
            removed.append(fp)
    return removed


def _find_special_cat_folder(work_dir, cat_name, pathology_no, patient_name=""):
    """Find an existing subfolder for patient under 工作目录/特殊患者归档/<cat_name>/ using primary-key pathology_no."""
    p_no_clean = sanitize_filename(str(pathology_no or "").strip())
    p_name_clean = sanitize_filename(str(patient_name or "").strip())
    prefix_base = f"{p_no_clean}_{p_name_clean}" if p_name_clean else ""
    cat_root = os.path.join(work_dir, "特殊患者归档", cat_name)
    if os.path.exists(cat_root):
        try:
            entries = sorted(os.listdir(cat_root))
            if prefix_base:
                for entry in entries:
                    full = os.path.join(cat_root, entry)
                    if os.path.isdir(full) and (entry.startswith(f"{prefix_base}_") or entry == prefix_base):
                        return full
            for entry in entries:
                full = os.path.join(cat_root, entry)
                if os.path.isdir(full) and _matches_patient_no(entry, pathology_no, allow_exact=True):
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

    # Migrate photos and remove any stale old-diagnosis/old-name folder in this category
    for entry in list(os.listdir(cat_root)):
        if entry != safe_folder_name and _matches_patient_no(entry, p_no, allow_exact=True):
            old_dir = os.path.join(cat_root, entry)
            if os.path.isdir(old_dir):
                _ensure_dir_docs_writable(old_dir)
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
        if not _paths_same_entry(report_file_path, target_doc_path):
            backup_before_overwrite(target_doc_path)
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

def archive_patient_report(work_dir, patient_data, report_file_path, is_from_review=False, extra_old_paths=None, extra_old_nos=None):
    """
    Archives a patient's original report:
      1. Regular archives (including IHC-only patients):
         Stored directly into: 工作目录/已归档报告/病理号_姓名_诊断.docx (Flat file, NO subfolders, NO '[病理报告]' prefix!)
         Photos are NOT copied to regular archive (solves file duplication / bloat).
         Maintains primary-key uniqueness per pathology_no (recycles stale renamed/re-diagnosed non-IHC archives).
         Does NOT overwrite or delete *_免疫组化报告.docx or 特殊患者归档/免疫组化/.
      2. Review special patients archives (待审核 / 重切):
         Moves package (pathology photos + [病理报告]...docx, NO clinical photos) from 待审核 / 重切 into 特殊患者归档/完成审核/.
    """
    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    p_diag = patient_data.get("pathology_diag") or patient_data.get("clinical_diagnosis") or "已诊断"
    no_cands = _collect_patient_no_candidates(p_no, extra_old_nos=extra_old_nos, extra_old_paths=extra_old_paths)

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

    # Clean up any legacy, renamed, or duplicate non-IHC archived files for this patient in 已归档报告/
    stale_candidates = []
    if os.path.exists(archive_root):
        for item in sorted(os.listdir(archive_root)):
            if (
                not item.startswith("~$")
                and _matches_any_patient_no(item, no_cands, allow_bracket_prefix=True)
                and item.lower().endswith(".docx")
                and not is_ihc_archived_filename(item)
            ):
                old_f = os.path.join(archive_root, item)
                if not _paths_same_entry(old_f, target_archive_file):
                    stale_candidates.append(old_f)

    for op in (extra_old_paths or []):
        if (
            op
            and os.path.isfile(op)
            and not is_ihc_archived_filename(op)
            and not _paths_same_entry(op, target_archive_file)
            and all(not _paths_same_entry(op, existing) for existing in stale_candidates)
        ):
            stale_candidates.append(op)

    for old_f in stale_candidates:
        _ensure_writable(old_f, lock_label="旧版已归档报告文件")
        safe_delete_path(old_f)

    # Copy final report to 已归档报告/ (recycling any existing file at target first)
    if report_file_path and os.path.exists(report_file_path):
        if not _paths_same_entry(report_file_path, target_archive_file):
            backup_before_overwrite(target_archive_file)
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
                    if _matches_any_patient_no(e, no_cands, allow_exact=True):
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

def find_patient_ihc_folder(work_dir, pathology_no, patient_name=""):
    """Finds existing folder under 特殊患者归档/免疫组化/ for this patient (primary-key matched)."""
    p_no_clean = sanitize_filename(str(pathology_no).strip())
    p_name_clean = sanitize_filename(str(patient_name or "").strip())
    prefix = f"{p_no_clean}_{p_name_clean}" if p_name_clean else ""
    ihc_root = os.path.join(work_dir, "特殊患者归档", "免疫组化")
    if os.path.exists(ihc_root):
        try:
            entries = sorted(os.listdir(ihc_root))
            if prefix:
                for entry in entries:
                    full = os.path.join(ihc_root, entry)
                    if os.path.isdir(full) and (entry.startswith(f"{prefix}_") or entry == prefix):
                        return full
            for entry in entries:
                full = os.path.join(ihc_root, entry)
                if os.path.isdir(full) and _matches_patient_no(entry, pathology_no, allow_exact=True):
                    return full
        except Exception:
            pass
    return ""


def find_patient_ihc_archive_path(work_dir, pathology_no, patient_name="", extra_archive_path="", extra_old_nos=None):
    """Finds the patient's archived IHC report (*_免疫组化报告.docx) in 工作目录/已归档报告/."""
    p_no_clean = sanitize_filename(str(pathology_no).strip())
    p_name_clean = sanitize_filename(str(patient_name or "").strip())
    prefix = f"{p_no_clean}_{p_name_clean}" if p_name_clean else ""
    no_cands = _collect_patient_no_candidates(pathology_no, extra_old_nos=extra_old_nos)
    archive_root = get_archive_root_dir(work_dir)
    if os.path.exists(archive_root):
        try:
            entries = sorted(os.listdir(archive_root))
            if prefix:
                for entry in entries:
                    if (entry.startswith(f"{prefix}_") or entry.startswith(f"{prefix}.")) and is_ihc_archived_filename(entry) and not entry.startswith("~$"):
                        return os.path.join(archive_root, entry)
            for entry in entries:
                if _matches_any_patient_no(entry, no_cands, allow_bracket_prefix=True) and is_ihc_archived_filename(entry):
                    return os.path.join(archive_root, entry)
        except Exception:
            pass
    if extra_archive_path and os.path.isfile(extra_archive_path) and is_ihc_archived_filename(extra_archive_path):
        return extra_archive_path
    return ""


def create_ihc_report_package(work_dir, patient_data, source_archive_file, force_overwrite=False):
    """
    Create an independent IHC report package in 工作目录/特殊患者归档/免疫组化/<folder>/
    by copying the patient's archived report and pathology photo directory (WITHOUT clinical photos).
    Refuses to silently overwrite an existing IHC draft unless force_overwrite=True.
    """
    if not source_archive_file or not os.path.exists(source_archive_file) or not os.path.isfile(source_archive_file):
        raise FileNotFoundError("未找到患者的已归档原报告文件，请先将原报告归档后再创建免疫组化报告")

    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    source_filename = os.path.basename(source_archive_file)
    source_stem = os.path.splitext(source_filename)[0]

    ihc_root = os.path.join(work_dir, "特殊患者归档", "免疫组化")
    ihc_dir = os.path.join(ihc_root, source_stem)
    ihc_photo_dir = os.path.join(ihc_dir, "病理照片")
    ihc_doc_path = os.path.join(ihc_dir, f"[病理报告]{source_stem}.docx")
    prefix = f"{sanitize_filename(p_no)}_{sanitize_filename(p_name)}"

    # Detect any existing IHC draft document BEFORE mutating files/folders
    existing_drafts = []
    if os.path.exists(ihc_doc_path):
        existing_drafts.append(ihc_doc_path)
    if os.path.exists(ihc_root):
        for entry in sorted(os.listdir(ihc_root)):
            if entry != source_stem and _matches_patient_no(entry, p_no, allow_exact=True):
                cand_dir = os.path.join(ihc_root, entry)
                if os.path.isdir(cand_dir):
                    for fn in os.listdir(cand_dir):
                        if not fn.startswith("~$") and fn.lower().endswith(".docx"):
                            existing_drafts.append(os.path.join(cand_dir, fn))

    if existing_drafts and not force_overwrite:
        raise ConfirmRequiredError(
            existing_drafts[0],
            "IHC_DRAFT_EXISTS",
            f"该患者已存在免疫组化报告草稿（{os.path.basename(existing_drafts[0])}），继续创建将使用当前【已归档报告】替换已有草稿（旧草稿将移入回收站，已有照片保留）。确定继续覆盖吗？"
        )

    os.makedirs(ihc_root, exist_ok=True)
    os.makedirs(ihc_dir, exist_ok=True)
    os.makedirs(ihc_photo_dir, exist_ok=True)

    # Preserve any photos if an older-named IHC folder existed for this patient
    for entry in list(os.listdir(ihc_root)):
        if entry != source_stem and _matches_patient_no(entry, p_no, allow_exact=True):
            old_dir = os.path.join(ihc_root, entry)
            if os.path.isdir(old_dir):
                _ensure_dir_docs_writable(old_dir)
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

    # Recycle existing IHC report docx first before overwriting
    if not _paths_same_entry(source_archive_file, ihc_doc_path):
        backup_before_overwrite(ihc_doc_path, lock_label="免疫组化报告文件")
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


def archive_ihc_report(work_dir, patient_data, ihc_data, extra_old_paths=None, extra_old_nos=None):
    """
    Archive the saved IHC report document (document only, no photos, no move to 完成审核)
    into 工作目录/已归档报告/<原已归档文件名>_免疫组化报告.docx without touching the original report.
    Ensures primary-key uniqueness and removes any stale/renamed *_免疫组化报告.docx for this patient.
    """
    ihc_doc_path = (ihc_data or {}).get("ihc_doc_path", "")
    if not ihc_doc_path or not os.path.exists(ihc_doc_path):
        return {"success": False, "error": "未找到已保存的免疫组化报告文件，请先创建并保存免疫组化报告"}

    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    p_no_clean = sanitize_filename(p_no)
    p_name_clean = sanitize_filename(p_name)
    expected_prefix = f"{p_no_clean}_{p_name_clean}_"
    no_cands = _collect_patient_no_candidates(
        p_no,
        extra_old_nos=extra_old_nos,
        extra_old_paths=list(extra_old_paths or []) + [(ihc_data or {}).get("archive_file", "")],
    )

    orig_arch = find_patient_archive_path(
        work_dir,
        p_no,
        p_name,
        extra_archive_path=patient_data.get("archive_dir", ""),
        extra_old_nos=extra_old_nos,
    )
    src_name = (ihc_data or {}).get("source_archive_filename", "")
    if orig_arch and os.path.basename(orig_arch).startswith(expected_prefix):
        src_name = os.path.basename(orig_arch)

    base_stem = ""
    if src_name:
        cand_stem = os.path.splitext(src_name)[0]
        if cand_stem.endswith("_免疫组化报告"):
            cand_stem = cand_stem[:-len("_免疫组化报告")]
        if cand_stem.startswith(expected_prefix):
            base_stem = cand_stem
        else:
            parts = cand_stem.split("_", 2)
            diag_tail = parts[2] if len(parts) >= 3 and parts[2] else ""
            if diag_tail:
                base_stem = f"{p_no_clean}_{p_name_clean}_{diag_tail}"

    if not base_stem:
        p_diag = (ihc_data or {}).get("pathology_diag") or patient_data.get("pathology_diag") or patient_data.get("clinical_diagnosis") or "已诊断"
        base_stem = f"{p_no_clean}_{p_name_clean}_{sanitize_filename(p_diag)}"

    if base_stem.endswith("_免疫组化报告"):
        base_stem = base_stem[:-len("_免疫组化报告")]

    target_filename = f"{base_stem}{IHC_ARCHIVE_SUFFIX}"
    archive_root = get_archive_root_dir(work_dir)
    target_archive_file = os.path.join(archive_root, target_filename)
    _ensure_writable(target_archive_file, lock_label="已归档免疫组化报告文件")

    # Remove only stale *_免疫组化报告.docx for this patient, NEVER touch the original archived report
    stale_candidates = []
    if os.path.exists(archive_root):
        for item in sorted(os.listdir(archive_root)):
            if (
                not item.startswith("~$")
                and _matches_any_patient_no(item, no_cands, allow_bracket_prefix=True)
                and is_ihc_archived_filename(item)
            ):
                old_f = os.path.join(archive_root, item)
                if not _paths_same_entry(old_f, target_archive_file):
                    stale_candidates.append(old_f)

    for op in list(extra_old_paths or []) + [(ihc_data or {}).get("archive_file", "")]:
        if (
            op
            and os.path.isfile(op)
            and is_ihc_archived_filename(op)
            and not _paths_same_entry(op, target_archive_file)
            and all(not _paths_same_entry(op, existing) for existing in stale_candidates)
        ):
            stale_candidates.append(op)

    for old_f in stale_candidates:
        _ensure_writable(old_f, lock_label="旧版已归档免疫组化报告文件")
        safe_delete_path(old_f)

    if not _paths_same_entry(ihc_doc_path, target_archive_file):
        backup_before_overwrite(target_archive_file, lock_label="已归档免疫组化报告文件")
        shutil.copy2(ihc_doc_path, target_archive_file)
    return {
        "success": True,
        "archive_file": target_archive_file,
        "filename": target_filename,
        "source_archive_filename": f"{base_stem}.docx",
        "source_archive_path": orig_arch or "",
    }


def revoke_ihc_archive_files(work_dir, patient_data, ihc_data=None):
    """
    Removes ONLY the archived IHC report (*_免疫组化报告.docx) from 已归档报告/,
    preserving the original archived report and 特殊患者归档/免疫组化/.
    """
    p_no = patient_data.get("pathology_no", "")
    arch_file = (ihc_data or {}).get("archive_file", "")
    no_cands = _collect_patient_no_candidates(p_no, extra_old_paths=[arch_file])
    removed = []

    if arch_file and os.path.exists(arch_file) and is_ihc_archived_filename(arch_file):
        _ensure_writable(arch_file, lock_label="已归档免疫组化报告文件")
        safe_delete_path(arch_file)
        removed.append(arch_file)

    archive_root = get_archive_root_dir(work_dir)
    if os.path.exists(archive_root):
        for entry in list(os.listdir(archive_root)):
            if _matches_any_patient_no(entry, no_cands, allow_bracket_prefix=True) and is_ihc_archived_filename(entry):
                full = os.path.join(archive_root, entry)
                if os.path.exists(full):
                    _ensure_writable(full, lock_label="已归档免疫组化报告文件")
                    safe_delete_path(full)
                    removed.append(full)

    return {"success": True, "removed": removed}


def delete_ihc_report_files(work_dir, patient_data, ihc_data=None):
    """
    Safely moves ALL IHC files & folders for a patient to Recycle Bin via safe_delete_path:
      1. 特殊患者归档/免疫组化/<patient_folder>
      2. 已归档报告/<pathology_no>_*_免疫组化报告.docx
    Does NOT touch the patient's regular original report, draft, or photos.
    """
    p_no = patient_data.get("pathology_no", "")
    ihc_dir_recorded = (ihc_data or {}).get("ihc_dir", "")
    ihc_doc_recorded = (ihc_data or {}).get("ihc_doc_path", "")
    ihc_arch_recorded = (ihc_data or {}).get("archive_file", "")
    no_cands = _collect_patient_no_candidates(p_no, extra_old_paths=[ihc_arch_recorded, ihc_dir_recorded])

    # Pre-check locks before deleting anything
    if ihc_doc_recorded and os.path.exists(ihc_doc_recorded):
        _ensure_writable(ihc_doc_recorded)
    if ihc_arch_recorded and os.path.exists(ihc_arch_recorded):
        _ensure_writable(ihc_arch_recorded)

    archive_root = get_archive_root_dir(work_dir)
    if os.path.exists(archive_root):
        for entry in os.listdir(archive_root):
            if _matches_any_patient_no(entry, no_cands, allow_bracket_prefix=True) and is_ihc_archived_filename(entry):
                _ensure_writable(os.path.join(archive_root, entry))

    ihc_root = os.path.join(work_dir, "特殊患者归档", "免疫组化")
    ihc_dirs_to_delete = []
    if ihc_dir_recorded and os.path.exists(ihc_dir_recorded):
        ihc_dirs_to_delete.append(os.path.normpath(ihc_dir_recorded))
    if os.path.exists(ihc_root):
        for entry in os.listdir(ihc_root):
            full = os.path.normpath(os.path.join(ihc_root, entry))
            if _matches_any_patient_no(entry, no_cands, allow_exact=True) and os.path.isdir(full):
                if all(not _paths_same_entry(full, existing) for existing in ihc_dirs_to_delete):
                    ihc_dirs_to_delete.append(full)

    for d in ihc_dirs_to_delete:
        for root_p, _, files in os.walk(d):
            for fn in files:
                if fn.lower().endswith((".docx", ".doc")):
                    _ensure_writable(os.path.join(root_p, fn))

    removed = []
    rev_arch = revoke_ihc_archive_files(work_dir, patient_data, ihc_data=ihc_data)
    removed.extend(rev_arch.get("removed", []))

    for d in ihc_dirs_to_delete:
        if os.path.exists(d):
            safe_delete_path(d)
            removed.append(d)

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

def delete_patient_all_files(work_dir, patient_data, ihc_data=None):
    """
    Deletes all files and directories related to the patient into Windows Recycle Bin:
    - Draft word report
    - Output word report
    - Pathology photo folder
    - Archived report file in 已归档报告 (including stale pre-rename archive paths)
    - Special archive folders
    Never touches 未归档报告/临床照片/.
    """
    deleted_items = []
    errors = []

    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    recorded_arch = patient_data.get("archive_dir", "")
    recorded_ihc_arch = (ihc_data or {}).get("archive_file", "")
    recorded_final = patient_data.get("final_report_path", "")
    no_cands = _collect_patient_no_candidates(
        p_no,
        extra_old_paths=[recorded_arch, recorded_ihc_arch, recorded_final],
    )

    # 1. Draft Word
    draft_p = patient_data.get("word_path") or get_patient_word_path(work_dir, p_no, p_name)
    if draft_p and os.path.exists(draft_p):
        try:
            safe_delete_path(draft_p)
            deleted_items.append(draft_p)
        except RecycleBinFailedError:
            raise
        except Exception as e:
            errors.append(f"删除草稿文件失败: {e}")

    # 2. Output Word
    out_dir = get_output_report_dir(work_dir)
    if os.path.exists(out_dir):
        for f in list(os.listdir(out_dir)):
            if not f.startswith("~$") and _matches_any_patient_no(f, no_cands) and f.lower().endswith(".docx"):
                fp = os.path.join(out_dir, f)
                try:
                    safe_delete_path(fp)
                    deleted_items.append(fp)
                except RecycleBinFailedError:
                    raise
                except Exception as e:
                    errors.append(f"删除输出报告失败: {e}")

    # 3. Pathology photos folder
    photo_p = patient_data.get("photo_dir") or get_patient_photo_dir(work_dir, p_no, p_name)
    if photo_p and os.path.exists(photo_p):
        try:
            safe_delete_path(photo_p)
            deleted_items.append(photo_p)
        except RecycleBinFailedError:
            raise
        except Exception as e:
            errors.append(f"删除病理照片目录失败: {e}")

    # 4. Archived reports in 已归档报告/ (and legacy 按患者归档/), plus stale recorded paths
    seen_arch = set()
    for extra_ap in (recorded_arch, recorded_ihc_arch):
        if extra_ap and os.path.exists(extra_ap):
            np_ap = os.path.normcase(os.path.normpath(extra_ap))
            if np_ap not in seen_arch:
                seen_arch.add(np_ap)
                try:
                    safe_delete_path(extra_ap)
                    deleted_items.append(extra_ap)
                except RecycleBinFailedError:
                    raise
                except Exception as e:
                    errors.append(f"删除归档报告失败: {e}")

    for arch_name in ["已归档报告", "按患者归档"]:
        arch_root = os.path.join(work_dir, arch_name)
        if os.path.exists(arch_root):
            for entry in list(os.listdir(arch_root)):
                if _matches_any_patient_no(entry, no_cands, allow_bracket_prefix=True):
                    target = os.path.join(arch_root, entry)
                    np_t = os.path.normcase(os.path.normpath(target))
                    if np_t in seen_arch or not os.path.exists(target):
                        continue
                    seen_arch.add(np_t)
                    try:
                        safe_delete_path(target)
                        deleted_items.append(target)
                    except RecycleBinFailedError:
                        raise
                    except Exception as e:
                        errors.append(f"删除归档报告失败: {e}")

    # 5. Special archives
    for sub in ["待审核", "免疫组化", "重切", "完成审核"]:
        spec_root = os.path.join(work_dir, "特殊患者归档", sub)
        if os.path.exists(spec_root):
            for entry in list(os.listdir(spec_root)):
                if _matches_any_patient_no(entry, no_cands, allow_exact=True):
                    target = os.path.join(spec_root, entry)
                    try:
                        safe_delete_path(target)
                        deleted_items.append(target)
                    except RecycleBinFailedError:
                        raise
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

def find_patient_archive_path(work_dir, pathology_no, patient_name="", extra_archive_path="", extra_old_nos=None):
    """
    Finds the patient's original (non-IHC) archived report in '工作目录/已归档报告/'.
    Uses primary-key matching on pathology_no (preferring current pathology_no_name prefix first)
    and excludes '*_免疫组化报告.docx'. Also falls back to extra_archive_path / legacy 按患者归档/.
    """
    p_no_clean = sanitize_filename(str(pathology_no or "").strip())
    p_name_clean = sanitize_filename(str(patient_name or "").strip())
    prefix = f"{p_no_clean}_{p_name_clean}" if p_name_clean else ""
    no_cands = _collect_patient_no_candidates(pathology_no, extra_old_nos=extra_old_nos, extra_old_paths=[extra_archive_path])

    for root_dir in (get_archive_root_dir(work_dir), os.path.join(work_dir, "按患者归档")):
        if not os.path.exists(root_dir):
            continue
        try:
            entries = sorted(os.listdir(root_dir))
            if prefix:
                for entry in entries:
                    if (entry.startswith(f"{prefix}_") or entry.startswith(f"{prefix}.")) and not is_ihc_archived_filename(entry) and not entry.startswith("~$"):
                        return os.path.join(root_dir, entry)
            for entry in entries:
                if _matches_any_patient_no(entry, no_cands, allow_bracket_prefix=True) and not is_ihc_archived_filename(entry):
                    return os.path.join(root_dir, entry)
        except Exception:
            pass

    if extra_archive_path and os.path.isfile(extra_archive_path) and not is_ihc_archived_filename(extra_archive_path):
        return extra_archive_path

    return ""

def find_patient_archive_dir(work_dir, pathology_no, patient_name="", extra_archive_path="", extra_old_nos=None):
    """Alias for backwards compatibility."""
    return find_patient_archive_path(
        work_dir,
        pathology_no,
        patient_name=patient_name,
        extra_archive_path=extra_archive_path,
        extra_old_nos=extra_old_nos,
    )

def revoke_patient_archive_files(work_dir, patient_data):
    """
    Removes the patient's original archived report from 已归档报告/ and review 特殊患者归档 folders.
    Preserves *_免疫组化报告.docx and 特殊患者归档/免疫组化/.
    """
    p_no = patient_data.get("pathology_no", "")
    recorded_arch = patient_data.get("archive_dir", "")
    no_cands = _collect_patient_no_candidates(p_no, extra_old_paths=[recorded_arch])

    removed = []
    seen = set()

    if recorded_arch and os.path.exists(recorded_arch) and not is_ihc_archived_filename(recorded_arch):
        _ensure_writable(recorded_arch, lock_label="已归档报告文件")
        safe_delete_path(recorded_arch)
        removed.append(recorded_arch)
        seen.add(os.path.normcase(os.path.normpath(recorded_arch)))

    # 1. 已归档报告 (and legacy 按患者归档), excluding *_免疫组化报告.docx
    for arch_name in ["已归档报告", "按患者归档"]:
        arch_root = os.path.join(work_dir, arch_name)
        if os.path.exists(arch_root):
            for entry in list(os.listdir(arch_root)):
                full = os.path.join(arch_root, entry)
                np_full = os.path.normcase(os.path.normpath(full))
                if np_full in seen:
                    continue
                if _matches_any_patient_no(entry, no_cands, allow_bracket_prefix=True) and not is_ihc_archived_filename(entry):
                    _ensure_writable(full, lock_label="已归档报告文件")
                    safe_delete_path(full)
                    removed.append(full)
                    seen.add(np_full)

    # 2. 特殊患者归档 (待审核, 重切, 完成审核; excluding 免疫组化)
    for sub in ["待审核", "重切", "完成审核"]:
        spec_root = os.path.join(work_dir, "特殊患者归档", sub)
        if os.path.exists(spec_root):
            for entry in list(os.listdir(spec_root)):
                full = os.path.join(spec_root, entry)
                if _matches_any_patient_no(entry, no_cands, allow_exact=True):
                    _ensure_dir_docs_writable(full)
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
    arch_dir = _find_special_cat_folder(work_dir, "完成审核", p_no, p_name)

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


# =========================================================================
# External Word Modification Detection (SHA-256 Fingerprints & 3-Tier Sync)
# =========================================================================

def normalize_sync_path(path):
    """Return case-folded normalized absolute path key for Windows sync tracking."""
    if not path:
        return ""
    return os.path.normcase(os.path.normpath(os.path.abspath(str(path))))


def compute_file_fingerprint(file_path):
    """Compute SHA-256 hash + stat metadata for a regular file, or None if missing."""
    if not file_path or not os.path.isfile(file_path):
        return None
    h = hashlib.sha256()
    with open(file_path, "rb") as f:
        while True:
            chunk = f.read(65536)
            if not chunk:
                break
            h.update(chunk)
    st = os.stat(file_path)
    return {
        "path": os.path.normpath(os.path.abspath(file_path)),
        "sha256": h.hexdigest(),
        "size": int(st.st_size),
        "mtime_ns": int(getattr(st, "st_mtime_ns", int(st.st_mtime * 1e9))),
    }


def _find_docx_inside_dir(dir_path):
    """Find primary non-temp .docx inside a special-category folder."""
    if not dir_path or not os.path.isdir(dir_path):
        return ""
    folder_name = os.path.basename(dir_path)
    preferred = os.path.join(dir_path, f"[病理报告]{folder_name}.docx")
    if os.path.isfile(preferred):
        return preferred
    try:
        for fn in sorted(os.listdir(dir_path)):
            if not fn.startswith("~$") and fn.lower().endswith(".docx"):
                full = os.path.join(dir_path, fn)
                if os.path.isfile(full):
                    return full
    except Exception:
        pass
    return ""


def collect_patient_doc_paths(work_dir, patient_data, ihc_data=None):
    """
    Collect all existing Word document paths across the 3 tiers for a patient (ignoring '~$' temp files):
      - Tier 1 ('entry'):   entry (未归档报告/病理报告录入)
      - Tier 2 ('draft'):   draft_output (病理报告输出), draft_pending (待审核), draft_recut (重切), draft_approved (完成审核)
      - Tier 3 ('archive'): archive (已归档报告 non-IHC)
      - IHC Tier 2 ('ihc_draft'):   ihc_draft (特殊患者归档/免疫组化)
      - IHC Tier 3 ('ihc_archive'): ihc_archive (已归档报告 *_免疫组化报告.docx)
    """
    if not work_dir or not patient_data:
        return {}
    p_no = patient_data.get("pathology_no", "")
    p_name = patient_data.get("name", "")
    docs = {}

    def _add(role, tier_group, label, path):
        if path and os.path.isfile(path) and not os.path.basename(path).startswith("~$"):
            docs[role] = {
                "role": role,
                "tier_group": tier_group,
                "label": label,
                "path": os.path.normpath(os.path.abspath(path)),
            }

    # Tier 1: Entry
    entry_p = patient_data.get("word_path") or get_patient_word_path(work_dir, p_no, p_name)
    _add("entry", "entry", "病理报告录入", entry_p)

    # Tier 2: Drafts
    out_p = patient_data.get("final_report_path") or ""
    if not out_p or not os.path.isfile(out_p):
        out_dir = get_output_report_dir(work_dir)
        if os.path.isdir(out_dir):
            p_diag = patient_data.get("pathology_diag") or ""
            if p_diag:
                cand_out = get_final_report_path(work_dir, p_no, p_name, p_diag)
                if os.path.isfile(cand_out):
                    out_p = cand_out
            if not out_p or not os.path.isfile(out_p):
                for fn in sorted(os.listdir(out_dir)):
                    if not fn.startswith("~$") and fn.lower().endswith(".docx") and _matches_patient_no(fn, p_no):
                        out_p = os.path.join(out_dir, fn)
                        break
    _add("draft_output", "draft", "病理报告输出", out_p)

    for role, cat_name, label in (
        ("draft_pending", "待审核", "特殊患者归档/待审核"),
        ("draft_recut", "重切", "特殊患者归档/重切"),
        ("draft_approved", "完成审核", "特殊患者归档/完成审核"),
    ):
        fdir = _find_special_cat_folder(work_dir, cat_name, p_no, p_name)
        if fdir:
            _add(role, "draft", label, _find_docx_inside_dir(fdir))

    # Tier 3: Regular archive
    arch_p = patient_data.get("archive_dir") or ""
    if not arch_p or not os.path.isfile(arch_p):
        arch_p = find_patient_archive_path(work_dir, p_no, p_name, extra_archive_path=arch_p)
    _add("archive", "archive", "已归档报告", arch_p)

    # IHC Tier 2 & 3
    if ihc_data is not None:
        ihc_doc_p = ihc_data.get("ihc_doc_path") or ""
        if not ihc_doc_p or not os.path.isfile(ihc_doc_p):
            ihc_dir = ihc_data.get("ihc_dir") or find_patient_ihc_folder(work_dir, p_no, p_name)
            if ihc_dir:
                ihc_doc_p = _find_docx_inside_dir(ihc_dir)
        _add("ihc_draft", "ihc_draft", "免疫组化草稿", ihc_doc_p)

        ihc_arch_p = ihc_data.get("archive_file") or ""
        if not ihc_arch_p or not os.path.isfile(ihc_arch_p):
            ihc_arch_p = find_patient_ihc_archive_path(work_dir, p_no, p_name, extra_archive_path=ihc_arch_p)
        _add("ihc_archive", "ihc_archive", "已归档免疫组化报告", ihc_arch_p)

    return docs


def _parse_sync_meta(existing_meta):
    if isinstance(existing_meta, str):
        try:
            import json
            existing_meta = json.loads(existing_meta) if existing_meta.strip() else {}
        except Exception:
            existing_meta = {}
    if not isinstance(existing_meta, dict):
        return {"files": {}}
    files = existing_meta.get("files")
    if not isinstance(files, dict):
        files = {}
    return {"files": {str(k): dict(v) for k, v in files.items() if isinstance(v, dict)}}


def update_doc_sync_meta_entries(existing_meta, role_path_pairs, prune_missing=True):
    """
    Update tracked SHA-256 fingerprints in a doc_sync_meta dict for the given role -> file_path pairs.
    If file_path is empty/missing for a role, removes stale records for that role.
    """
    meta = _parse_sync_meta(existing_meta)
    files = meta["files"]

    items = role_path_pairs.items() if isinstance(role_path_pairs, dict) else list(role_path_pairs or [])
    for role, fp in items:
        # Remove previous keys bound to this role
        stale_keys = [k for k, rec in files.items() if rec.get("role") == role]
        for k in stale_keys:
            files.pop(k, None)
        if fp and os.path.isfile(fp) and not os.path.basename(fp).startswith("~$"):
            info = compute_file_fingerprint(fp)
            if info:
                info["role"] = role
                files[normalize_sync_path(fp)] = info

    if prune_missing:
        missing_keys = [
            k for k, rec in files.items()
            if not rec.get("path") or not os.path.isfile(rec.get("path"))
        ]
        for k in missing_keys:
            files.pop(k, None)

    return {"files": files}


def remap_doc_sync_meta_paths(existing_meta, renamed_pairs):
    """
    Update stored path keys after rename_patient_files_and_folders without altering recorded SHA-256 hashes.
    """
    meta = _parse_sync_meta(existing_meta)
    files = meta["files"]
    if not files or not renamed_pairs:
        return meta

    norm_pairs = sorted(
        [(normalize_sync_path(old_p), os.path.normpath(os.path.abspath(new_p))) for old_p, new_p in renamed_pairs if old_p and new_p],
        key=lambda item: len(item[0]),
        reverse=True,
    )
    updated_files = {}
    for k, rec in files.items():
        cur_path = rec.get("path") or ""
        cur_norm = normalize_sync_path(cur_path) or k
        new_rec_path = cur_path
        matched = False
        for old_norm, new_abs in norm_pairs:
            if cur_norm == old_norm:
                new_rec_path = new_abs
                matched = True
                break
        if not matched:
            for old_norm, new_abs in norm_pairs:
                sep_old = old_norm + os.sep
                if cur_norm.startswith(sep_old):
                    rel = cur_path[len(old_norm):].lstrip("\\/")
                    new_rec_path = os.path.join(new_abs, rel)
                    break
        rec_copy = dict(rec)
        rec_copy["path"] = new_rec_path
        updated_files[normalize_sync_path(new_rec_path)] = rec_copy
    return {"files": updated_files}


def inspect_patient_doc_sync(work_dir, patient_data, ihc_data=None):
    """
    Compare current disk files against recorded SHA-256 fingerprints in patients.doc_sync_meta
    (and ihc_reports.doc_sync_meta if ihc_data is provided).
    Returns detailed external-modification state grouped cleanly by the 3 tiers.
    """
    existing_docs = collect_patient_doc_paths(work_dir, patient_data, ihc_data=ihc_data)
    patient_files_meta = _parse_sync_meta((patient_data or {}).get("doc_sync_meta")).get("files", {})
    ihc_files_meta = _parse_sync_meta((ihc_data or {}).get("doc_sync_meta")).get("files", {}) if ihc_data is not None else {}

    # Any tracked Tier 2 draft hash lets us detect untracked externally dropped Tier 2 sibling copies too
    tracked_draft_hashes = {
        rec.get("sha256")
        for rec in patient_files_meta.values()
        if str(rec.get("role", "")).startswith("draft_") and rec.get("sha256")
    }

    modified_files = []
    ihc_modified_files = []

    for role, info in existing_docs.items():
        fpath = info["path"]
        norm_p = normalize_sync_path(fpath)
        is_ihc_role = role.startswith("ihc_")
        tracked_map = ihc_files_meta if is_ihc_role else patient_files_meta
        rec = tracked_map.get(norm_p)

        try:
            st = os.stat(fpath)
            cur_size = int(st.st_size)
            cur_mtime_ns = int(getattr(st, "st_mtime_ns", int(st.st_mtime * 1e9)))
            cur_mtime = float(st.st_mtime)
        except OSError:
            continue

        is_changed = False
        cur_sha = ""
        rec_sha = ""

        if rec and rec.get("sha256"):
            rec_sha = rec.get("sha256", "")
            if cur_size != int(rec.get("size", -1)) or cur_mtime_ns != int(rec.get("mtime_ns", -1)):
                fp = compute_file_fingerprint(fpath)
                if fp and fp["sha256"] != rec_sha:
                    is_changed = True
                    cur_sha = fp["sha256"]
        elif (not is_ihc_role) and info["tier_group"] == "draft" and tracked_draft_hashes:
            fp = compute_file_fingerprint(fpath)
            if fp and fp["sha256"] not in tracked_draft_hashes:
                is_changed = True
                cur_sha = fp["sha256"]
                rec_sha = next(iter(tracked_draft_hashes), "")

        if is_changed:
            entry_info = {
                "role": role,
                "tier_group": info["tier_group"],
                "label": info["label"],
                "path": fpath,
                "filename": os.path.basename(fpath),
                "current_sha256": cur_sha,
                "recorded_sha256": rec_sha,
                "mtime": cur_mtime,
            }
            if is_ihc_role:
                ihc_modified_files.append(entry_info)
            else:
                modified_files.append(entry_info)

    # Sort newest-modified first so UI/default source selection prioritizes the latest externally saved file
    modified_files.sort(key=lambda x: (-x.get("mtime", 0), x.get("role", "")))
    ihc_modified_files.sort(key=lambda x: (-x.get("mtime", 0), x.get("role", "")))

    return {
        "external_modified": bool(modified_files),
        "ihc_external_modified": bool(ihc_modified_files),
        "any_external_modified": bool(modified_files or ihc_modified_files),
        "modified_files": modified_files,
        "ihc_modified_files": ihc_modified_files,
        "modified_tiers": sorted(list({m["tier_group"] for m in modified_files})),
        "ihc_modified_tiers": sorted(list({m["tier_group"] for m in ihc_modified_files})),
        "doc_paths": {k: v["path"] for k, v in existing_docs.items()},
    }


def distribute_docx_copy(src_path, dst_paths, lock_label="报告文件"):
    """
    Distribute a source .docx (preserving 100% of external custom styles/colors/comments)
    to sibling paths, backing up every overwritten destination to the Recycle Bin first.
    """
    if not src_path or not os.path.isfile(src_path):
        raise FileNotFoundError(f"源文件不存在: {src_path}")
    copied = []
    for dst_p in (dst_paths or []):
        if not dst_p:
            continue
        if _paths_same_entry(src_path, dst_p):
            copied.append(dst_p)
            continue
        os.makedirs(os.path.dirname(dst_p), exist_ok=True)
        _ensure_writable(dst_p, lock_label=lock_label)
        backup_before_overwrite(dst_p, lock_label=lock_label)
        shutil.copy2(src_path, dst_p)
        copied.append(dst_p)
    return copied
