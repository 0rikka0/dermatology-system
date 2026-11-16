# -*- coding: utf-8 -*-
"""
Dermatology Pathology Reporting System - Template & Diagnosis Manager
Handles:
1. Parsing and caching normal and IHC microscopic description templates from docx
2. Intelligent template recommendation based on photo diagnoses
3. Microscope photo filename parsing (序号 + 诊断)
4. Auto-generating standardized pathological diagnosis text
"""
import os
import re
import docx

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEXT_TEMPLATES_DIR = os.path.join(BASE_DIR, "文字模板")

_TEMPLATES_CACHE = None

def load_text_templates(templates_dir=None):
    """
    Parses 普通模板.docx and 免疫组化模板.docx into structured records.
    Cached in memory.
    """
    global _TEMPLATES_CACHE
    if _TEMPLATES_CACHE is not None:
        return _TEMPLATES_CACHE

    target_dir = templates_dir or TEXT_TEMPLATES_DIR
    templates = []
    
    # 1. Parse 普通模板.docx
    normal_path = os.path.join(target_dir, "普通模板.docx")
    if os.path.exists(normal_path):
        try:
            doc = docx.Document(normal_path)
            for idx, p in enumerate(doc.paragraphs):
                t = p.text.strip()
                if not t:
                    continue
                delim = "：" if "：" in t else (":" if ":" in t else None)
                if delim:
                    parts = t.split(delim, 1)
                    name = parts[0].strip()
                    desc = parts[1].strip()
                else:
                    name = t[:20].strip()
                    desc = t.strip()
                if name and desc:
                    templates.append({
                        "id": f"norm_{idx}",
                        "name": name,
                        "desc": desc,
                        "category": "普通模板",
                        "full_text": t
                    })
        except Exception as e:
            print(f"[TemplateManager] Error loading 普通模板: {e}")

    # 2. Parse 免疫组化模板.docx
    ihc_path = os.path.join(target_dir, "免疫组化模板.docx")
    if os.path.exists(ihc_path):
        try:
            doc = docx.Document(ihc_path)
            for idx, p in enumerate(doc.paragraphs):
                t = p.text.strip()
                if not t:
                    continue
                delim = "：" if "：" in t else (":" if ":" in t else None)
                if delim:
                    parts = t.split(delim, 1)
                    name = parts[0].strip()
                    desc = parts[1].strip()
                else:
                    name = t[:20].strip()
                    desc = t.strip()
                if name and desc:
                    templates.append({
                        "id": f"ihc_{idx}",
                        "name": name,
                        "desc": desc,
                        "category": "免疫组化模板",
                        "full_text": t
                    })
        except Exception as e:
            print(f"[TemplateManager] Error loading 免疫组化模板: {e}")

    _TEMPLATES_CACHE = templates
    return _TEMPLATES_CACHE

def reload_text_templates():
    """Clear cache and reload text templates."""
    global _TEMPLATES_CACHE
    _TEMPLATES_CACHE = None
    return load_text_templates()

def search_templates(keyword=None, category=None):
    """
    Search templates by keyword and/or category.
    """
    all_tmpls = load_text_templates()
    results = all_tmpls
    
    if category and category != "全部":
        results = [t for t in results if t["category"] == category]
        
    if keyword:
        kw = keyword.strip().lower()
        results = [
            t for t in results
            if kw in t["name"].lower() or kw in t["desc"].lower()
        ]
        
    return results

def recommend_templates(diagnoses):
    """
    Recommend templates based on a list of diagnoses extracted from photos.
    Sorts exact matches to the top.
    """
    all_tmpls = load_text_templates()
    if not diagnoses:
        return []

    recommended = []
    seen_ids = set()

    for diag in diagnoses:
        d = diag.strip().lower()
        if not d:
            continue
        
        # 1. Exact match in name
        for t in all_tmpls:
            if t["id"] in seen_ids:
                continue
            t_name_lower = t["name"].lower()
            if d == t_name_lower:
                recommended.append({**t, "match_type": "exact", "matched_diag": diag})
                seen_ids.add(t["id"])

        # 2. Substring match in name
        for t in all_tmpls:
            if t["id"] in seen_ids:
                continue
            t_name_lower = t["name"].lower()
            if d in t_name_lower or t_name_lower in d:
                recommended.append({**t, "match_type": "name_contains", "matched_diag": diag})
                seen_ids.add(t["id"])

        # 3. Substring match in description
        for t in all_tmpls:
            if t["id"] in seen_ids:
                continue
            if d in t["desc"].lower():
                recommended.append({**t, "match_type": "desc_contains", "matched_diag": diag})
                seen_ids.add(t["id"])

    return recommended

def parse_photo_filename(filename):
    """
    Extracts (indices, diagnosis) from a microscope photo filename.
    Examples:
      '1交界痣.jpg' -> indices=['1'], diagnosis='交界痣'
      '1.2.交界痣.jpg' -> indices=['1', '2'], diagnosis='交界痣'
      '1-表皮囊肿.png' -> indices=['1'], diagnosis='表皮囊肿'
      '表皮囊肿.jpg' -> indices=[], diagnosis='表皮囊肿'
    """
    name, ext = os.path.splitext(filename)
    ext = ext.lower()
    if ext not in [".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"]:
        return None

    # Match leading numbers/dots/dashes/underscores
    m = re.match(r"^([\d\.\-_、]+)(.*)$", name)
    if m:
        prefix_str = m.group(1)
        diag = m.group(2).strip()
        nums = re.findall(r"\d+", prefix_str)
        diag = diag.strip(" .-_、")
        if not diag:
            diag = name
        return {
            "filename": filename,
            "indices": nums,
            "diagnosis": diag,
            "index_num": int(nums[0]) if nums else 999
        }
    else:
        return {
            "filename": filename,
            "indices": [],
            "diagnosis": name.strip(),
            "index_num": 999
        }

def format_pathology_diagnosis(items, image_count=None):
    """
    Synthesizes standard pathological diagnosis from photo diagnoses.
    Format examples:
      - Single-image report (image_count == 1 or single item): '表皮囊肿' (no index prefix)
      - Multiple: '1.2.交界痣；3.皮内痣'
      - Different: '1.表皮囊肿；2.脂肪瘤'
    """
    if not items:
        return ""

    cleaned = []
    for it in items:
        if isinstance(it, (list, tuple)) and len(it) >= 2:
            if isinstance(it[0], (list, tuple)):
                idx = ".".join([str(x).strip() for x in it[0] if str(x).strip()]).rstrip(".").rstrip("-")
            else:
                idx = str(it[0]).strip().rstrip(".").rstrip("-")
            diag = str(it[1]).strip()
        elif isinstance(it, dict):
            indices = it.get("indices")
            if indices and isinstance(indices, list):
                idx = ".".join([str(x) for x in indices])
            else:
                idx = str(it.get("index", "")).strip().rstrip(".").rstrip("-")
            diag = str(it.get("diagnosis", "")).strip()
        else:
            continue
        if diag:
            cleaned.append((idx, diag))

    if not cleaned:
        return ""

    is_single_image_report = False
    if image_count is not None:
        try:
            is_single_image_report = (int(image_count) == 1)
        except (TypeError, ValueError):
            is_single_image_report = False

    if is_single_image_report:
        unique_diags = []
        for _, diag in cleaned:
            if diag not in unique_diags:
                unique_diags.append(diag)
        return "；".join(unique_diags)

    if len(cleaned) == 1 and image_count is None:
        return cleaned[0][1]

    if len(cleaned) == 1 and not cleaned[0][0]:
        return cleaned[0][1]

    # Group by diagnosis preserving first appearance order
    groups = []
    diag_map = {}
    for idx, diag in cleaned:
        if diag not in diag_map:
            diag_map[diag] = len(groups)
            groups.append((diag, [idx] if idx else []))
        else:
            if idx:
                groups[diag_map[diag]][1].append(idx)

    # If only 1 group and no indices were ever found
    if len(groups) == 1 and not groups[0][1]:
        return groups[0][0]

    parts = []
    for diag, indices in groups:
        # filter out empty indices
        valid_indices = []
        for x in indices:
            if "." in x:
                valid_indices.extend(x.split("."))
            elif x:
                valid_indices.append(x)
                
        # deduplicate while preserving order
        dedup_indices = []
        for v in valid_indices:
            if v not in dedup_indices:
                dedup_indices.append(v)

        if dedup_indices:
            idx_prefix = ".".join(dedup_indices) + "."
            parts.append(f"{idx_prefix}{diag}")
        else:
            parts.append(diag)

    return "；".join(parts)
