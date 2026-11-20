# -*- coding: utf-8 -*-
"""
Dermatology Pathology Reporting System - Document Preview Renderer
Generates an A4 medical report HTML preview from docx and patient data.
Supports exact multi-image table layout (1-9 images) and 2-page preview.
"""
import os
import re
import json
import html
import urllib.parse
import docx


def _normalize_extracted_date(raw_text):
    """Normalize Chinese or ISO date text into (cn_str, raw_yyyy_mm_dd)."""
    s = str(raw_text or "").strip()
    if not s:
        return "", ""
    m_cn = re.search(r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日?", s)
    if m_cn:
        y, m, d = m_cn.groups()
        return f"{y}年{int(m):02d}月{int(d):02d}日", f"{y}-{int(m):02d}-{int(d):02d}"
    m_iso = re.search(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", s)
    if m_iso:
        y, m, d = m_iso.groups()
        return f"{y}年{int(m):02d}月{int(d):02d}日", f"{y}-{int(m):02d}-{int(d):02d}"
    return s, ""


def _parse_table0_fields(table0):
    """Parse key-value metadata cells from Table 0 while deduplicating merged cells."""
    fields = {}
    label_map = {
        "姓名": "name",
        "性别": "gender",
        "年龄": "age",
        "病理号": "pathology_no",
        "送检单位": "send_unit",
        "科室": "department",
        "送检日期": "submission_date",
        "住院号": "inpatient_no",
        "门诊号": "outpatient_no",
        "送检医生": "referring_doctor",
        "临床诊断": "clinical_diagnosis",
        "送检部位": "single_biopsy_site",
    }
    seen_tc = set()
    for row in table0.rows:
        for cell in row.cells:
            tc = cell._tc
            if tc in seen_tc:
                continue
            seen_tc.add(tc)
            cell_text = re.sub(r"\s+", " ", cell.text.replace("\r", " ").replace("\n", " ")).strip()
            if not cell_text:
                continue
            m = re.match(r"^([^：:]+?)[：:]\s*(.*)$", cell_text)
            if not m:
                continue
            lbl, val = m.group(1).strip(), m.group(2).strip()
            if lbl in label_map:
                fields[label_map[lbl]] = val
    return fields


def _parse_image_site_table(tables):
    """Parse numbered biopsy sites (1. xxx, 2. xxx ...) below image grid in Table 1."""
    if len(tables) < 2:
        return None, []
    site_map = {}
    seen_tc = set()
    for tbl in tables[1:]:
        for row in tbl.rows:
            for cell in row.cells:
                tc = cell._tc
                if tc in seen_tc:
                    continue
                seen_tc.add(tc)
                t = cell.text.strip()
                if not t:
                    continue
                m = re.match(r"^(\d+)\s*[.、．]\s*(.*)$", t, re.S)
                if m:
                    idx = int(m.group(1))
                    if 1 <= idx <= 9:
                        site_map[idx] = m.group(2).strip()
    if not site_map:
        return None, []
    max_idx = max(site_map.keys())
    sites = [site_map.get(i, "") for i in range(1, max_idx + 1)]
    return max_idx, sites


def extract_docx_details(doc_path):
    """
    Full reverse parser for generated or externally edited pathology Word documents.
    Extracts Table 0 patient info, Table 1 multi-site labels, microscopic description,
    pathology diagnosis, doctor signature line, and medical history.
    """
    if not doc_path or not os.path.exists(doc_path):
        return None
    try:
        doc = docx.Document(doc_path)
        content = {
            "microscopic_desc": "",
            "pathology_diag": "",
            "history_line": "",
            "medical_history": "",
        }

        if doc.tables:
            t0_fields = _parse_table0_fields(doc.tables[0])
            content.update(t0_fields)
            if "pathology_no" in content:
                content["pathology_no"] = content["pathology_no"].strip().upper()
            if "submission_date" in content:
                sub_cn, sub_raw = _normalize_extracted_date(content["submission_date"])
                content["submission_date"] = sub_cn or content["submission_date"]
                content["submission_date_raw"] = sub_raw

        img_cnt_t1, sites_t1 = _parse_image_site_table(doc.tables)
        if img_cnt_t1 is not None:
            content["image_count"] = img_cnt_t1
            content["biopsy_sites"] = sites_t1
        elif "single_biopsy_site" in content:
            single_site = content["single_biopsy_site"]
            content["image_count"] = 1
            content["biopsy_sites"] = [single_site] if single_site else []

        current_section = None
        for p in doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue
            if text.startswith("镜下描述：") or text.startswith("镜下描述:"):
                current_section = "micro"
                desc = re.sub(r"^镜下描述[：:]\s*", "", text).strip()
                if desc:
                    content["microscopic_desc"] = desc
            elif text.startswith("病理诊断：") or text.startswith("病理诊断:"):
                current_section = "diag"
                diag = re.sub(r"^病理诊断[：:]\s*", "", text).strip()
                if diag:
                    content["pathology_diag"] = diag
            elif ("报告医生" in text and ("审核医生" in text or "报告日期" in text)) or text.startswith("报告医生：") or text.startswith("报告医生:"):
                current_section = None
                m_rep = re.search(r"报告医生[：:]\s*(.*?)(?=\s*(?:审核医生|报告日期)[：:]|$)", text)
                m_rev = re.search(r"审核医生[：:]\s*(.*?)(?=\s*报告日期[：:]|$)", text)
                m_date = re.search(r"报告日期[：:]\s*(.*)$", text)
                if m_rep:
                    content["report_doctor"] = m_rep.group(1).strip()
                if m_rev:
                    content["review_doctor"] = m_rev.group(1).strip()
                if m_date:
                    rep_cn, rep_raw = _normalize_extracted_date(m_date.group(1).strip())
                    content["report_date"] = rep_cn
                    content["report_date_raw"] = rep_raw
            elif text.startswith("病史：") or text.startswith("病史:"):
                current_section = "history"
                hist_val = re.sub(r"^病史[：:]\s*", "", text).strip()
                content["history_line"] = hist_val
                content["medical_history"] = hist_val
            else:
                if current_section == "micro":
                    content["microscopic_desc"] += ("\n" + text if content["microscopic_desc"] else text)
                elif current_section == "diag":
                    content["pathology_diag"] += ("\n" + text if content["pathology_diag"] else text)
                elif current_section == "history":
                    content["history_line"] += ("\n" + text if content["history_line"] else text)
                    content["medical_history"] = content["history_line"]

        return content
    except Exception as e:
        print(f"[DocPreview] Error reading docx: {e}")
        return None

def render_preview_html(patient_data, doc_path=None):
    """
    Renders an HTML snippet representing the A4 pathology report.
    """
    docx_content = extract_docx_details(doc_path) if doc_path else None
    
    name = html.escape(str(patient_data.get("name", "")))
    gender = html.escape(str(patient_data.get("gender", "")))
    age = html.escape(str(patient_data.get("age", "")))
    if age and not age.endswith(("岁", "月", "日", "天")):
        age += "岁"
    pathology_no = html.escape(str(patient_data.get("pathology_no", "")))
    sub_date = html.escape(str(patient_data.get("submission_date", "")))
    inpatient_no = html.escape(str(patient_data.get("inpatient_no", "")))
    outpatient_no = html.escape(str(patient_data.get("outpatient_no", "")))
    ref_doc = html.escape(str(patient_data.get("referring_doctor", "")))
    diag = html.escape(str(patient_data.get("clinical_diagnosis", "")))
    image_count = int(patient_data.get("image_count", 1))
    
    biopsy_sites = patient_data.get("biopsy_sites", [])
    if isinstance(biopsy_sites, str):
        biopsy_sites = [s.strip() for s in biopsy_sites.split(";") if s.strip()]
        
    send_unit = html.escape(str(patient_data.get("send_unit") or "本院"))
    department = html.escape(str(patient_data.get("department") or "皮肤科"))
    rep_doc = html.escape(str(patient_data.get("report_doctor", "")))
    rev_doc = html.escape(str(patient_data.get("review_doctor", "")))
    rep_date = html.escape(str(patient_data.get("report_date", "")))
    history = html.escape(docx_content.get("history_line") if docx_content and docx_content.get("history_line") else str(patient_data.get("medical_history", "")))
    
    # Table 0 Row 3 layout
    if image_count == 1:
        single_site = html.escape(biopsy_sites[0]) if biopsy_sites else ""
        row3_html = f"""
        <tr>
            <td colspan="2" class="table-cell"><strong>临床诊断：</strong>{diag}</td>
            <td colspan="2" class="table-cell"><strong>送检部位：</strong>{single_site}</td>
        </tr>
        """
    else:
        row3_html = f"""
        <tr>
            <td colspan="4" class="table-cell"><strong>临床诊断：</strong>{diag}</td>
        </tr>
        """
        
    # Parse selected_images
    selected_images = patient_data.get("selected_images", [])
    if isinstance(selected_images, str):
        try:
            selected_images = json.loads(selected_images)
        except Exception:
            selected_images = []
            
    img_paths = [None] * image_count
    if isinstance(selected_images, list):
        for idx, item in enumerate(selected_images):
            if isinstance(item, dict):
                slot = item.get("slot", idx + 1) - 1
                p = item.get("path")
                if 0 <= slot < image_count:
                    img_paths[slot] = p
            elif isinstance(item, str) and idx < image_count:
                img_paths[idx] = item

    # Render Image & Site Grid (Table Layout in WebUI)
    if image_count == 1:
        p0 = img_paths[0] if (img_paths and img_paths[0] and os.path.exists(img_paths[0])) else None
        if p0:
            p_url = f"/api/photo_file?path={urllib.parse.quote(p0)}"
            inner_img = f'<img src="{p_url}" alt="照片 1" style="max-width:100%; max-height:220px; object-fit:contain; border-radius:4px;">'
        else:
            inner_img = '<div class="img-box-inner"><span class="img-icon">🖼️</span><span class="img-label">照片 1 (待选择)</span></div>'

        images_grid_html = f"""
        <div class="report-single-img-container">
            <div class="report-img-placeholder single">
                {inner_img}
            </div>
        </div>
        """
    else:
        cols = 2 if image_count == 2 else 3
        num_rows = (image_count + cols - 1) // cols
        
        table_rows_html = []
        for r_idx in range(num_rows):
            img_cells = []
            lbl_cells = []
            for c_idx in range(cols):
                idx = r_idx * cols + c_idx + 1
                if idx <= image_count:
                    s_name = html.escape(biopsy_sites[idx - 1]) if len(biopsy_sites) >= idx else ""
                    lbl_text = f"{idx}. {s_name}" if s_name else f"{idx}."
                    
                    p_curr = img_paths[idx - 1] if (idx - 1 < len(img_paths) and img_paths[idx - 1] and os.path.exists(img_paths[idx - 1])) else None
                    if p_curr:
                        p_url = f"/api/photo_file?path={urllib.parse.quote(p_curr)}"
                        box_content = f'<img src="{p_url}" alt="照片 {idx}" style="width:100%; height:100%; object-fit:cover; border-radius:3px;">'
                    else:
                        box_content = f'<div class="img-box-inner"><span class="img-icon">🖼️</span><span class="img-label">照片 {idx}</span></div>'

                    img_cells.append(f"""
                        <td class="preview-img-cell">
                            <div class="report-img-placeholder grid-box">
                                {box_content}
                            </div>
                        </td>
                    """)
                    lbl_cells.append(f"""
                        <td class="preview-lbl-cell">
                            <div class="site-label-aligned">{lbl_text}</div>
                        </td>
                    """)
                else:
                    # Empty cell for 4, 5, 7, 8
                    img_cells.append('<td class="preview-img-cell empty"></td>')
                    lbl_cells.append('<td class="preview-lbl-cell empty"></td>')
                    
            table_rows_html.append(f"<tr>{''.join(img_cells)}</tr>")
            table_rows_html.append(f"<tr>{''.join(lbl_cells)}</tr>")
            
        images_grid_html = f"""
        <table class="report-image-table cols-{cols}">
            <tbody>
                {''.join(table_rows_html)}
            </tbody>
        </table>
        """

    micro_desc = html.escape(patient_data.get("microscopic_desc") or (docx_content.get("microscopic_desc", "") if docx_content else ""))
    patho_diag = html.escape(patient_data.get("pathology_diag") or (docx_content.get("pathology_diag", "") if docx_content else ""))

    html_out = f"""
    <div class="a4-report-sheet">
        <!-- PAGE 1 -->
        <div class="a4-page page-1">
            <div class="report-header">
                <h1 class="hospital-title">某三甲医院皮肤科</h1>
                <h2 class="report-title">皮肤组织病理报告</h2>
            </div>
            
            <table class="report-table">
                <tbody>
                    <tr>
                        <td class="table-cell w-25"><strong>姓名：</strong>{name}</td>
                        <td class="table-cell w-25"><strong>性别：</strong>{gender}</td>
                        <td class="table-cell w-25"><strong>年龄：</strong>{age}</td>
                        <td class="table-cell w-25"><strong>病理号：</strong>{pathology_no}</td>
                    </tr>
                    <tr>
                        <td class="table-cell"><strong>送检单位：</strong>{send_unit}</td>
                        <td class="table-cell"><strong>科室：</strong>{department}</td>
                        <td colspan="2" class="table-cell"><strong>送检日期：</strong>{sub_date}</td>
                    </tr>
                    <tr>
                        <td class="table-cell"><strong>住院号：</strong>{inpatient_no}</td>
                        <td class="table-cell"><strong>门诊号：</strong>{outpatient_no}</td>
                        <td colspan="2" class="table-cell"><strong>送检医生：</strong>{ref_doc}</td>
                    </tr>
                    {row3_html}
                </tbody>
            </table>
            
            {images_grid_html}
            
            <div class="report-section">
                <div class="section-title">镜下描述：</div>
                <div class="section-content text-placeholder">{micro_desc or '（待医生录入镜下描述）'}</div>
            </div>
            
            <div class="report-section">
                <div class="section-title">病理诊断：</div>
                <div class="section-content text-placeholder">{patho_diag or '（待医生录入病理诊断）'}</div>
            </div>
            
            <div class="report-footer-signatures">
                <div class="signatures">
                    <span><strong>报告医生：</strong>{rep_doc}</span>
                    <span><strong>审核医生：</strong>{rev_doc}</span>
                    <span><strong>报告日期：</strong>{rep_date}</span>
                </div>
            </div>
        </div>

        <!-- PAGE BREAK INDICATOR -->
        <div class="preview-page-divider">
            <span>✂️ --- 第 2 页 (病史) ---</span>
        </div>

        <!-- PAGE 2 -->
        <div class="a4-page page-2">
            <div class="history-section">
                <strong>病史：</strong>{history or '（无特殊病史）'}
            </div>
        </div>
    </div>
    """
    return html_out
