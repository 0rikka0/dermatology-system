# -*- coding: utf-8 -*-
"""
Dermatology Pathology Reporting System - Document Preview Renderer
Generates an A4 medical report HTML preview from docx and patient data.
Supports exact multi-image table layout (1-9 images) and 2-page preview.
"""
import os
import json
import html
import urllib.parse
import docx

def extract_docx_details(doc_path):
    """Extract text fields from docx if file exists."""
    if not os.path.exists(doc_path):
        return None
    try:
        doc = docx.Document(doc_path)
        content = {
            "microscopic_desc": "",
            "pathology_diag": "",
            "history_line": ""
        }
        
        current_section = None
        for p in doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue
            if "镜下描述：" in text:
                current_section = "micro"
                desc = text.replace("镜下描述：", "").strip()
                if desc:
                    content["microscopic_desc"] = desc
            elif "病理诊断：" in text:
                current_section = "diag"
                diag = text.replace("病理诊断：", "").strip()
                if diag:
                    content["pathology_diag"] = diag
            elif "报告医生：" in text and "审核医生：" in text:
                current_section = None
            elif text.startswith("病史："):
                current_section = None
                content["history_line"] = text.replace("病史：", "").strip()
            else:
                if current_section == "micro":
                    content["microscopic_desc"] += ("\n" + text if content["microscopic_desc"] else text)
                elif current_section == "diag":
                    content["pathology_diag"] += ("\n" + text if content["pathology_diag"] else text)
                    
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
