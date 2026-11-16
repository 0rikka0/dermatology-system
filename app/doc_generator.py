# -*- coding: utf-8 -*-
"""
Dermatology Pathology Reporting System - Word Document Generator & Updater
Implements robust Table-based layout for image placeholders and biopsy site alignment.
Handles:
- Exact 1-9 images layout
- For 3 images: 1 row of 3 images, subsequent sections shift up, signatures on Page 1, history on Page 2
- For 4-5 images: 2 rows of 3 images, extra blank placeholders removed, layout preserved
- For 6 images: 2 full rows of 3 images
- For 7-9 images: automatically adds 3rd row, compacts empty lines to ensure signatures stay on Page 1
- Exact alignment of biopsy site labels directly below each respective picture
- Guaranteed Page 1 for signatures and Page 2 for history
"""
import os
import re
import math
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATES_DIR = os.path.join(BASE_DIR, "模板")
PLACEHOLDER_IMG = os.path.join(TEMPLATES_DIR, "placeholder.jpeg")

def count_visual_lines(text, chars_per_line=39.0):
    """Calculate the estimated visual lines of text at 12pt 宋体 on A4 page."""
    if not text or not text.strip():
        return 1
    eff_len = sum(1.0 if ord(c) > 127 else 0.55 for c in text.strip())
    return max(1, math.ceil(eff_len / chars_per_line))

def set_run_font(run, font_name="宋体", size_pt=12, bold=False):
    """Apply East Asian and Latin font formatting to a run."""
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    run.font.size = Pt(size_pt)
    run.bold = bold

def set_cell_margins(cell, top=30, bottom=30, left=30, right=30):
    """Set inner padding of a table cell in dxa (1 pt = 20 dxa)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def remove_extra_separator_above_table(doc, table0):
    """
    Remove floating VML lines (e.g. 直线 2) or borders in paragraphs above Table 0
    to ensure there is no duplicate separator line above the table.
    """
    for p in doc.paragraphs:
        p_elem = p._p
        # If paragraph is above or at table0
        pict_nodes = p_elem.findall('.//' + qn('w:pict'))
        for pict in pict_nodes:
            for child in pict.iter():
                if child.tag.endswith('line'):
                    parent = pict.getparent()
                    if parent is not None:
                        parent.remove(pict)
                    break
        if p_elem.pPr is not None:
            pBdr = p_elem.pPr.find(qn('w:pBdr'))
            if pBdr is not None:
                p_elem.pPr.remove(pBdr)

def set_table0_borders(table):
    """Ensure Table 0 has explicit top border and bottom border (three-line table style) with no duplicates."""
    tblPr = table._tbl.tblPr
    existing_bdr = tblPr.find(qn('w:tblBorders'))
    if existing_bdr is not None:
        tblPr.remove(existing_bdr)
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/>\n'
        f'  <w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>\n'
        f'  <w:left w:val="none"/>\n'
        f'  <w:right w:val="none"/>\n'
        f'  <w:insideH w:val="none"/>\n'
        f'  <w:insideV w:val="none"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

    # Clean any conflicting cell-level borders on rows
    for row in table.rows:
        for c in row.cells:
            tcPr = c._tc.get_or_add_tcPr()
            tcBorders = tcPr.find(qn('w:tcBorders'))
            if tcBorders is not None:
                tcPr.remove(tcBorders)

def remove_table_borders(table):
    """Remove all grid borders from a table."""
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="none"/>\n'
        f'  <w:left w:val="none"/>\n'
        f'  <w:bottom w:val="none"/>\n'
        f'  <w:right w:val="none"/>\n'
        f'  <w:insideH w:val="none"/>\n'
        f'  <w:insideV w:val="none"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def ensure_placeholder_image(templates_dir):
    """Ensure placeholder.jpeg exists in templates_dir."""
    target = os.path.join(templates_dir, "placeholder.jpeg")
    if not os.path.exists(target):
        for fn in ["图片数大于等于3.docx", "图片数等于1.docx", "图片数等于2.docx"]:
            p = os.path.join(templates_dir, fn)
            if os.path.exists(p):
                try:
                    d = docx.Document(p)
                    for r in d.part.rels.values():
                        if "image" in str(r.target_ref):
                            with open(target, "wb") as f:
                                f.write(r.target_part.blob)
                            return target
                except Exception:
                    pass
    return target

def generate_report_file(templates_dir, output_path, data):
    """
    Generate Word report docx file adhering strictly to layout specifications.
    """
    image_count = int(data.get("image_count", 1))
    placeholder_img = ensure_placeholder_image(templates_dir)

    # Base template
    if image_count == 1:
        base_tmpl = os.path.join(templates_dir, "图片数等于1.docx")
    elif image_count == 2:
        base_tmpl = os.path.join(templates_dir, "图片数等于2.docx")
    else:
        base_tmpl = os.path.join(templates_dir, "图片数大于等于3.docx")

    if not os.path.exists(base_tmpl):
        raise FileNotFoundError(f"未找到模板文件: {base_tmpl}")

    doc = docx.Document(base_tmpl)
    table0 = doc.tables[0]

    name = data.get("name", "")
    gender = data.get("gender", "")
    age = str(data.get("age", ""))
    if age and not age.endswith(("岁", "月", "日", "天")):
        age += "岁"
    pathology_no = data.get("pathology_no", "")
    sub_date = data.get("submission_date", "")
    inpatient_no = data.get("inpatient_no", "")
    outpatient_no = data.get("outpatient_no", "")
    ref_doc = data.get("referring_doctor", "")
    diag = data.get("clinical_diagnosis", "")

    def format_cell(cell, text):
        cell.text = text
        if cell.paragraphs and cell.paragraphs[0].runs:
            set_run_font(cell.paragraphs[0].runs[0], "宋体", 12)

    # Row 0: 姓名、性别、年龄、病理号
    format_cell(table0.rows[0].cells[0], f"姓名：{name}")
    format_cell(table0.rows[0].cells[1], f"性别：{gender}")
    format_cell(table0.rows[0].cells[2], f"年龄：{age}")
    format_cell(table0.rows[0].cells[3], f"病理号：{pathology_no}")

    # Row 1: 送检单位、科室、送检日期
    send_unit = str(data.get("send_unit") or "本院").strip() or "本院"
    department = str(data.get("department") or "皮肤科").strip() or "皮肤科"
    format_cell(table0.rows[1].cells[0], f"送检单位：{send_unit}")
    format_cell(table0.rows[1].cells[1], f"科室：{department}")
    format_cell(table0.rows[1].cells[2], f"送检日期：{sub_date}")
    if len(table0.rows[1].cells) > 3 and table0.rows[1].cells[3] != table0.rows[1].cells[2]:
        format_cell(table0.rows[1].cells[3], f"送检日期：{sub_date}")

    # Row 2: 住院号、门诊号、送检医生
    format_cell(table0.rows[2].cells[0], f"住院号：{inpatient_no}")
    format_cell(table0.rows[2].cells[1], f"门诊号：{outpatient_no}")
    format_cell(table0.rows[2].cells[2], f"送检医生：{ref_doc}")
    if len(table0.rows[2].cells) > 3 and table0.rows[2].cells[3] != table0.rows[2].cells[2]:
        format_cell(table0.rows[2].cells[3], f"送检医生：{ref_doc}")

    # Row 3: 临床诊断 & 送检部位 (仅1图时送检部位在表0)
    raw_sites = data.get("biopsy_sites", [])
    if isinstance(raw_sites, str):
        biopsy_sites = [s.strip() for s in raw_sites.split(";") if s.strip()]
    elif isinstance(raw_sites, list):
        biopsy_sites = [str(s).strip() for s in raw_sites if str(s).strip()]
    else:
        biopsy_sites = []

    if image_count == 1:
        single_site = biopsy_sites[0] if biopsy_sites else ""
        format_cell(table0.rows[3].cells[0], f"临床诊断：{diag}")
        if len(table0.rows[3].cells) > 1 and table0.rows[3].cells[1] != table0.rows[3].cells[0]:
            format_cell(table0.rows[3].cells[1], f"临床诊断：{diag}")
        format_cell(table0.rows[3].cells[2], f"送检部位：{single_site}")
        if len(table0.rows[3].cells) > 3 and table0.rows[3].cells[3] != table0.rows[3].cells[2]:
            format_cell(table0.rows[3].cells[3], f"送检部位：{single_site}")
    else:
        for c in table0.rows[3].cells:
            format_cell(c, f"临床诊断：{diag}")

    # Ensure Table 0 has top and bottom borders (three-line table style) and remove redundant lines above it
    remove_extra_separator_above_table(doc, table0)
    set_table0_borders(table0)

    # Locate key section paragraphs
    body = doc._body._element
    p_micro_elem = None
    p_diag_elem = None
    p_sig_elem = None
    p_hist_elem = None

    for p in doc.paragraphs:
        t = p.text.strip()
        if "镜下描述：" in t and p_micro_elem is None:
            p_micro_elem = p._p
        elif "病理诊断：" in t and p_diag_elem is None:
            p_diag_elem = p._p
        elif "报告医生：" in t and "审核医生：" in t and p_sig_elem is None:
            p_sig_elem = p._p
        elif t.startswith("病史：") and p_hist_elem is None:
            p_hist_elem = p._p

    # Remove legacy floating elements between Table 0 and "镜下描述："
    table0_elem = table0._tbl
    elements_to_remove = []
    found_table = False
    for child in list(body):
        if child == table0_elem:
            found_table = True
            continue
        if found_table:
            if child == p_micro_elem:
                break
            elements_to_remove.append(child)

    for elem in elements_to_remove:
        body.remove(elem)

    # Clean up empty paragraphs between sections
    for p in list(doc.paragraphs):
        elem = p._p
        if elem in [p_micro_elem, p_diag_elem, p_sig_elem, p_hist_elem]:
            continue
        if not p.text.strip():
            try:
                elem.getparent().remove(elem)
            except Exception:
                pass

    # Insert Image Section right before "镜下描述："
    selected_images = data.get("selected_images") or []
    # Normalize selected_images into a list of file paths indexed 0..image_count-1
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

    if image_count == 1:
        p_img = doc.add_paragraph()
        p_micro_elem.addprevious(p_img._p)
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(4)
        p_img.paragraph_format.space_after = Pt(2)
        r = p_img.add_run()
        img_to_insert = img_paths[0] if (img_paths and img_paths[0] and os.path.exists(img_paths[0])) else placeholder_img
        if os.path.exists(img_to_insert):
            r.add_picture(img_to_insert, width=Pt(185), height=Pt(139))
    else:
        # Table Layout: 2 columns for 2 images; 3 columns for 3 to 9 images
        cols = 2 if image_count == 2 else 3
        col_width_pt = 260 if cols == 2 else 174
        if cols == 2:
            img_w = 175
            img_h = 131
        elif image_count <= 6:
            img_w = 145
            img_h = 109
        else:
            # 7-9 images: compact height so 3 rows fit cleanly on Page 1
            img_w = 135
            img_h = 100

        num_rows_needed = (image_count + cols - 1) // cols

        img_table = doc.add_table(rows=0, cols=cols)
        p_micro_elem.addprevious(img_table._tbl)
        img_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        remove_table_borders(img_table)

        for r_idx in range(num_rows_needed):
            img_row = img_table.add_row()
            lbl_row = img_table.add_row()

            for c_idx in range(cols):
                img_idx = r_idx * cols + c_idx + 1
                cell_img = img_row.cells[c_idx]
                cell_lbl = lbl_row.cells[c_idx]
                cell_img.width = Pt(col_width_pt)
                cell_lbl.width = Pt(col_width_pt)
                set_cell_margins(cell_img, top=10, bottom=4, left=8, right=8)
                set_cell_margins(cell_lbl, top=4, bottom=8, left=8, right=8)

                p_i = cell_img.paragraphs[0]
                p_i.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_i.paragraph_format.space_before = Pt(0)
                p_i.paragraph_format.space_after = Pt(0)
                p_i.paragraph_format.line_spacing = 1.0

                p_l = cell_lbl.paragraphs[0]
                p_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_l.paragraph_format.space_before = Pt(0)
                p_l.paragraph_format.space_after = Pt(2)
                p_l.paragraph_format.line_spacing = 1.0

                if img_idx <= image_count:
                    # Insert image (actual or placeholder)
                    run_i = p_i.add_run()
                    chosen_p = img_paths[img_idx - 1] if (img_idx - 1 < len(img_paths) and img_paths[img_idx - 1]) else None
                    img_to_insert = chosen_p if (chosen_p and os.path.exists(chosen_p)) else placeholder_img
                    if os.path.exists(img_to_insert):
                        run_i.add_picture(img_to_insert, width=Pt(img_w), height=Pt(img_h))

                    # Insert site label directly beneath
                    s_name = biopsy_sites[img_idx - 1] if len(biopsy_sites) >= img_idx else ""
                    label_str = f"{img_idx}. {s_name}" if s_name else f"{img_idx}."
                    run_l = p_l.add_run(label_str)
                    set_run_font(run_l, "宋体", 11)
                else:
                    # Omit placeholder for unused slots (e.g. slot 5, 6 when image_count=4)
                    pass

    # Ensure empty line between image section and "镜下描述：" (Requirement 1: 行距固定15磅，图片下固定为一行)
    p_img_spacer = doc.add_paragraph()
    p_micro_elem.addprevious(p_img_spacer._p)
    p_img_spacer.paragraph_format.space_before = Pt(0)
    p_img_spacer.paragraph_format.space_after = Pt(0)
    p_img_spacer.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p_img_spacer.paragraph_format.line_spacing = Pt(15)

    # Reference target lines per image specification:
    # 1.1. 单图: 镜下描述 15行 (含标题), 病理诊断 10行 (含标题)
    # 1.2. 双图: 镜下描述 14行 (含标题), 病理诊断 10行 (含标题)
    # 1.3. 三图: 镜下描述 15行 (含标题), 病理诊断 10行 (含标题)
    # 1.4. 4-6图 (两行): 镜下描述 10行 (含标题), 病理诊断 6行 (含标题)
    # 1.5. 7-9图 (三行): 镜下描述 5行 (含标题), 病理诊断 3行 (含标题)
    if image_count == 1:
        target_micro_lines, target_diag_lines = 15, 10
    elif image_count == 2:
        target_micro_lines, target_diag_lines = 14, 10
    elif image_count == 3:
        target_micro_lines, target_diag_lines = 15, 10
    elif image_count <= 6:
        target_micro_lines, target_diag_lines = 10, 6
    else:
        target_micro_lines, target_diag_lines = 5, 3

    # Format 镜下描述 paragraph
    p_micro_obj = docx.text.paragraph.Paragraph(p_micro_elem, doc)
    p_micro_obj.text = "镜下描述："
    p_micro_obj.paragraph_format.space_before = Pt(0)
    p_micro_obj.paragraph_format.space_after = Pt(0)
    p_micro_obj.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p_micro_obj.paragraph_format.line_spacing = Pt(15)
    if p_micro_obj.runs:
        set_run_font(p_micro_obj.runs[0], "宋体", 12)

    # Populate microscopic description content lines
    raw_micro = str(data.get("microscopic_desc") or "").strip()
    if raw_micro.startswith("镜下描述："):
        raw_micro = raw_micro[len("镜下描述："):].strip()
    elif raw_micro.startswith("镜下描述:"):
        raw_micro = raw_micro[len("镜下描述:"):].strip()

    desc_lines = [l.strip() for l in raw_micro.splitlines() if l.strip()]
    micro_content_visual_lines = 0
    for line in desc_lines:
        p_desc = doc.add_paragraph()
        p_diag_elem.addprevious(p_desc._p)
        p_desc.paragraph_format.space_before = Pt(0)
        p_desc.paragraph_format.space_after = Pt(0)
        p_desc.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        p_desc.paragraph_format.line_spacing = Pt(15)
        r = p_desc.add_run(line)
        set_run_font(r, "宋体", 12)
        micro_content_visual_lines += count_visual_lines(line)

    total_micro_lines = 1 + micro_content_visual_lines
    pad_micro_lines = max(0, target_micro_lines - total_micro_lines)
    for _ in range(pad_micro_lines):
        p_pad = doc.add_paragraph()
        p_diag_elem.addprevious(p_pad._p)
        p_pad.paragraph_format.space_before = Pt(0)
        p_pad.paragraph_format.space_after = Pt(0)
        p_pad.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        p_pad.paragraph_format.line_spacing = Pt(15)

    # Format 病理诊断 paragraph (Requirement 1: 换行后写诊断内容)
    p_diag_obj = docx.text.paragraph.Paragraph(p_diag_elem, doc)
    p_diag_obj.text = "病理诊断："
    p_diag_obj.paragraph_format.space_before = Pt(0)
    p_diag_obj.paragraph_format.space_after = Pt(0)
    p_diag_obj.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p_diag_obj.paragraph_format.line_spacing = Pt(15)
    if p_diag_obj.runs:
        set_run_font(p_diag_obj.runs[0], "宋体", 12)

    raw_diag = str(data.get("pathology_diag") or "").strip()
    if raw_diag.startswith("病理诊断："):
        raw_diag = raw_diag[len("病理诊断："):].strip()
    elif raw_diag.startswith("病理诊断:"):
        raw_diag = raw_diag[len("病理诊断:"):].strip()

    diag_lines = [l.strip() for l in raw_diag.splitlines() if l.strip()]
    if not diag_lines and raw_diag:
        diag_lines = [raw_diag]

    diag_content_visual_lines = 0
    for line in diag_lines:
        p_diag_line = doc.add_paragraph()
        p_sig_elem.addprevious(p_diag_line._p)
        p_diag_line.paragraph_format.space_before = Pt(0)
        p_diag_line.paragraph_format.space_after = Pt(0)
        p_diag_line.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        p_diag_line.paragraph_format.line_spacing = Pt(15)
        r = p_diag_line.add_run(line)
        set_run_font(r, "宋体", 12)
        diag_content_visual_lines += count_visual_lines(line)

    total_diag_lines = 1 + diag_content_visual_lines
    pad_diag_lines = max(0, target_diag_lines - total_diag_lines)
    for _ in range(pad_diag_lines):
        p_pad = doc.add_paragraph()
        p_sig_elem.addprevious(p_pad._p)
        p_pad.paragraph_format.space_before = Pt(0)
        p_pad.paragraph_format.space_after = Pt(0)
        p_pad.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        p_pad.paragraph_format.line_spacing = Pt(15)

    # Format 报告医生 / 审核医生 / 报告日期 signature paragraph
    p_sig_obj = docx.text.paragraph.Paragraph(p_sig_elem, doc)
    rep_doc = data.get("report_doctor") or "张医生"
    rev_doc = data.get("review_doctor") or "李主任"
    rep_date = data.get("report_date")
    if not rep_date:
        import datetime
        rep_date = datetime.date.today().strftime("%Y年%m月%d日")
    p_sig_obj.text = f"报告医生：{rep_doc}                审核医生：{rev_doc}                报告日期：{rep_date}"
    p_sig_obj.paragraph_format.space_before = Pt(0)
    p_sig_obj.paragraph_format.space_after = Pt(0)
    p_sig_obj.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p_sig_obj.paragraph_format.line_spacing = Pt(15)
    if p_sig_obj.runs:
        set_run_font(p_sig_obj.runs[0], "宋体", 12)

    # Format 病史 paragraph
    p_hist_obj = docx.text.paragraph.Paragraph(p_hist_elem, doc)
    med_hist = data.get("medical_history", "")
    p_hist_obj.text = f"病史：{med_hist}"
    p_hist_obj.paragraph_format.space_before = Pt(4)
    p_hist_obj.paragraph_format.space_after = Pt(4)
    if p_hist_obj.runs:
        set_run_font(p_hist_obj.runs[0], "宋体", 12)

    # Insert explicit page break before "病史：" to guarantee it is on Page 2
    p_pb = doc.add_paragraph()
    p_hist_elem.addprevious(p_pb._p)
    p_pb.paragraph_format.space_before = Pt(0)
    p_pb.paragraph_format.space_after = Pt(0)
    r_pb = p_pb.add_run()
    r_pb.add_break(docx.enum.text.WD_BREAK.PAGE)

    dir_p = os.path.dirname(output_path)
    if dir_p:
        os.makedirs(dir_p, exist_ok=True)

    if os.path.exists(output_path):
        try:
            with open(output_path, "a+b"):
                pass
        except (PermissionError, IOError) as pe:
            raise PermissionError(
                f"报告文件被其他软件（如 Word 或 WPS）打开占用，无法写入！请先关闭该文件后重试保存。路径: {os.path.basename(output_path)}"
            ) from pe

    try:
        doc.save(output_path)
    except (PermissionError, IOError) as pe:
        raise PermissionError(
            f"保存报告文件失败（文件可能被 Word/WPS 打开占用或权限不足）: {os.path.basename(output_path)}"
        ) from pe
    return output_path

def update_report_file(templates_dir, doc_path, data):
    """
    Update report file. Uses generate_report_file to ensure layout consistency.
    """
    return generate_report_file(templates_dir, doc_path, data)
