# -*- coding: utf-8 -*-
"""
Dermatology Pathology Reporting System - Signing Sheet Generator
Generates report pickup / signing sheet Excel files adhering strictly to '报告领取签名单示例.xlsx'.
Outputs to: 工作目录/领取表/
Also provides HTML table preview and printing capabilities.
"""
import os
import html
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from app.path_manager import backup_before_overwrite

SURGERY_DATE_NOTICE = "本手术日期仅供参考，若本页未找到可在相近日期的页面中寻找"

def get_previous_workday(date_str):
    """Return the previous weekday (Mon-Fri) before date_str (YYYY-MM-DD)."""
    start_s, _ = get_surgery_date_bounds(date_str)
    return start_s


def get_surgery_date_bounds(date_str):
    """
    For a submission date (YYYY-MM-DD), return (start_str, end_str) of the surgery window:
    - When crossing a weekend (previous workday is Friday, e.g. Monday submission),
      includes the weekend days up to the day before submission (Fri ~ Sun).
    - When not crossing a weekend (Tue-Fri submission), both bounds equal the single previous workday.
    """
    if not date_str:
        return "", ""
    try:
        dt = datetime.strptime(str(date_str).strip()[:10], "%Y-%m-%d").date()
    except ValueError:
        return "", ""
    prev_wd = dt - timedelta(days=1)
    while prev_wd.weekday() >= 5:  # 5=Saturday, 6=Sunday
        prev_wd -= timedelta(days=1)
    prev_cal = dt - timedelta(days=1)
    end_dt = prev_cal if prev_cal > prev_wd else prev_wd
    return prev_wd.strftime("%Y-%m-%d"), end_dt.strftime("%Y-%m-%d")


def compute_surgery_date_str(start_date=None, end_date=None, patients=None):
    """Compute surgery date string, expanding Friday previous workday to include weekend."""
    s = (start_date or "").strip()
    e = (end_date or "").strip()
    if s and e:
        ps, _ = get_surgery_date_bounds(s)
        _, pe = get_surgery_date_bounds(e)
        if ps and pe:
            return ps if ps == pe else f"{ps} 至 {pe}"
    elif s:
        ps, _ = get_surgery_date_bounds(s)
        if ps:
            return f"{ps} 起"
    elif e:
        _, pe = get_surgery_date_bounds(e)
        if pe:
            return f"截至 {pe}"

    if patients:
        raw_dates = sorted({
            str(p.get("submission_date_raw") or "").strip()[:10]
            for p in patients
            if str(p.get("submission_date_raw") or "").strip()
        })
        if raw_dates:
            ps, _ = get_surgery_date_bounds(raw_dates[0])
            _, pe = get_surgery_date_bounds(raw_dates[-1])
            if ps and pe:
                return ps if ps == pe else f"{ps} 至 {pe}"
    return "-"

def _build_signing_filename(ext, start_date=None, end_date=None, prefixes=None, custom_filename=None, timestamp_str=None):
    today_str = datetime.now().strftime("%Y%m%d")
    now_time_str = timestamp_str or datetime.now().strftime("%H%M%S")
    dot_ext = f".{ext.lstrip('.')}"

    if custom_filename:
        stem = os.path.splitext(custom_filename)[0]
        return f"{stem}{dot_ext}"

    if start_date and end_date:
        s_clean = start_date.replace("-", "")
        e_clean = end_date.replace("-", "")
        date_part = f"{s_clean}_{e_clean}" if s_clean != e_clean else s_clean
    elif start_date:
        date_part = start_date.replace("-", "")
    elif end_date:
        date_part = end_date.replace("-", "")
    else:
        date_part = today_str

    pfx_tag = ""
    if prefixes and "ALL" not in prefixes and "所有" not in prefixes:
        clean_p = [p for p in prefixes if p]
        if clean_p:
            pfx_tag = f"_{''.join(clean_p)}"

    return f"报告领取签名单_{date_part}{pfx_tag}_{now_time_str}{dot_ext}"

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SAMPLE_TEMPLATE = os.path.join(BASE_DIR, "报告领取签名单示例.xlsx")

def get_thin_border():
    thin = Side(border_style="thin", color="000000")
    return Border(top=thin, bottom=thin, left=thin, right=thin)

def generate_signing_sheet_excel(work_dir, patients, start_date=None, end_date=None, prefixes=None, custom_filename=None):
    """
    Generate signing sheet Excel workbook matching '报告领取签名单示例.xlsx'.
    Saves in work_dir/领取表/.
    """
    output_dir = os.path.join(work_dir, "领取表")
    os.makedirs(output_dir, exist_ok=True)

    filename = _build_signing_filename(
        "xlsx",
        start_date=start_date,
        end_date=end_date,
        prefixes=prefixes,
        custom_filename=custom_filename
    )

    out_path = os.path.join(output_dir, filename)

    # Load or create workbook
    if os.path.exists(SAMPLE_TEMPLATE):
        wb = openpyxl.load_workbook(SAMPLE_TEMPLATE)
        ws = wb.active
        # Clear rows 2 to max_row
        for r in range(2, max(ws.max_row + 1, 35)):
            for c in range(1, 4):
                ws.cell(r, c).value = None
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.column_dimensions['A'].width = 18.0
        ws.column_dimensions['B'].width = 18.0
        ws.column_dimensions['C'].width = 22.0

        # Row 1 header
        headers = ["病理号", "姓名", "签名"]
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(1, c_idx, h)
            cell.font = Font(name="等线", size=11, bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = get_thin_border()
        ws.row_dimensions[1].height = 28

    thin_border = get_thin_border()
    data_font = Font(name="等线", size=11)
    center_align = Alignment(horizontal="center", vertical="center")

    current_row = 2
    for p in patients:
        p_no = str(p.get("pathology_no", "")).strip()
        p_name = str(p.get("name", "")).strip()

        # Cell A: 病理号
        cell_a = ws.cell(current_row, 1, p_no)
        cell_a.font = data_font
        cell_a.alignment = center_align
        cell_a.border = thin_border

        # Cell B: 姓名
        cell_b = ws.cell(current_row, 2, p_name)
        cell_b.font = data_font
        cell_b.alignment = center_align
        cell_b.border = thin_border

        # Cell C: 签名 (留空手签)
        cell_c = ws.cell(current_row, 3, "")
        cell_c.font = data_font
        cell_c.alignment = center_align
        cell_c.border = thin_border

        ws.row_dimensions[current_row].height = 25
        current_row += 1

    # Pre-populate empty rows up to row 32 (standard A4 page size in template)
    while current_row <= 32:
        for col in range(1, 4):
            cell = ws.cell(current_row, col, "")
            cell.font = data_font
            cell.alignment = center_align
            cell.border = thin_border
        ws.row_dimensions[current_row].height = 25
        current_row += 1

    # Page setup
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    if os.path.exists(out_path):
        backup_before_overwrite(out_path, lock_label="签名单 Excel 文件")

    try:
        wb.save(out_path)
    except (PermissionError, IOError) as pe:
        raise PermissionError(
            f"保存签名单 Excel 文件失败（可能被 Excel/WPS 打开占用或权限不足）: {os.path.basename(out_path)}"
        ) from pe

    return {
        "success": True,
        "file_path": out_path,
        "filename": filename,
        "count": len(patients)
    }

def _set_doc_run_font(run, font_name="宋体", size_pt=11, bold=False, color=None):
    run.font.name = font_name
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:eastAsia"), font_name)
    run.font.size = Pt(size_pt)
    run.bold = bold
    if color is not None:
        run.font.color.rgb = color

def _set_cell_borders_all(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    existing = tcPr.find(qn("w:tcBorders"))
    if existing is not None:
        tcPr.remove(existing)
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
        f'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
        f'  <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
        f'  <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)

def generate_signing_sheet_word(work_dir, patients, start_date=None, end_date=None, prefixes=None, custom_filename=None, date_range_str=None):
    """
    Generate signing sheet Word (.docx) file in work_dir/领取表/.
    - Removes all empty table rows (only header + patient rows).
    - Includes 手术日期 (previous workday of submission date, prominently highlighted with notice)
      and 送检日期 metadata.
    """
    output_dir = os.path.join(work_dir, "领取表")
    os.makedirs(output_dir, exist_ok=True)

    filename = _build_signing_filename(
        "docx",
        start_date=start_date,
        end_date=end_date,
        prefixes=prefixes,
        custom_filename=custom_filename
    )
    out_path = os.path.join(output_dir, filename)

    doc = docx.Document()
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(1.8)
        section.right_margin = Cm(1.8)

    # Header titles
    p_hosp = doc.add_paragraph()
    p_hosp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_hosp.paragraph_format.space_before = Pt(0)
    p_hosp.paragraph_format.space_after = Pt(2)
    _set_doc_run_font(p_hosp.add_run("某三甲医院皮肤科"), "宋体", 14, bold=True)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(8)
    _set_doc_run_font(p_title.add_run("皮肤组织病理报告领取签名单"), "宋体", 18, bold=True)

    # Prominent Surgery Date Box (5.3)
    surgery_date_str = compute_surgery_date_str(start_date=start_date, end_date=end_date, patients=patients)
    p_surg = doc.add_paragraph()
    p_surg.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_surg.paragraph_format.space_before = Pt(2)
    p_surg.paragraph_format.space_after = Pt(2)
    pPr = p_surg._p.get_or_add_pPr()
    pPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="FEF3C7"/>'))
    pPr.append(parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'  <w:top w:val="single" w:sz="8" w:space="4" w:color="D97706"/>'
        f'  <w:bottom w:val="single" w:sz="8" w:space="4" w:color="D97706"/>'
        f'  <w:left w:val="single" w:sz="8" w:space="6" w:color="D97706"/>'
        f'  <w:right w:val="single" w:sz="8" w:space="6" w:color="D97706"/>'
        f'</w:pBdr>'
    ))
    _set_doc_run_font(p_surg.add_run(f"手术日期：{surgery_date_str}"), "宋体", 14, bold=True, color=RGBColor(0xB9, 0x1C, 0x1C))
    p_surg.add_run("\n")
    _set_doc_run_font(p_surg.add_run(f"（提示：{SURGERY_DATE_NOTICE}）"), "宋体", 10, bold=True, color=RGBColor(0x92, 0x40, 0x0E))

    # Submission Date & Metadata Line (5.1)
    if not date_range_str:
        if start_date and end_date:
            date_range_str = start_date if start_date == end_date else f"{start_date} 至 {end_date}"
        else:
            date_range_str = start_date or end_date or datetime.now().strftime("%Y-%m-%d")

    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_meta.paragraph_format.space_before = Pt(6)
    p_meta.paragraph_format.space_after = Pt(6)
    meta_text = f"送检日期：{date_range_str}        报告份数：{len(patients)} 份        打印日期：{datetime.now().strftime('%Y-%m-%d %H:%M')}"
    _set_doc_run_font(p_meta.add_run(meta_text), "宋体", 10.5, bold=False)

    # Signing Table (strictly no empty rows - 5.2)
    table = doc.add_table(rows=1 + len(patients), cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    col_widths = [Cm(1.6), Cm(4.8), Cm(4.2), Cm(6.4)]
    headers = ["序号", "病理号", "姓名", "领取人签名"]

    hdr_row = table.rows[0]
    trPr = hdr_row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:tblHeader"))
    trPr.append(OxmlElement("w:cantSplit"))
    trPr.append(parse_xml(f'<w:trHeight {nsdecls("w")} w:val="420" w:hRule="atLeast"/>'))

    for c_idx, h_text in enumerate(headers):
        cell = hdr_row.cells[c_idx]
        cell.width = col_widths[c_idx]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _set_cell_borders_all(cell)
        cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="F1F5F9"/>'))
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        _set_doc_run_font(p.add_run(h_text), "宋体", 11, bold=True)

    for idx, pat in enumerate(patients, start=1):
        row = table.rows[idx]
        r_trPr = row._tr.get_or_add_trPr()
        r_trPr.append(OxmlElement("w:cantSplit"))
        r_trPr.append(parse_xml(f'<w:trHeight {nsdecls("w")} w:val="400" w:hRule="atLeast"/>'))

        vals = [
            str(idx),
            str(pat.get("pathology_no", "")).strip(),
            str(pat.get("name", "")).strip(),
            ""
        ]
        for c_idx, val in enumerate(vals):
            cell = row.cells[c_idx]
            cell.width = col_widths[c_idx]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            _set_cell_borders_all(cell)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            _set_doc_run_font(p.add_run(val), "宋体", 11, bold=False)

    # Footer
    p_foot = doc.add_paragraph()
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_foot.paragraph_format.space_before = Pt(12)
    p_foot.paragraph_format.space_after = Pt(0)
    _set_doc_run_font(p_foot.add_run("经手人：____________            发放日期：______年____月____日"), "宋体", 10.5)

    if os.path.exists(out_path):
        backup_before_overwrite(out_path, lock_label="签名单 Word 文件")

    try:
        doc.save(out_path)
    except (PermissionError, IOError) as pe:
        raise PermissionError(
            f"保存签名单 Word 文件失败（可能被 Word/WPS 打开占用或权限不足）: {os.path.basename(out_path)}"
        ) from pe

    return {
        "success": True,
        "file_path": out_path,
        "filename": filename,
        "surgery_date": surgery_date_str,
        "count": len(patients)
    }

def render_signing_sheet_html(patients, title=None, date_range_str=None, surgery_date_str=None):
    """
    Render signing sheet as high-fidelity HTML for WebUI preview and browser printing.
    """
    total = len(patients)
    date_info = html.escape(date_range_str or datetime.now().strftime("%Y-%m-%d"))
    surgery_info = html.escape(surgery_date_str or compute_surgery_date_str(patients=patients))
    sheet_title = html.escape(title or "皮肤组织病理报告领取签名单")

    rows_html = []
    for idx, p in enumerate(patients, start=1):
        p_no = html.escape(str(p.get("pathology_no", "")))
        p_name = html.escape(str(p.get("name", "")))

        rows_html.append(f"""
        <tr>
            <td style="text-align: center; width: 50px;">{idx}</td>
            <td style="text-align: center; font-family: monospace; font-weight: 600; width: 140px;">{p_no}</td>
            <td style="text-align: center; font-weight: 600; width: 110px;">{p_name}</td>
            <td style="width: 160px; min-height: 32px;"></td>
        </tr>
        """)

    html_out = f"""
    <div class="signing-sheet-container">
        <div class="signing-sheet-header">
            <h2 class="signing-hospital">某三甲医院皮肤科</h2>
            <h1 class="signing-title">{sheet_title}</h1>
            <div class="signing-surgery-banner">
                <div class="signing-surgery-date">手术日期：{surgery_info}</div>
                <div class="signing-surgery-notice">提示：{SURGERY_DATE_NOTICE}</div>
            </div>
            <div class="signing-meta">
                <span><strong>送检日期：</strong>{date_info}</span>
                <span><strong>报告份数：</strong>{total} 份</span>
                <span><strong>打印日期：</strong>{datetime.now().strftime('%Y-%m-%d %H:%M')}</span>
            </div>
        </div>

        <table class="signing-sheet-table">
            <thead>
                <tr>
                    <th style="width: 50px;">序号</th>
                    <th style="width: 140px;">病理号</th>
                    <th style="width: 110px;">姓名</th>
                    <th style="width: 160px;">领取人签名</th>
                </tr>
            </thead>
            <tbody>
                {''.join(rows_html)}
            </tbody>
        </table>

        <div class="signing-sheet-footer">
            <span>经手人：____________</span>
            <span>发放日期：______年____月____日</span>
            <span>第 1 页 / 共 1 页</span>
        </div>
    </div>
    """
    return html_out
