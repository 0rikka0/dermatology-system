# -*- coding: utf-8 -*-
"""
Dermatology Pathology Reporting System - Signing Sheet Generator
Generates report pickup / signing sheet Excel files adhering strictly to '报告领取签名单示例.xlsx'.
Outputs to: 工作目录/领取表/
Also provides HTML table preview and printing capabilities.
"""
import os
import html
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

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

    today_str = datetime.now().strftime("%Y%m%d")
    now_time_str = datetime.now().strftime("%H%M%S")

    # Determine filename
    if custom_filename:
        filename = custom_filename if custom_filename.endswith(".xlsx") else f"{custom_filename}.xlsx"
    else:
        date_part = ""
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

        filename = f"报告领取签名单_{date_part}{pfx_tag}_{now_time_str}.xlsx"

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

    wb.save(out_path)
    return {
        "success": True,
        "file_path": out_path,
        "filename": filename,
        "count": len(patients)
    }

def render_signing_sheet_html(patients, title=None, date_range_str=None):
    """
    Render signing sheet as high-fidelity HTML for WebUI preview and browser printing.
    """
    total = len(patients)
    date_info = html.escape(date_range_str or datetime.now().strftime("%Y年%m月%d日"))
    sheet_title = html.escape(title or "皮肤组织病理报告领取签名单")

    rows_html = []
    for idx, p in enumerate(patients, start=1):
        p_no = html.escape(str(p.get("pathology_no", "")))
        p_name = html.escape(str(p.get("name", "")))
        p_diag = html.escape(str(p.get("pathology_diag") or p.get("clinical_diagnosis") or ""))

        rows_html.append(f"""
        <tr>
            <td style="text-align: center; width: 50px;">{idx}</td>
            <td style="text-align: center; font-family: monospace; font-weight: 600; width: 140px;">{p_no}</td>
            <td style="text-align: center; font-weight: 600; width: 110px;">{p_name}</td>
            <td style="width: 160px; min-height: 32px;"></td>
        </tr>
        """)

    # Fill extra empty rows to reach at least 25 rows for a full page look
    min_rows = max(25, total)
    for idx in range(total + 1, min_rows + 1):
        rows_html.append(f"""
        <tr>
            <td style="text-align: center; color: #cbd5e1;">{idx}</td>
            <td></td>
            <td></td>
            <td style="min-height: 32px;"></td>
        </tr>
        """)

    html_out = f"""
    <div class="signing-sheet-container">
        <div class="signing-sheet-header">
            <h2 class="signing-hospital">某三甲医院皮肤科</h2>
            <h1 class="signing-title">{sheet_title}</h1>
            <div class="signing-meta">
                <span><strong>统计范围：</strong>{date_info}</span>
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
