# -*- coding: utf-8 -*-
"""
皮肤病理报告系统 - 快速入门操作手册 PDF 生成脚本 (v6.0 标准 SOP 版)
定位：纯临床实操标准作业程序 (SOP)，聚焦“需要做什么、怎么做、不能做什么”，
      语言通俗简练，配系统流程图与脱敏 WebUI 实景标注截图。
"""

import argparse
import os
import sys

from PIL import Image as PILImage, ImageDraw, ImageFont
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    HRFlowable,
    Image,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

# ================= 1. 字体与配色 =================
FONT_REG = "C:/Windows/Fonts/msyh.ttc"
FONT_BOLD = "C:/Windows/Fonts/msyhbd.ttc"
if not os.path.exists(FONT_REG):
    FONT_REG = "C:/Windows/Fonts/simhei.ttf"
    FONT_BOLD = FONT_REG

pdfmetrics.registerFont(TTFont("YaHei", FONT_REG))
pdfmetrics.registerFont(TTFont("YaHeiBold", FONT_BOLD))

NAVY = colors.HexColor("#1E3A5F")
BLUE = colors.HexColor("#2563EB")
LIGHT_BLUE = colors.HexColor("#EFF6FF")
GREEN = colors.HexColor("#15803D")
LIGHT_GREEN = colors.HexColor("#F0FDF4")
RED = colors.HexColor("#B91C1C")
LIGHT_RED = colors.HexColor("#FEF2F2")
AMBER = colors.HexColor("#B45309")
LIGHT_AMBER = colors.HexColor("#FFFBEB")
DARK_TEXT = colors.HexColor("#1F2937")
GRAY_TEXT = colors.HexColor("#4B5563")
BORDER_GRAY = colors.HexColor("#CBD5E1")
ROW_ALT = colors.HexColor("#F8FAFC")


def _pil_font(size, bold=False):
    path = FONT_BOLD if bold else FONT_REG
    try:
        return ImageFont.truetype(path, size)
    except Exception:
        return ImageFont.load_default()


# ================= 2. 矢量级简明示意图生成 =================
def create_workflow_diagram(out_path):
    """生成 6 步全流程极简看板图"""
    W, H = 1600, 360
    img = PILImage.new("RGB", (W, H), "#FFFFFF")
    draw = ImageDraw.Draw(img)

    f_title = _pil_font(26, bold=True)
    f_step = _pil_font(22, bold=True)
    f_desc = _pil_font(18, bold=False)

    draw.text((40, 20), "日常工作 6 步极简总览（按从左到右顺序完成）", fill="#1E3A5F", font=f_title)

    steps = [
        ("1. 点图标进入", "双击桌面快捷方式\n自动拉起网页\n(有后台不重开)", "#2563EB", "#EFF6FF"),
        ("2. 登记建档", "选前缀点自动编号\n填信息与图数(1-9)\n建初稿与照片夹", "#0284C7", "#F0F9FF"),
        ("3. 放图出初稿", "照片放入对应目录\n勾选常用诊断词条\n生成并预览报告", "#0D9488", "#F0FDFA"),
        ("4. 复核/挂起", "无误直接点归档\n需重切进报告审核\n支持双草稿并行改", "#D97706", "#FFFBEB"),
        ("5. 加做组化", "如加做免疫组化\n进独立组化页签\n生成第二份报告", "#7C3AED", "#F5F3FF"),
        ("6. 输出与收尾", "批量打印/签名单\n历史按日随时查改\n下班点停止系统", "#15803D", "#F0FDF4"),
    ]

    box_w, box_h = 228, 220
    start_x, gap = 36, 30
    y0 = 86

    for idx, (title, body, border_hex, bg_hex) in enumerate(steps):
        x0 = start_x + idx * (box_w + gap)
        x1, y1 = x0 + box_w, y0 + box_h
        draw.rounded_rectangle([x0, y0, x1, y1], radius=14, fill=bg_hex, outline=border_hex, width=3)
        draw.rounded_rectangle([x0, y0, x1, y0 + 50], radius=14, fill=border_hex)
        draw.rectangle([x0, y0 + 30, x1, y0 + 50], fill=border_hex)
        draw.text((x0 + 16, y0 + 11), title, fill="#FFFFFF", font=f_step)
        draw.multiline_text((x0 + 16, y0 + 68), body, fill="#1F2937", font=f_desc, spacing=10)

        if idx < len(steps) - 1:
            ax0, ax1 = x1 + 4, x1 + gap - 4
            ay = y0 + box_h // 2
            draw.line([(ax0, ay), (ax1 - 6, ay)], fill="#64748B", width=4)
            draw.polygon([(ax1, ay), (ax1 - 10, ay - 7), (ax1 - 10, ay + 7)], fill="#64748B")

    img.save(out_path, "PNG")
    return out_path


def create_layout_diagram(out_path):
    """生成 1-9 张图 1 页纸紧凑排版示意图"""
    W, H = 1500, 320
    img = PILImage.new("RGB", (W, H), "#FFFFFF")
    draw = ImageDraw.Draw(img)

    f_title = _pil_font(24, bold=True)
    f_sub = _pil_font(20, bold=True)
    f_small = _pil_font(17, bold=False)

    draw.text((30, 14), "图文排版速览：系统按 1-9 张图自动切换行列比例，尽量保证 1 页 A4 打完", fill="#1E3A5F", font=f_title)

    cards = [
        ("1~3 张图", "1 行横排", (1, 3), "#2563EB"),
        ("4 张图", "2 x 2 田字格", (2, 2), "#0284C7"),
        ("5~6 张图", "3 列 x 2 行", (2, 3), "#0D9488"),
        ("7~9 张图", "3 列 x 3 行", (3, 3), "#7C3AED"),
    ]
    cw, ch = 330, 220
    for idx, (lbl, mode, (rows, cols), col_hex) in enumerate(cards):
        x0 = 30 + idx * (cw + 40)
        y0 = 64
        draw.rounded_rectangle([x0, y0, x0 + cw, y0 + ch], radius=12, fill="#F8FAFC", outline=col_hex, width=3)
        draw.text((x0 + 18, y0 + 12), f"{lbl}：{mode}", fill=col_hex, font=f_sub)

        gx, gy = x0 + 24, y0 + 54
        gw, gh = cw - 48, ch - 74
        cell_w = (gw - (cols - 1) * 8) // cols
        cell_h = (gh - (rows - 1) * 8) // rows
        num = 1
        for r in range(rows):
            for c in range(cols):
                cx0 = gx + c * (cell_w + 8)
                cy0 = gy + r * (cell_h + 8)
                draw.rectangle([cx0, cy0, cx0 + cell_w, cy0 + cell_h], fill="#DBEAFE", outline="#60A5FA", width=2)
                draw.text((cx0 + cell_w // 2 - 10, cy0 + cell_h // 2 - 10), f"图{num}", fill="#1E3A5F", font=f_small)
                num += 1

    img.save(out_path, "PNG")
    return out_path


# ================= 3. PDF 组件构建 =================
def _styles():
    base = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "ManualTitle",
            parent=base["Title"],
            fontName="YaHeiBold",
            fontSize=22,
            leading=28,
            textColor=colors.white,
            alignment=TA_CENTER,
        ),
        "subtitle": ParagraphStyle(
            "ManualSubtitle",
            parent=base["Normal"],
            fontName="YaHei",
            fontSize=10.5,
            leading=15,
            textColor=colors.HexColor("#E2E8F0"),
            alignment=TA_CENTER,
        ),
        "h1": ParagraphStyle(
            "H1",
            parent=base["Heading1"],
            fontName="YaHeiBold",
            fontSize=13.5,
            leading=19,
            textColor=NAVY,
            spaceBefore=6,
            spaceAfter=4,
        ),
        "body": ParagraphStyle(
            "Body",
            parent=base["Normal"],
            fontName="YaHei",
            fontSize=9.5,
            leading=14.5,
            textColor=DARK_TEXT,
            wordWrap="CJK",
        ),
        "sop_head": ParagraphStyle(
            "SopHead",
            parent=base["Normal"],
            fontName="YaHeiBold",
            fontSize=10,
            leading=14,
            textColor=NAVY,
            wordWrap="CJK",
        ),
        "sop_cell": ParagraphStyle(
            "SopCell",
            parent=base["Normal"],
            fontName="YaHei",
            fontSize=9,
            leading=14.2,
            textColor=DARK_TEXT,
            wordWrap="CJK",
        ),
        "caption": ParagraphStyle(
            "Caption",
            parent=base["Normal"],
            fontName="YaHei",
            fontSize=8.5,
            leading=12,
            textColor=GRAY_TEXT,
            alignment=TA_CENTER,
            wordWrap="CJK",
        ),
    }


def _sop_card(st, sop_id, title, need_do, how_to, do_not):
    """生成标准化三栏 SOP 卡片：需要做什么 / 怎么做 / 不能做什么"""
    header = Paragraph(f"<b>{sop_id}｜{title}</b>", ParagraphStyle(
        "CardTitle", fontName="YaHeiBold", fontSize=11, leading=15, textColor=colors.white
    ))
    rows = [
        [header, ""],
        [
            Paragraph("<b>需要做什么</b>", ParagraphStyle("L1", fontName="YaHeiBold", fontSize=9.5, leading=13, textColor=BLUE)),
            Paragraph(need_do, st["sop_cell"]),
        ],
        [
            Paragraph("<b>怎么做</b>", ParagraphStyle("L2", fontName="YaHeiBold", fontSize=9.5, leading=13, textColor=GREEN)),
            Paragraph(how_to, st["sop_cell"]),
        ],
        [
            Paragraph("<b>不能做什么</b>", ParagraphStyle("L3", fontName="YaHeiBold", fontSize=9.5, leading=13, textColor=RED)),
            Paragraph(do_not, st["sop_cell"]),
        ],
    ]
    tbl = Table(rows, colWidths=[26 * mm, 148 * mm])
    tbl.setStyle(TableStyle([
        ("SPAN", (0, 0), (1, 0)),
        ("BACKGROUND", (0, 0), (1, 0), NAVY),
        ("BACKGROUND", (0, 1), (0, 1), LIGHT_BLUE),
        ("BACKGROUND", (1, 1), (1, 1), colors.white),
        ("BACKGROUND", (0, 2), (0, 2), LIGHT_GREEN),
        ("BACKGROUND", (1, 2), (1, 2), colors.white),
        ("BACKGROUND", (0, 3), (0, 3), LIGHT_RED),
        ("BACKGROUND", (1, 3), (1, 3), colors.HexColor("#FFFBFB")),
        ("BOX", (0, 0), (-1, -1), 0.9, NAVY),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, BORDER_GRAY),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
    ]))
    return tbl


def _add_screenshot(story, st, screenshots_dir, filename, caption, max_w_mm=174, max_h_mm=108):
    if not screenshots_dir:
        return
    fpath = os.path.join(screenshots_dir, filename)
    if not os.path.exists(fpath):
        return
    with PILImage.open(fpath) as im:
        pw, ph = im.size
    ratio = min((max_w_mm * mm) / pw, (max_h_mm * mm) / ph)
    w, h = pw * ratio, ph * ratio
    img_flow = Image(fpath, width=w, height=h)
    box = Table([[img_flow], [Paragraph(caption, st["caption"])]], colWidths=[max_w_mm * mm])
    box.setStyle(TableStyle([
        ("BOX", (0, 0), (0, 0), 0.6, BORDER_GRAY),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    story.append(Spacer(1, 3 * mm))
    story.append(box)


def _page_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("YaHei", 8)
    canvas.setFillColor(GRAY_TEXT)
    canvas.drawString(18 * mm, 10 * mm, "皮肤病理报告系统 v6.0｜标准操作手册 (SOP)")
    canvas.drawRightString(A4[0] - 18 * mm, 10 * mm, f"第 {doc.page} 页")
    canvas.restoreState()


def build_manual_pdf(pdf_path, screenshots_dir=None):
    work_dir = os.path.dirname(os.path.abspath(pdf_path))
    img_dir = os.path.join(work_dir, ".temp_imgs")
    os.makedirs(img_dir, exist_ok=True)

    diag_workflow = create_workflow_diagram(os.path.join(img_dir, "diag_sop_workflow.png"))
    diag_layout = create_layout_diagram(os.path.join(img_dir, "diag_sop_layout.png"))

    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=14 * mm,
        bottomMargin=16 * mm,
        title="皮肤病理报告系统 快速入门操作手册 (SOP v6.0)",
    )
    st = _styles()
    story = []

    # -------- 第 1 页：极简封面条 + 全流程总览 + SOP-01 --------
    banner = Table(
        [
            [Paragraph("皮肤病理报告系统 · 快速入门操作手册", st["title"])],
            [Paragraph("版本：v6.0 标准操作指南 (SOP)｜极简上手·一看就会·直接照做", st["subtitle"])],
        ],
        colWidths=[174 * mm],
    )
    banner.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), NAVY),
        ("TOPPADDING", (0, 0), (0, 0), 10),
        ("BOTTOMPADDING", (0, 1), (0, 1), 10),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
    ]))
    story.append(banner)
    story.append(Spacer(1, 4 * mm))

    story.append(Paragraph("本手册只讲三件事：<b>需要做什么</b>、<b>怎么做</b>、<b>不能做什么</b>。日常操作按下方 1→6 顺序执行即可。", st["body"]))
    story.append(Spacer(1, 3 * mm))
    story.append(Image(diag_workflow, width=174 * mm, height=39 * mm))
    story.append(Spacer(1, 4 * mm))

    story.append(_sop_card(
        st,
        "SOP-01",
        "日常开机启动、单图标进入与下班关机",
        need_do="桌面仅保留一个图标启动或回到系统页面，避免黑框干扰或悄悄多开后台。",
        how_to=(
            "<b>1. 首次/再次进入：</b>直接双击桌面快捷方式「<b>启动系统(无黑框后台版).vbs</b>」，浏览器会自动弹出工作台。<br/>"
            "<b>2. 关掉网页后再进：</b>不用找任务栏，再次双击同一个「.vbs」图标，系统检测到后台已在运行，会直接打开网页。<br/>"
            "<b>3. 下班彻底退出：</b>点击网页右上角「<b>退出系统</b>」，或双击安装目录里的「<b>停止系统.bat</b>」。"
        ),
        do_not=(
            "• <b>不要</b>把软件文件夹里的单独程序拖到桌面（请在「启动系统(无黑框后台版).vbs」上右键→发送到→桌面快捷方式）。<br/>"
            "• <b>不要</b>在正在生成 Word 或批量打印时直接拔电源或用任务管理器强行结束进程。<br/>"
            "• <b>不要</b>在首次配置向导时跳过工作目录设置，首次选好「工作目录」即可长期自动记住。"
        ),
    ))

    story.append(Spacer(1, 4 * mm))
    # 编号前缀速查小表
    prefix_tbl = Table(
        [
            [Paragraph("<b>病理编号前缀速查（登记时按报告分类点选）</b>", ParagraphStyle("PH", fontName="YaHeiBold", fontSize=9.5, textColor=colors.white)), "", "", "", ""],
            ["无前缀", "B 前缀", "L 前缀", "H 前缀", "T 前缀"],
            ["普通门诊报告", "病房报告", "分院区门诊报告", "会诊报告", "其他特殊报告"],
        ],
        colWidths=[34.8 * mm] * 5,
    )
    prefix_tbl.setStyle(TableStyle([
        ("SPAN", (0, 0), (4, 0)),
        ("BACKGROUND", (0, 0), (4, 0), BLUE),
        ("BACKGROUND", (0, 1), (4, 1), LIGHT_BLUE),
        ("FONTNAME", (0, 1), (-1, -1), "YaHei"),
        ("FONTNAME", (0, 1), (4, 1), "YaHeiBold"),
        ("FONTSIZE", (0, 1), (-1, -1), 8.5),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOX", (0, 0), (-1, -1), 0.7, BLUE),
        ("INNERGRID", (0, 0), (-1, -1), 0.4, BORDER_GRAY),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(prefix_tbl)

    # -------- 第 2 页：SOP-02 患者录入 + 实景截图 1 --------
    story.append(PageBreak())
    story.append(_sop_card(
        st,
        "SOP-02",
        "登记新患者与创建病理照片文件夹（示例：张建国）",
        need_do="为每位新患者生成不重复的当日病理号、Word 报告初稿，以及存放显微照片的专属文件夹。",
        how_to=(
            "<b>① 编号：</b>选好来源前缀（无 / B / L / H / T），点击【<b>自动编号</b>】。<br/>"
            "<b>② 填表：</b>输入患者姓名（如：张建国）、性别、年龄（新生儿可填 3月 / 15天），点击图片数 <b>1~9</b> 并补全每张图对应的送检部位。<br/>"
            "<b>③ 建档：</b>点绿色的【<b>保存并生成病理报告</b>】，右侧“今日列表”会立刻出现该患者，同时硬盘里自动建好照片文件夹。<br/>"
            "<b>④ 对照：</b>选中右侧列表中的患者后点【<b>临床照片查询</b>】，可一键查看关联的皮肤科临床皮损照片。"
        ),
        do_not=(
            "• <b>不要</b>自己硬编和别人一样的病理号——系统会提示重号并拦截。<br/>"
            "• <b>不要</b>在外部用 Word 手动改病理号或改“临床照片”文件夹名；如需改病理号请回网页「修改患者信息」改，改名后若遇弹窗按需确认立即重归档即可。<br/>"
            "• <b>不要</b>选了 4 张图却只填 2 个送检部位，部位框要和图片张数保持一致。"
        ),
    ))
    _add_screenshot(story, st, screenshots_dir, "shot1_entry.png", "附图 1：患者录入界面按钮引导（①自动编号 → ②填表选图数 → ③保存建档 → ④临床对照）", max_h_mm=128)

    # -------- 第 3 页：SOP-03 报告制作 + 实景截图 2 --------
    story.append(PageBreak())
    story.append(_sop_card(
        st,
        "SOP-03",
        "放显微照片、选用词条模板与「今日批量归档」",
        need_do="把拍好的显微照片放进患者目录，选好词条模板与流转状态生成 Word 报告，确认后单人或批量归档。",
        how_to=(
            "<b>① 批量收尾：</b>切到「报告制作」页签；若当天常规/纯组化报告都已保存检查好，点右侧【<b>批量归档今日常规/纯组化</b>】一键收尾。<br/>"
            "<b>② 放图核对：</b>平时制作单份病例时，先点蓝色【<b>打开/复制路径</b>】把显微图片放入文件夹（点【<b>临床照片查询</b>】可对照皮损）。<br/>"
            "<b>③ 选图写诊断：</b>选好照片排版、套用常用词条模板（如有需要勾选“待审核 / 需重切 / 仅免疫组化”），点【<b>保存报告并输出</b>】生成 Word 并刷新预览。<br/>"
            "<b>④ 归档/打一页：</b>常规或已审核通过病例点【<b>一键归档</b>】，需要打核对单可点右侧【<b>打印第一页</b>】。"
        ),
        do_not=(
            "• <b>不要</b>开着电脑版 Word 占用同一份报告，同时又在网页里点“保存报告并输出”——先保存并关掉 Word 再回网页。<br/>"
            "• <b>不要</b>在外部用 Word 改完字后无视弹窗盲目覆盖；列表出现「Word已外改」时点【<b>确定</b>】即可一键导入外部修改并保留字体颜色和批注。<br/>"
            "• <b>不要</b>把还勾着“待审核 / 需重切”的疑难病例强行直接归档，先保存分流进审核队列。"
        ),
    ))
    _add_screenshot(story, st, screenshots_dir, "shot2_making.png", "附图 2：报告制作工作区（①批量归档今日 → ②放照片核对 → ③选图词条保存输出 → ④一键归档/打印）", max_h_mm=124)

    # -------- 第 4 页：SOP-04 报告审核 + 实景截图 3 --------
    story.append(PageBreak())
    story.append(_sop_card(
        st,
        "SOP-04",
        "报告审核队列、重切补图与双草稿同步",
        need_do="集中处理待审核或需重切的特殊病例，放好重切照片、修改确认后再回常规流完成归档。",
        how_to=(
            "<b>① 放重切图：</b>进入「报告审核」页签，选中左侧待审患者，把重切照片放入特殊归档照片夹，点【<b>重新读取照片</b>】。<br/>"
            "<b>② 改稿同步：</b>修改医生/日期/诊断后点【<b>保存同步草稿</b>】（此时不会误改已归档正式库，系统自动维护双草稿一致）。<br/>"
            "<b>③ 边存边打：</b>需要打纸质第 1 页核对单请医生/主任（如李主任）签核时，点【<b>一键打印第一页</b>】。<br/>"
            "<b>④ 审核放行：</b>核对无误后点绿色【<b>完成审核</b>】，再切回「报告制作」点击一键归档即可进入报告输出。"
        ),
        do_not=(
            "• <b>不要</b>在电脑文件夹窗口里手动 Delete 删掉待审核患者目录，任何退回/作废均走网页按钮。<br/>"
            "• <b>不要</b>忘记点【完成审核】就去报告输出里找，必须经过“完成审核”才可最终归档。<br/>"
            "• <b>不要</b>在外部 Word 未保存关闭时点击网页归档移动文件，外改润色后先关闭 Word 再回网页保存/归档。"
        ),
    ))
    _add_screenshot(story, st, screenshots_dir, "shot3_review.png", "附图 3：报告审核工作区（①重切取图 → ②保存同步草稿 → ③打印第一页 → ④完成审核）", max_h_mm=124)

    # -------- 第 5 页：SOP-05 独立免疫组化 + 实景截图 4 --------
    story.append(PageBreak())
    story.append(_sop_card(
        st,
        "SOP-05",
        "同一患者加做「免疫组化」独立第二份报告",
        need_do="在保留已归档常规病理报告原件不变的前提下，为加做免疫组化的患者独立生成第二份「*_免疫组化报告」。",
        how_to=(
            "<b>① 建组化副本：</b>切到「免疫组化」页签，搜索选中原报告已归档的患者，点【<b>创建/重建组化报告</b>】复制出独立工作副本。<br/>"
            "<b>② 放染色照片：</b>点【<b>打开/复制照片路径</b>】放入免疫组化显微图，可同时点【<b>临床照片查询</b>】对照。<br/>"
            "<b>③ 独立存与归档：</b>录入染色结果后先点【<b>保存</b>】，预览无误再点【<b>归档</b>】，输出独立「*_免疫组化报告」文件。<br/>"
            "<b>④ 误标作废：</b>如不再加做或建错，点红色【<b>删除免疫组化报告</b>】安全清理，绝不损伤原常规报告。"
        ),
        do_not=(
            "• <b>不要</b>直接去修改已归档的原常规病理 Word 冒充免疫组化报告，必须走独立「免疫组化」页签。<br/>"
            "• <b>不要</b>手改「*_免疫组化报告」的文件名后缀，避免「报告输出」右栏识别不到。<br/>"
            "• <b>不要</b>混淆删除按钮：此处的删除仅清组化第二份报告，不会删除常规原报告。"
        ),
    ))
    _add_screenshot(story, st, screenshots_dir, "shot4_ihc.png", "附图 4：独立免疫组化工作区（①建工作副本 → ②放染色照片 → ③保存并归档 → ④安全删除不伤原件）", max_h_mm=124)

    # -------- 第 6 页：SOP-06 报告输出、签名单与综合查询 + 实景截图 5 --------
    story.append(PageBreak())
    story.append(_sop_card(
        st,
        "SOP-06",
        "报告三栏批量打印、领取签名单与病历综合查询",
        need_do="分类批量打印已归档报告、导出 Word/Excel 双版本领取签名单，以及随时反查历史报告并补改。",
        how_to=(
            "<b>①②③ 三栏分批打印：</b>进入「报告输出」页签，左栏【常规已归档】、中栏【特殊审核完成】、右栏【免疫组化已归档】相互独立，支持各自一键【<b>批量打印</b>】或【<b>撤销归档/审核</b>】。<br/>"
            "<b>④ 精准日期过滤：</b>三栏顶部均支持切换“报告日期 / 送检日期”、时间范围及多前缀复选过滤。<br/>"
            "<b>⑤ 导签名单：</b>切到顶部「<b>报告签名单</b>」页签，一键生成 Word 版和 Excel 版患者报告领取签名单。<br/>"
            "<b>⑥ 查历史补改：</b>切到「<b>病历综合查询</b>」页签，可按姓名、病理号、医生或<b>报告日期</b>反查全库，未归档老记录也可直接打开继续编辑。"
        ),
        do_not=(
            "• <b>不要</b>连续狂点【批量打印】，点一次后等待默认打印机依次出纸即可。<br/>"
            "• <b>不要</b>打印前忘记检查电脑默认打印机是否联机、纸盒是否有纸。<br/>"
            "• <b>不要</b>手拖 Word 底部医生签名栏前面的空格，系统已用精确定位保证单行右对齐、不折行。"
        ),
    ))
    _add_screenshot(story, st, screenshots_dir, "shot5_output.png", "附图 5：报告输出三列表（左栏常规 / 中栏特殊审核 / 右栏独立免疫组化，支持日期及前缀多选）", max_h_mm=104)

    # -------- 第 7 页：1-9 图排版图 + 常见问题 10 秒自查表 --------
    story.append(PageBreak())
    story.append(Paragraph("附表 1：图文自动排版规则（1~9 张图均无需手工调格）", st["h1"]))
    story.append(Image(diag_layout, width=174 * mm, height=37 * mm))
    story.append(Spacer(1, 5 * mm))

    story.append(Paragraph("附表 2：常见操作问题 10 秒自查表", st["h1"]))
    faq_rows = [
        [
            Paragraph("<b>遇到什么情况</b>", ParagraphStyle("FH", fontName="YaHeiBold", fontSize=9.5, textColor=colors.white)),
            Paragraph("<b>原因说明（极简版）</b>", ParagraphStyle("FH", fontName="YaHeiBold", fontSize=9.5, textColor=colors.white)),
            Paragraph("<b>立刻怎么做</b>", ParagraphStyle("FH", fontName="YaHeiBold", fontSize=9.5, textColor=colors.white)),
        ],
        [
            "关掉浏览器后想再进系统，双击图标没反应或怕多开",
            "后台服务本来就还开着，不会多开新窗口",
            "直接再双击一次桌面「启动系统(无黑框后台版).vbs」，网页会自动弹出来",
        ],
        [
            "患者列表显示黄标「Word已外改」或保存被拦截",
            "有人直接用电脑 Word 改过并保存了这份文件",
            "点患者后在弹窗点【确定】即可把 Word 里的改字/红字/批注同步进系统（不愿保留点后续确认按系统重盖）",
        ],
        [
            "点生成或归档提示“文件被占用 / 无法保存”",
            "您用电脑版 Word 打开了这份报告还没关",
            "把任务栏里对应的 Word 窗口保存并关闭，回网页再点一次按钮即可",
        ],
        [
            "删除或覆盖时弹出“回收站移入失败—防永久删除拦截”",
            "工作目录处于网络共享盘或权限受限，无法进回收站",
            "不建议直接永久删除，请核对无误后再确认；日常旧版本报告会自动先存一份到系统回收站",
        ],
        [
            "想给某位患者补改几天前已经归档的报告",
            "报告已进入历史归档库",
            "去「病历综合查询」搜姓名/病理号/日期找到他，点编辑或撤销归档后再修改",
        ],
        [
            "打印出来的报告最后签名跑到了第 2 页",
            "文字行数偏多或手改了多余空行",
            "删掉报告里自己多敲的空行后在网页点保存重生成（勿手拉空格）",
        ],
        [
            "点打印没反应或报错（尤其旧系统/老 Word）",
            "默认打印机未设好，或正在弹窗确认",
            "进控制面板确认默认打印机在线有纸；使用软件前确保 Word 未卡在弹窗状态",
        ],
        [
            "软件异常想彻底重启刷新",
            "后台服务需要冷重启一次",
            "双击软件目录内的「停止系统.bat」，等 2 秒后再双击桌面「.vbs」启动",
        ],
    ]
    p_rows = []
    for i, r in enumerate(faq_rows):
        if i == 0:
            p_rows.append(r)
        else:
            p_rows.append([Paragraph(c, st["sop_cell"]) for c in r])

    faq_tbl = Table(p_rows, colWidths=[50 * mm, 52 * mm, 72 * mm])
    faq_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, ROW_ALT]),
        ("BOX", (0, 0), (-1, -1), 0.8, NAVY),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, BORDER_GRAY),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(faq_tbl)

    story.append(Spacer(1, 6 * mm))
    tip_box = Table(
        [[Paragraph(
            "<b>日常口诀（贴在工位旁）：</b><br/>"
            "上班双击 .vbs 一图标直接进 → 编号建档放照片 → 选词条预览没问题就归档 → "
            "需重切进「报告审核」→ 加做组化走独立页签 → 批量打印配签名单 → 下班点退出或「停止系统.bat」。",
            ParagraphStyle("Tip", fontName="YaHei", fontSize=9.5, leading=14.5, textColor=NAVY, wordWrap="CJK"),
        )]],
        colWidths=[174 * mm],
    )
    tip_box.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), LIGHT_AMBER),
        ("BOX", (0, 0), (-1, -1), 0.9, AMBER),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
    ]))
    story.append(tip_box)

    doc.build(story, onFirstPage=_page_footer, onLaterPages=_page_footer)
    # 清理临时示意图
    for f in (diag_workflow, diag_layout):
        if os.path.exists(f):
            try:
                os.remove(f)
            except OSError:
                pass
    try:
        os.rmdir(img_dir)
    except OSError:
        pass


def main():
    parser = argparse.ArgumentParser(description="生成皮肤病理报告系统快速入门操作手册 PDF")
    parser.add_argument(
        "--output",
        default=os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "文档",
            "皮肤病理报告系统_快速入门操作手册.pdf",
        ),
        help="输出 PDF 路径",
    )
    parser.add_argument(
        "--screenshots-dir",
        default=None,
        help="可选：脱敏标注截图所在目录（包含 shot1_entry.png ~ shot5_output.png）",
    )
    args = parser.parse_args()
    build_manual_pdf(args.output, screenshots_dir=args.screenshots_dir)
    print(f"[OK] Generated SOP Manual PDF: {args.output}")


if __name__ == "__main__":
    main()
