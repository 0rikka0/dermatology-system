# -*- coding: utf-8 -*-
"""
生成皮肤病理报告系统快速入门与标准操作规程 (SOP) 手册 (PDF 格式, 含高清示意图)
严格执行脱敏处理，排版清晰易读，覆盖全流程 SOP。
"""

import os
import sys
from PIL import Image as PILImage, ImageDraw, ImageFont
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, KeepTogether, HRFlowable, PageBreak
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

FONT_REGULAR = "C:/Windows/Fonts/msyh.ttc"
FONT_BOLD = "C:/Windows/Fonts/msyhbd.ttc"

pdfmetrics.registerFont(TTFont("YaHei", FONT_REGULAR))
pdfmetrics.registerFont(TTFont("YaHeiBold", FONT_BOLD))


def get_pil_font(size, bold=False):
    return ImageFont.truetype(FONT_BOLD if bold else FONT_REGULAR, size)


def generate_diagram_workflow(out_path):
    """生成系统全景标准工作流示意图"""
    w, h = 1600, 390
    im = PILImage.new("RGB", (w, h), (248, 250, 252))
    draw = ImageDraw.Draw(im)

    f_title = get_pil_font(24, bold=True)
    f_step = get_pil_font(20, bold=True)
    f_desc = get_pil_font(17, bold=False)

    steps = [
        ("01", "患者信息录入", "填写基本资料与病史\n选1-9图与送检部位\n自动建初稿及照片夹", (59, 130, 246)),
        ("02", "报告制作排版", "读取病理照片与词库\n常规及纯组化直接归档\n待审/重切双草稿分流", (16, 185, 129)),
        ("03", "待审/重切审核", "待审核与重切完全等价\n病理照读复制后目录\n保存双同步·完成待归档", (245, 158, 11)),
        ("04", "独立免疫组化", "基于已归档原报告创建\n独立编辑保存不动原件\n加后缀双份报告并存", (139, 92, 246)),
        ("05", "三列表输出打印", "常规/特殊/组化物理隔离\nWord COM 仅打第1页\n生成Word+Excel签名单", (239, 68, 68)),
    ]

    card_w = 274
    card_h = 290
    gap = 36
    start_x = 43
    start_y = 48

    for i, (num, title, desc, color) in enumerate(steps):
        x = start_x + i * (card_w + gap)
        y = start_y

        draw.rounded_rectangle([x + 3, y + 3, x + card_w + 3, y + card_h + 3], radius=16, fill=(226, 232, 240))
        draw.rounded_rectangle([x, y, x + card_w, y + card_h], radius=16, fill=(255, 255, 255), outline=color, width=3)

        draw.rounded_rectangle([x, y, x + card_w, y + 66], radius=14, fill=color)
        draw.rectangle([x, y + 48, x + card_w, y + 66], fill=color)

        draw.text((x + 16, y + 16), f"SOP-{num}", font=f_step, fill=(255, 255, 255))
        draw.text((x + 18, y + 86), title, font=f_title, fill=(30, 41, 59))

        dy = y + 136
        for line in desc.split("\n"):
            draw.text((x + 18, dy), "• " + line, font=f_desc, fill=(100, 116, 139))
            dy += 36

        if i < len(steps) - 1:
            ax = x + card_w + 6
            ay = y + card_h // 2
            draw.line([ax, ay, ax + 20, ay], fill=(148, 163, 184), width=4)
            draw.polygon([(ax + 20, ay - 8), (ax + 30, ay), (ax + 20, ay + 8)], fill=(148, 163, 184))

    im.save(out_path, quality=95)


def generate_diagram_special_flow(out_path):
    """生成“待审核/重切双草稿流”与“独立免疫组化双报告流”对照示意图"""
    w, h = 1540, 520
    im = PILImage.new("RGB", (w, h), (248, 250, 252))
    draw = ImageDraw.Draw(im)

    f_title = get_pil_font(22, bold=True)
    f_badge = get_pil_font(17, bold=True)
    f_sub = get_pil_font(16, bold=False)

    # Outer frame
    draw.rounded_rectangle([20, 16, w - 20, h - 16], radius=16, fill=(255, 255, 255), outline=(203, 213, 225), width=2)

    # ---- Track 1: 待审核 / 重切 ----
    draw.rounded_rectangle([40, 36, 1500, 252], radius=12, fill=(255, 251, 235), outline=(245, 158, 11), width=2)
    draw.text((60, 48), "轨道 1：待审核 / 重切（处理逻辑完全等价 · 同一份报告双草稿同步）", font=f_title, fill=(180, 83, 9))

    t1_boxes = [
        (60, "1. 制作保存分流", "输出草稿至 病理报告输出\\\n仅复制病理照片+草稿至\n特殊归档/待审核(或重切)", (254, 243, 199), (217, 119, 6), (180, 83, 9)),
        (415, "2. 审核区照片与编辑", "临床照查 未归档\\临床照片\n病理照打开/复制 特殊归档目录\n保存双向同步两份草稿+新增图", (255, 255, 255), (217, 119, 6), (30, 41, 59)),
        (770, "3. 完成审核标记", "审核保存不动 已归档报告\n点 [完成审核] 改状态为 完成审核\n(已归档被改则置 archived=2)", (236, 253, 245), (16, 185, 129), (4, 120, 87)),
        (1125, "4. 回制作区归档", "回 [报告制作] 点 [一键归档]\n写入 已归档报告/<原名>.docx\n待审核/重切目录移入 完成审核", (240, 253, 250), (13, 148, 136), (15, 118, 110)),
    ]
    bw, bh = 315, 140
    for idx, (bx, btitle, btxt, bg, bd, fg) in enumerate(t1_boxes):
        by = 92
        draw.rounded_rectangle([bx, by, bx + bw, by + bh], radius=10, fill=bg, outline=bd, width=2)
        draw.text((bx + 14, by + 12), btitle, font=f_badge, fill=fg)
        dy = by + 44
        for ln in btxt.split("\n"):
            draw.text((bx + 14, dy), ln, font=f_sub, fill=(55, 65, 81))
            dy += 28
        if idx < len(t1_boxes) - 1:
            ax = bx + bw + 6
            ay = by + bh // 2
            draw.line([ax, ay, ax + 24, ay], fill=(217, 119, 6), width=3)
            draw.polygon([(ax + 24, ay - 7), (ax + 34, ay), (ax + 24, ay + 7)], fill=(217, 119, 6))

    # ---- Track 2: 独立免疫组化 ----
    draw.rounded_rectangle([40, 276, 1500, 492], radius=12, fill=(245, 243, 255), outline=(124, 58, 237), width=2)
    draw.text((60, 288), "轨道 2：需免疫组化（独立双报告机制 · 原报告常规出片后，再独立制作第二份免疫组化报告）", font=f_title, fill=(109, 40, 217))

    t2_boxes = [
        (60, "1. 原报告按常规走", "仅标 需免疫组化 不视作待审\n可在制作区直接归档和打印原报告\n(若兼待审核则按轨道1先审归)", (238, 242, 255), (99, 102, 241), (67, 56, 202)),
        (415, "2. 创建免疫组化副本", "进 [免疫组化] 点 [创建组化报告]\n须原报告已在 已归档报告\\\n复制原报告+病理照片(无临床照)", (255, 255, 255), (124, 58, 237), (30, 41, 59)),
        (770, "3. 独立编辑与保存", "打开/复制组化照片夹加特染图\n点 [保存] 仅改 特殊归档/免疫组化\n绝不动原报告与 patients 主表", (255, 255, 255), (124, 58, 237), (30, 41, 59)),
        (1125, "4. 独立归档并存输出", "已归档报告 生成 _免疫组化报告\n与原报告并存 · 不进 完成审核\n在 [报告输出] 列表3独立打印", (243, 232, 255), (147, 51, 234), (107, 33, 168)),
    ]
    for idx, (bx, btitle, btxt, bg, bd, fg) in enumerate(t2_boxes):
        by = 332
        draw.rounded_rectangle([bx, by, bx + bw, by + bh], radius=10, fill=bg, outline=bd, width=2)
        draw.text((bx + 14, by + 12), btitle, font=f_badge, fill=fg)
        dy = by + 44
        for ln in btxt.split("\n"):
            draw.text((bx + 14, dy), ln, font=f_sub, fill=(55, 65, 81))
            dy += 28
        if idx < len(t2_boxes) - 1:
            ax = bx + bw + 6
            ay = by + bh // 2
            draw.line([ax, ay, ax + 24, ay], fill=(124, 58, 237), width=3)
            draw.polygon([(ax + 24, ay - 7), (ax + 34, ay), (ax + 24, ay + 7)], fill=(124, 58, 237))

    im.save(out_path, quality=95)


def generate_diagram_page_layout(out_path):
    """生成版面控制与首页单页打印示意图"""
    w, h = 1400, 380
    im = PILImage.new("RGB", (w, h), (248, 250, 252))
    draw = ImageDraw.Draw(im)

    f_title = get_pil_font(22, bold=True)
    f_sub = get_pil_font(17, bold=False)
    f_badge = get_pil_font(17, bold=True)

    p1_x, p1_y, p1_w, p1_h = 90, 32, 460, 316
    draw.rounded_rectangle([p1_x + 4, p1_y + 4, p1_x + p1_w + 4, p1_y + p1_h + 4], radius=12, fill=(226, 232, 240))
    draw.rounded_rectangle([p1_x, p1_y, p1_x + p1_w, p1_y + p1_h], radius=12, fill=(255, 255, 255), outline=(59, 130, 246), width=2)
    draw.rectangle([p1_x, p1_y, p1_x + p1_w, p1_y + 36], fill=(59, 130, 246))
    draw.text((p1_x + 15, p1_y + 8), "【第 1 页】法定病理报告单 (Word COM 仅打此页)", font=f_badge, fill=(255, 255, 255))

    draw.rectangle([p1_x + 20, p1_y + 48, p1_x + p1_w - 20, p1_y + 84], fill=(241, 245, 249), outline=(203, 213, 225))
    draw.text((p1_x + 30, p1_y + 56), "表头: 病理号 / 姓名 / 性别 / 年龄 / 送检部位", font=f_sub, fill=(71, 85, 105))

    draw.rectangle([p1_x + 20, p1_y + 94, p1_x + p1_w - 20, p1_y + 164], fill=(238, 242, 255), outline=(199, 210, 254))
    draw.text((p1_x + 30, p1_y + 104), "显微图像排版区 (1-9图自动行数补偿)", font=f_badge, fill=(67, 56, 202))
    draw.text((p1_x + 30, p1_y + 134), "图片下方精确对齐部位标注 (1.部位名)", font=f_sub, fill=(100, 116, 139))

    draw.rectangle([p1_x + 20, p1_y + 174, p1_x + p1_w - 20, p1_y + 244], fill=(240, 253, 244), outline=(187, 247, 208))
    draw.text((p1_x + 30, p1_y + 184), "镜下描述与病理诊断 (锁定 15 磅固定行距)", font=f_badge, fill=(22, 101, 52))
    draw.text((p1_x + 30, p1_y + 214), "自动补偿空行，严防签名溢出到第二页", font=f_sub, fill=(100, 116, 139))

    draw.rectangle([p1_x + 20, p1_y + 254, p1_x + p1_w - 20, p1_y + 296], fill=(254, 243, 199), outline=(252, 211, 77))
    draw.text((p1_x + 30, p1_y + 264), "署名栏: 报告医生 / 审核医生 / 报告日期", font=f_badge, fill=(180, 83, 9))

    draw.line([595, 190, 725, 190], fill=(37, 99, 235), width=6)
    draw.polygon([(725, 180), (745, 190), (725, 200)], fill=(37, 99, 235))
    draw.rounded_rectangle([585, 122, 735, 166], radius=8, fill=(37, 99, 235))
    draw.text((602, 134), "Word COM 打印", font=f_badge, fill=(255, 255, 255))
    draw.text((600, 210), "只打第 1 页", font=f_title, fill=(220, 38, 38))
    draw.text((585, 244), "自动跳过第 2 页附页", font=f_sub, fill=(100, 116, 139))

    p2_x, p2_y, p2_w, p2_h = 790, 32, 460, 316
    draw.rounded_rectangle([p2_x + 4, p2_y + 4, p2_x + p2_w + 4, p2_y + p2_h + 4], radius=12, fill=(226, 232, 240))
    draw.rounded_rectangle([p2_x, p2_y, p2_x + p2_w, p2_y + p2_h], radius=12, fill=(255, 255, 255), outline=(148, 163, 184), width=2)
    draw.rectangle([p2_x, p2_y, p2_x + p2_w, p2_y + 36], fill=(148, 163, 184))
    draw.text((p2_x + 15, p2_y + 8), "【第 2 页】临床病史附页 (电子留档专用)", font=f_badge, fill=(255, 255, 255))

    draw.rectangle([p2_x + 20, p2_y + 58, p2_x + p2_w - 20, p2_y + 200], fill=(248, 250, 252), outline=(226, 232, 240))
    draw.text((p2_x + 30, p2_y + 74), "详细临床病史与送检特征记录:", font=f_badge, fill=(71, 85, 105))
    draw.text((p2_x + 30, p2_y + 112), "• 完整留存 .docx 归档文件内可追溯", font=f_sub, fill=(100, 116, 139))
    draw.text((p2_x + 30, p2_y + 148), "• 纸质输出自动跳过，节约临床耗材", font=f_sub, fill=(100, 116, 139))

    draw.rounded_rectangle([p2_x + 75, p2_y + 234, p2_x + 385, p2_y + 286], radius=10, fill=(254, 226, 226), outline=(239, 68, 68), width=2)
    draw.text((p2_x + 102, p2_y + 247), "免打印附页 · 防止多页废纸", font=f_title, fill=(220, 38, 38))

    im.save(out_path, quality=95)


class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        print(f"[INFO] Total generated PDF pages: {num_pages}")
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_page_decorations(self, page_count):
        self.saveState()
        self.setFont("YaHei", 8.5)
        self.setFillColor(colors.HexColor("#64748b"))

        if self._pageNumber > 1:
            self.drawString(50, 804, "皮肤病理报告系统 v1.4 — 快速入门与标准操作规程 (SOP) 手册 (脱敏图文版)")
            self.setStrokeColor(colors.HexColor("#e2e8f0"))
            self.setLineWidth(0.5)
            self.line(50, 796, 545, 796)

        page_str = f"第 {self._pageNumber} 页 / 共 {page_count} 页"
        self.drawRightString(545, 34, page_str)
        self.drawString(50, 34, "标准操作 SOP · 质控安全隔离 · 离线便携运行")
        self.setStrokeColor(colors.HexColor("#e2e8f0"))
        self.setLineWidth(0.5)
        self.line(50, 46, 545, 46)
        self.restoreState()


def build_manual_pdf(pdf_path):
    work_dir = os.path.dirname(pdf_path)
    img_dir = os.path.join(work_dir, ".temp_imgs")
    os.makedirs(img_dir, exist_ok=True)

    diag1 = os.path.join(img_dir, "diag1_workflow.png")
    diag2 = os.path.join(img_dir, "diag2_special.png")
    diag3 = os.path.join(img_dir, "diag3_layout.png")

    generate_diagram_workflow(diag1)
    generate_diagram_special_flow(diag2)
    generate_diagram_page_layout(diag3)

    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        leftMargin=50,
        rightMargin=50,
        topMargin=52,
        bottomMargin=52
    )

    s_title = ParagraphStyle("DocTitle", fontName="YaHeiBold", fontSize=22, leading=28, textColor=colors.HexColor("#0f172a"), alignment=1, spaceAfter=6)
    s_subtitle = ParagraphStyle("DocSubtitle", fontName="YaHei", fontSize=11.5, leading=17, textColor=colors.HexColor("#475569"), alignment=1, spaceAfter=10)
    s_badge = ParagraphStyle("Badge", fontName="YaHei", fontSize=8.8, leading=13, textColor=colors.HexColor("#0284c7"), alignment=1, spaceAfter=14)
    s_h1 = ParagraphStyle("H1", fontName="YaHeiBold", fontSize=13.5, leading=20, textColor=colors.HexColor("#1e3a8a"), spaceBefore=12, spaceAfter=6, keepWithNext=True)
    s_h2 = ParagraphStyle("H2", fontName="YaHeiBold", fontSize=11, leading=16, textColor=colors.HexColor("#0369a1"), spaceBefore=8, spaceAfter=4, keepWithNext=True)
    s_body = ParagraphStyle("Body", fontName="YaHei", fontSize=9.2, leading=14.5, textColor=colors.HexColor("#334155"), spaceAfter=4)
    s_bold = ParagraphStyle("BodyBold", fontName="YaHeiBold", fontSize=9.2, leading=14.5, textColor=colors.HexColor("#1e293b"), spaceBefore=5, spaceAfter=2, keepWithNext=True)
    s_tip = ParagraphStyle("Tip", fontName="YaHei", fontSize=8.8, leading=13.8, textColor=colors.HexColor("#0f766e"))
    s_warn = ParagraphStyle("Warn", fontName="YaHei", fontSize=8.8, leading=14, textColor=colors.HexColor("#991b1b"))
    s_tbl = ParagraphStyle("Tbl", fontName="YaHei", fontSize=8.3, leading=12, textColor=colors.HexColor("#334155"))
    s_tbl_b = ParagraphStyle("TblB", fontName="YaHeiBold", fontSize=8.4, leading=12.5, textColor=colors.HexColor("#0f172a"))
    s_tbl_hdr = ParagraphStyle("TblHdr", fontName="YaHeiBold", fontSize=8.6, leading=13, textColor=colors.HexColor("#ffffff"))

    story = []
    W = 495

    def make_callout(html_text, bg="#f0fdfa", border="#99f6e4", style=s_tip, border_w=1):
        t = Table([[Paragraph(html_text, style)]], colWidths=[W])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(bg)),
            ('BOX', (0, 0), (-1, -1), border_w, colors.HexColor(border)),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
            ('RIGHTPADDING', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 7),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 7),
        ]))
        return t

    def add_sop_block(sop_code, title, scope, steps_list, rule_note):
        hdr = [[Paragraph(f"{sop_code}：{title}", s_tbl_hdr), Paragraph(f"适用范围：{scope}", s_tbl_hdr)]]
        rows = list(hdr)
        for idx, (step_name, detail) in enumerate(steps_list, 1):
            rows.append([
                Paragraph(f"<b>{idx}. {step_name}</b>", s_tbl_b),
                Paragraph(detail, s_tbl)
            ])
        if rule_note:
            rows.append([
                Paragraph("<b>关键红线/核查点</b>", ParagraphStyle("RedL", fontName="YaHeiBold", fontSize=8.3, leading=12, textColor=colors.HexColor("#b91c1c"))),
                Paragraph(rule_note, ParagraphStyle("RedR", fontName="YaHei", fontSize=8.3, leading=12.2, textColor=colors.HexColor("#991b1b")))
            ])

        tbl = Table(rows, colWidths=[125, W - 125])
        ts = [
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1e3a8a")),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor("#f8fafc"), colors.HexColor("#ffffff")]),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 7),
            ('RIGHTPADDING', (0, 0), (-1, -1), 7),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ]
        if rule_note:
            ts.append(('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#fff1f2")))
        tbl.setStyle(TableStyle(ts))
        story.append(KeepTogether([tbl, Spacer(1, 8)]))

    # Cover
    story.append(Paragraph("皮肤病理报告系统", s_title))
    story.append(Paragraph("快速入门与标准操作规程 (SOP) 手册 (v1.4 脱敏图文版)", s_subtitle))
    story.append(Paragraph("【适用：报告医师 / 审核专家 / 病理技师 / 规培进修生】  |  【八模块全流程 · 独立免疫组化 · Word 单页打印】", s_badge))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#3b82f6"), spaceAfter=10))

    story.append(make_callout(
        "<b>★ 核心作业分工总览（必读）：</b> 本系统严格区分三类业务轨道——① <b>常规病例及仅标注“需免疫组化”的病例</b>：原初诊报告按普通流程在【报告制作】完成保存、打印与一键归档；② <b>标注“待审核”或“需重切”的病例</b>（二者完全等价）：仅复制病理照片到特殊待审目录，在【报告审核】协同修改，保存时双向同步两份草稿，审核完成后回【报告制作】归档；③ <b>免疫组化第二份报告</b>：在【免疫组化】模块基于已归档原报告创建独立副本，独立保存并归档为 <code>_免疫组化报告.docx</code>，与原报告并行不覆盖。"
    ))
    story.append(Spacer(1, 10))

    # Overview Diagram
    story.append(Paragraph("一、系统全流程标准作业蓝图 (SOP 总览)", s_h1))
    story.append(Image(diag1, width=W, height=120))
    story.append(Spacer(1, 8))
    story.append(Image(diag2, width=W, height=167))
    story.append(PageBreak())

    # SOPs
    story.append(Paragraph("二、各岗位标准操作规程 (SOP)", s_h1))

    add_sop_block(
        "SOP-01",
        "系统启停、初始化与环境适配",
        "每日开闭系统及首次部署（Win10/Win11 一体包 & Win7 专版）",
        [
            ("启动服务", "双击根目录 <code>启动系统.bat</code>（或无黑框后台版 <code>.vbs</code>），等待浏览器自动打开 <code>http://127.0.0.1:8765</code>。Windows 7 用户请使用 Chrome 109 / Edge 109 / 360极速模式（禁止使用 IE 内核）。"),
            ("确认免责声明", "每次进入 WebUI 需阅读执业复核免责提示，点击【我已知晓并同意】方可作业；点击【退出】将安全关闭后台服务。"),
            ("首次目录初始化", "在向导中设定大容量磁盘路径作为“工作目录”（系统自动建 <code>已归档报告</code>、<code>未归档报告</code>、<code>特殊患者归档</code>、<code>领取表</code>、<code>数据库</code> 5 大主目录）及默认医生姓名（如演示值：张医生 / 李主任）。"),
            ("安全关停", "工作结束后双击 <code>停止系统.bat</code>，释放本地 8765 端口并确保数据库事务提交。")
        ],
        "禁止直接强删运行中的 database/pathology.db；文档统一位于软件根目录 文档/ 下，请勿移动模板/与文字模板/目录。"
    )

    add_sop_block(
        "SOP-02",
        "患者信息登记建档与基本信息修改",
        "【患者信息录入】工作区",
        [
            ("选择前缀与取号", "按标本来源选择前缀：无前缀（门诊）、<code>B</code>（病房）、<code>L</code>（分院区）、<code>H</code>（会诊）、<code>T</code>（特殊）。点击【自动编号】生成当日独立递增号（各前缀互不冲号）。"),
            ("填写信息与部位", "录入姓名、性别、年龄（支持下拉 <code>岁/月/日</code> 或 <code>自由输入</code> 如 <code>1岁3月</code>）。<code>H/T</code> 前缀支持修改送检单位与科室（其余固定为 本院/皮肤科）。选定图片数（1-9）并在下方依次填写送检部位。"),
            ("生成初稿与照片夹", "点击【保存并生成病理报告】，系统写入 <code>未归档报告\\病理报告录入\\病理号_姓名.docx</code> 并创建 <code>未归档报告\\病理照片\\病理号_姓名\\</code>。选中列表行可点【临床照片查询】打开 <code>未归档报告\\临床照片\\</code>。"),
            ("已录信息修改与重归档", "修改已保存/已归档患者信息保存时，系统同步更新录入稿、照片夹、输出稿及特殊归档（待审核/重切/完成审核/免疫组化），并弹窗询问是否立即重新归档。")
        ],
        "修改患者基本信息时严禁系统直接覆写 已归档报告\\ 中的正式文件；点“否”延迟重归档则状态自动标记为 需要重新归档(2)。"
    )

    add_sop_block(
        "SOP-03",
        "常规病理报告与原初诊报告制作、归档及单页打印",
        "【报告制作】工作区（普通患者 & 仅标注“需免疫组化”的原报告）",
        [
            ("显微摄片与命名", "显微镜拍摄后将图片存入 <code>未归档报告\\病理照片\\病理号_姓名\\</code>，推荐命名为 <code>1_诊断.jpg</code> / <code>1.2.诊断.jpg</code>。"),
            ("图片排版与模板填入", "左侧按病理号正序选定患者；使用【临床照片查询】对照皮损照；点选图片插槽或【一键自动排版】，单图诊断默认去序号、多图保留序号；结合顶部推荐芯片或 249 篇词库弹窗写入镜下描述。"),
            ("特殊标注勾选判断", "仅勾选 <b>[需免疫组化]</b> 时：<b>视同普通患者</b>处理初版原报告，保存时<b>不会</b>生成特殊归档目录；若需上级把关或切片复查，再勾选 <b>[待审核]</b> 或 <b>[需重切]</b>。"),
            ("保存、归档与单页打印", "点击【保存报告并输出】写入 <code>未归档报告\\病理报告输出\\</code>；若未勾选待审核/重切，可直接点击【一键归档】平铺存入 <code>已归档报告\\病理号_姓名_诊断.docx</code>，并点【一键打印第一页】调用 Word COM 仅打第 1 页。")
        ],
        "Word/WPS 若正打开目标报告将被占锁保护安全拦截（HTTP 409），请先关闭对应 Word 窗口再保存或归档。"
    )

    add_sop_block(
        "SOP-04",
        "待审核与重切病例质控审核规范（两者完全等价）",
        "【报告制作】->【报告审核】-> 返回【报告制作】归档闭环",
        [
            ("触发分流（制作区保存）", "在制作区勾选 <code>待审核</code> 和/或 <code>需重切</code> 后保存：系统写入 <code>未归档报告\\病理报告输出\\</code>，并仅复制该报告草稿与 <code>病理照片\\</code> 到 <code>特殊患者归档\\待审核（或重切）\\&lt;患者目录&gt;\\</code>（<b>绝不复制临床照片</b>）。此时制作区禁止直接归档。"),
            ("审核区照片调阅规则", "进入【报告审核】（不显示纯免疫组化患者）：点击【临床照片查询】打开 <code>未归档报告\\临床照片\\</code>；如需补充重切或复核照片，点击界面【病理照片目录】或【复制病理照片路径】获取<b>复制后目录</b>（<code>特殊患者归档\\待审核(或重切)\\...\\病理照片\\</code>），放图后点【重新读取照片】。"),
            ("双草稿双向完全同步", "在审核区编辑报告医生、审核医生、报告日期、照片插槽、镜下所见或病理诊断后点击【保存修改并同步草稿】：新增照片自动回同步至 <code>未归档报告\\病理照片\\</code> 防止后续丢图，且同时同步更新 <code>未归档报告\\病理报告输出\\</code> 与 <code>特殊患者归档\\</code> 两份草稿（诊断更名时自动清理旧诊断文件/目录）。在制作区修改保存同样反向同步至特殊归档草稿。"),
            ("完成审核与回制作区归档", "审核确认无误后点击【完成审核】。无论首次归档，还是对<b>已归档病例</b>在审核区保存修订（状态置为 <code>archived=2</code>，审核保存不直接改动已归档目录），均须切换到【报告制作】点击【一键归档】：系统正式更新 <code>已归档报告\\</code>，并将待审目录移入 <code>特殊患者归档\\完成审核\\</code>。")
        ],
        "审核区保存绝不直接改写 已归档报告\\！已归档病例经审核修订后必须回到【报告制作】执行重新归档，方可对外打印新版报告。"
    )

    add_sop_block(
        "SOP-05",
        "独立免疫组化报告创建与归档",
        "【免疫组化】工作区（同一患者的第二份独立报告）",
        [
            ("前置条件核验", "确保患者的第一份原初诊报告已完成【已归档报告】归档（若原报告兼有待审核/重切，须先走完 SOP-04 完成归档）。原报告未归档时点击创建将被系统拦截提示。"),
            ("创建免疫组化报告", "在【免疫组化】列表选中患者，点击【创建免疫组化报告】：系统从 <code>已归档报告\\</code> 复制原报告与 <code>未归档报告\\病理照片\\</code> 到 <code>工作目录\\特殊患者归档\\免疫组化\\&lt;原归档名&gt;\\</code>（<b>不复制临床照片</b>）。"),
            ("独立编辑与组化照片录入", "点击【打开组化照片目录】或【复制组化照片路径】放入免疫组化染色新照片，点击【重新读取照片】刷新；独立修改报告/审核医生、日期、镜下所见（支持词库弹窗与推荐）与诊断（临床照片查询仍打开 <code>未归档报告\\临床照片\\</code>）。"),
            ("独立保存与独立归档", "点击【保存免疫组化报告】仅修改 <code>特殊患者归档\\免疫组化\\</code> 内草稿与 <code>ihc_reports</code> 表，<b>完全不动原报告（未归档与已归档原件均不动）</b>；点击【归档免疫组化报告】仅把 Word 复制到 <code>已归档报告\\&lt;原文件名&gt;_免疫组化报告.docx</code>（不覆写原报告，且<b>不移动</b> <code>特殊患者归档\\免疫组化\\</code> 目录，不放入 <code>完成审核\\</code>）。")
        ],
        "原报告与免疫组化报告在逻辑和文件上完全视作两份报告！如已归档的组化报告再次修改保存，须在【免疫组化】页重新点击【归档】更新 *_免疫组化报告.docx。"
    )

    add_sop_block(
        "SOP-06",
        "报告输出三列表、签名单与查询",
        "【报告输出】、【报告领取签名单】与【病历综合查询】",
        [
            ("三列表精准分发打印", "【报告输出】三栏并排隔离：① <b>常规已归档</b>（含普通患者与仅需免疫组化的原报告）；② <b>特殊审核完成</b>（待审/重切完成审核）；③ <b>免疫组化已归档</b>（<code>*_免疫组化报告.docx</code>）。支持按报告/送检日期和前缀过滤、单独或批量调用 Word COM 仅打第 1 页。"),
            ("独立撤销归档 / 撤销审核", "列表 1/2 撤销仅回退原报告归档或审核队列（<b>绝不误删</b> <code>*_免疫组化报告.docx</code> 与 <code>特殊患者归档\\免疫组化\\</code>）；列表 3 撤销仅移出 <code>*_免疫组化报告.docx</code>（保留原报告和组化草稿目录）。"),
            ("生成与打印签名单", "在【报告领取签名单】按送检/报告日期区间与多前缀生成 Excel + 无空行 Word 签名单至 <code>领取表\\</code>（Word 自动标注前一工作日参考手术日期），点击打印直接调用 Word 输出。"),
            ("综合查询与安全删除", "在【病历综合查询】按状态/日期组合查件，点【定位】可直达原报告及独立组化副本路径；删除病例统一通过 Windows 回收站软删除，保障数据可追溯恢复。")
        ],
        "所有打印操作均强制经 Word/WPS COM 引擎执行单页硬件控制；切勿使用浏览器打印替代 Word 排版打印。"
    )

    # Layout Control Section
    story.append(KeepTogether([
        Paragraph("三、排版控制与“单页打印保护”规范", s_h1),
        Paragraph("为防止第二页临床病史被打印浪费耗材，引擎锁定 15 磅固定行距并按图片数量自动调节建议行数，确保第一页严丝合缝止步于医生签名栏：", s_body),
        Image(diag3, width=W, height=134),
        Spacer(1, 8),
    ]))

    layout_table_data = [
        [Paragraph("图片规格", s_tbl_hdr), Paragraph("图片布局", s_tbl_hdr), Paragraph("镜下描述可用行数", s_tbl_hdr), Paragraph("病理诊断可用行数", s_tbl_hdr), Paragraph("第一页收尾位置", s_tbl_hdr)],
        [Paragraph("1 张图 (单图)", s_tbl), Paragraph("居中大图 (1行)", s_tbl), Paragraph("15 行 (含标题)", s_tbl), Paragraph("10 行 (含标题)", s_tbl), Paragraph("精确收于报告日期", s_tbl)],
        [Paragraph("2 张图 (双图)", s_tbl), Paragraph("左右并排 (1行)", s_tbl), Paragraph("14 行 (含标题)", s_tbl), Paragraph("10 行 (含标题)", s_tbl), Paragraph("精确收于报告日期", s_tbl)],
        [Paragraph("3 张图 (三图)", s_tbl), Paragraph("单行三图紧凑", s_tbl), Paragraph("15 行 (含标题)", s_tbl), Paragraph("10 行 (含标题)", s_tbl), Paragraph("精确收于报告日期", s_tbl)],
        [Paragraph("4 - 6 张图", s_tbl), Paragraph("双行阵列 (2行)", s_tbl), Paragraph("10 行 (含标题)", s_tbl), Paragraph("6 行 (含标题)", s_tbl), Paragraph("压缩空行，首页保全", s_tbl)],
        [Paragraph("7 - 9 张图", s_tbl), Paragraph("九宫阵列 (3行)", s_tbl), Paragraph("5 行 (含标题)", s_tbl), Paragraph("3 行 (含标题)", s_tbl), Paragraph("极度紧凑，绝不超页", s_tbl)]
    ]
    lt_flow = Table(layout_table_data, colWidths=[90, 95, 108, 108, 94])
    lt_flow.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1e3a8a")),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor("#f8fafc"), colors.HexColor("#ffffff")]),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    story.append(lt_flow)

    story.append(Spacer(1, 10))
    story.append(Paragraph("四、常见问题与排错指南 (FAQ)", s_h1))

    faq_items = [
        ("Q1：为什么“待审核”列表里看不到仅勾选了“需免疫组化”的患者？",
         "答：新版将免疫组化拆分为独立模块（顶部【免疫组化】标签）。仅勾选“需免疫组化”的患者先按常规流程在【报告制作】归档和打印原初诊报告，再进入【免疫组化】创建第二份组化报告；只有勾选了“待审核”或“需重切”的患者才进入【报告审核】。"),
        ("Q2：在【免疫组化】点击“创建免疫组化报告”提示原报告尚未归档怎么办？",
         "答：免疫组化报告必须在患者原报告已归档于 <code>已归档报告\\</code> 后方可创建副本。请先在【报告制作】（或完成待审后回【报告制作】）将原报告点【一键归档】，然后再创建免疫组化报告。"),
        ("Q3：我在【报告审核】里新放入了重切/复核病理照片，为什么【报告制作】还能同步看到、归档不会丢图？",
         "答：审核区的打开/复制病理照片按钮指向复制后的 <code>特殊患者归档\\待审核(或重切)\\...\\病理照片\\</code>；在审核区点【保存修改并同步草稿】时，系统会自动将新增照片回同步至 <code>未归档报告\\病理照片\\</code>，并保持两处 Word 草稿完全一致。"),
        ("Q4：为什么已归档的报告在【报告审核】改完保存后，还要求去【报告制作】重新归档？",
         "答：为防止静默篡改已封存档案，审核保存仅同步两份草稿并将归档状态标为 <code>需重新归档(2)</code>，不直接覆写 <code>已归档报告\\</code>；需前往【报告制作】核验后点【一键归档】正式更新归档文件。"),
        ("Q5：点击“临床照片查询”提示找不到或存在重名怎么办？",
         "答：待审核/免疫组化均不复制临床照片。系统直接在 <code>未归档报告\\临床照片\\</code> 中检索，单匹配自动打开，重名或未命中则打开母目录由操作员人工确认。"),
    ]
    for q, a in faq_items:
        story.append(KeepTogether([
            Paragraph(f"<b>{q}</b>", s_bold),
            Paragraph(a, s_body),
            Spacer(1, 3)
        ]))

    story.append(Spacer(1, 12))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#e2e8f0"), spaceAfter=8))

    footer_table = Table([[
        Paragraph("<b>编制单位：</b> 皮肤病理系统研发团队<br/><b>部署支持：</b> 离线绿包 / Win7 专版", s_tbl),
        Paragraph("<b>培训脱敏声明：</b><br/>本文档内人名（张医生、李主任、张建国等）及号段均系脱敏教学示例，无真实机构或个人隐私。", s_tbl),
        Paragraph("<b>版本标识：</b> v1.4.0-SOP<br/><b>文档归档：</b> <code>文档/</code>", s_tbl)
    ]], colWidths=[150, 220, 125])
    footer_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#f8fafc")),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor("#cbd5e1")),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    story.append(footer_table)

    doc.build(story, canvasmaker=NumberedCanvas)
    print(f"[OK] PDF 生成成功: {pdf_path}")

    try:
        for f in [diag1, diag2, diag3]:
            if os.path.exists(f):
                os.remove(f)
        if os.path.exists(img_dir):
            os.rmdir(img_dir)
    except Exception as e:
        print(f"清理临时图片提示: {e}")


if __name__ == "__main__":
    target_pdf = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "文档",
        "皮肤病理报告系统_快速入门操作手册.pdf"
    )
    build_manual_pdf(target_pdf)
