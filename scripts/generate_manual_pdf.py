# -*- coding: utf-8 -*-
"""
生成皮肤病理报告系统快速入门操作手册 (PDF 格式, 含高清示意图)
严格执行脱敏处理，排版美观、简单易懂，适合初学者/新入职医护人员
"""

import os
import sys
from PIL import Image as PILImage, ImageDraw, ImageFont
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak, KeepTogether, HRFlowable
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

# 1. 注册微软雅黑中文字体
FONT_REGULAR = "C:/Windows/Fonts/msyh.ttc"
FONT_BOLD = "C:/Windows/Fonts/msyhbd.ttc"

pdfmetrics.registerFont(TTFont("YaHei", FONT_REGULAR))
pdfmetrics.registerFont(TTFont("YaHeiBold", FONT_BOLD))

# -------------------------------------------------------------
# 绘图工具函数：生成 2x 高清示意图
# -------------------------------------------------------------
def get_pil_font(size, bold=False):
    font_path = FONT_BOLD if bold else FONT_REGULAR
    return ImageFont.truetype(font_path, size)

def generate_diagram_workflow(out_path):
    """生成日常5步工作流高清图"""
    w, h = 1600, 380
    im = PILImage.new("RGB", (w, h), (248, 250, 252))
    draw = ImageDraw.Draw(im)

    f_title = get_pil_font(26, bold=True)
    f_step = get_pil_font(22, bold=True)
    f_desc = get_pil_font(18, bold=False)

    steps = [
        ("01", "患者登记录入", "输入基本信息与病史\n规划图片数(1-9图)\n生成初稿并建照片夹", (59, 130, 246)),
        ("02", "显微摄片导入", "显微镜拍摄病理照片\n存入患者专属文件夹\n以序号+诊断规范命名", (16, 185, 129)),
        ("03", "报告制作排版", "一键调阅临床大体照\n装配显微图与智能词库\n自适应排版行数对齐", (139, 92, 246)),
        ("04", "质控审核把关", "常规病例直接归档\n待审核/免疫组化送审\n审核医师双签确认", (245, 158, 11)),
        ("05", "极速单页打印", "双列表严格安全隔离\n仅打印第一页防废纸\n一键生成签领单交接", (239, 68, 68)),
    ]

    card_w = 270
    card_h = 280
    gap = 40
    start_x = 45
    start_y = 50

    for i, (num, title, desc, color) in enumerate(steps):
        x = start_x + i * (card_w + gap)
        y = start_y

        # 绘制阴影与卡片背景
        draw.rounded_rectangle([x + 3, y + 3, x + card_w + 3, y + card_h + 3], radius=16, fill=(226, 232, 240))
        draw.rounded_rectangle([x, y, x + card_w, y + card_h], radius=16, fill=(255, 255, 255), outline=color, width=3)

        # 头部色块
        draw.rounded_rectangle([x, y, x + card_w, y + 68], radius=14, fill=color)
        draw.rectangle([x, y + 50, x + card_w, y + 68], fill=color) # 消除下圆角

        # 序号与标题
        draw.text((x + 16, y + 16), f"步骤 {num}", font=f_step, fill=(255, 255, 255))
        draw.text((x + 20, y + 90), title, font=f_title, fill=(30, 41, 59))

        # 描述多行
        dy = y + 140
        for line in desc.split("\n"):
            draw.text((x + 20, dy), "• " + line, font=f_desc, fill=(100, 116, 139))
            dy += 34

        # 绘制箭头 (除最后一个外)
        if i < len(steps) - 1:
            ax = x + card_w + 8
            ay = y + card_h // 2
            draw.line([ax, ay, ax + 22, ay], fill=(148, 163, 184), width=4)
            draw.polygon([(ax + 22, ay - 8), (ax + 32, ay), (ax + 22, ay + 8)], fill=(148, 163, 184))

    im.save(out_path, quality=95)

def generate_diagram_special_flow(out_path):
    """生成特殊病例质控闭环流转图"""
    w, h = 1500, 420
    im = PILImage.new("RGB", (w, h), (248, 250, 252))
    draw = ImageDraw.Draw(im)

    f_title = get_pil_font(24, bold=True)
    f_sub = get_pil_font(18, bold=False)
    f_badge = get_pil_font(18, bold=True)

    # 绘制外边框与标题栏
    draw.rounded_rectangle([30, 20, w - 30, h - 20], radius=18, fill=(255, 255, 255), outline=(203, 213, 225), width=2)

    # 节点 1: 制作台
    draw.rounded_rectangle([60, 110, 320, 310], radius=14, fill=(238, 242, 255), outline=(79, 70, 229), width=3)
    draw.text((85, 130), "报告制作工作区", font=f_title, fill=(67, 56, 202))
    draw.text((85, 175), "勾选 [待审核] 或", font=f_sub, fill=(55, 65, 81))
    draw.text((85, 208), "勾选 [免疫组化]", font=f_sub, fill=(55, 65, 81))
    # 红色警示标签
    draw.rounded_rectangle([85, 250, 295, 290], radius=8, fill=(254, 226, 226))
    draw.text((100, 258), "× 严格禁止直接归档", font=f_badge, fill=(220, 38, 38))

    # 连线 -> 自动分流
    draw.line([320, 210, 410, 210], fill=(79, 70, 229), width=4)
    draw.polygon([(410, 202), (425, 210), (410, 218)], fill=(79, 70, 229))

    # 节点 2: 特殊归档双队列
    draw.rounded_rectangle([425, 70, 715, 350], radius=14, fill=(254, 243, 199), outline=(217, 119, 6), width=3)
    draw.text((455, 95), "特殊患者归档队列", font=f_title, fill=(180, 83, 9))
    # 子盒子 1
    draw.rounded_rectangle([455, 145, 685, 225], radius=10, fill=(255, 255, 255), outline=(245, 158, 11), width=2)
    draw.text((475, 160), "特殊患者归档/待审核", font=f_badge, fill=(180, 83, 9))
    draw.text((475, 190), "疑难病例待上级把关", font=f_sub, fill=(107, 114, 128))
    # 子盒子 2
    draw.rounded_rectangle([455, 245, 685, 325], radius=10, fill=(255, 255, 255), outline=(245, 158, 11), width=2)
    draw.text((475, 260), "特殊患者归档/免疫组化", font=f_badge, fill=(180, 83, 9))
    draw.text((475, 290), "等待特染/抗体染色结果", font=f_sub, fill=(107, 114, 128))

    # 连线 -> 审核工作区
    draw.line([715, 210, 805, 210], fill=(217, 119, 6), width=4)
    draw.polygon([(805, 202), (820, 210), (805, 218)], fill=(217, 119, 6))

    # 节点 3: 报告审核工作区
    draw.rounded_rectangle([820, 110, 1080, 310], radius=14, fill=(236, 253, 245), outline=(5, 150, 105), width=3)
    draw.text((845, 130), "报告审核工作区", font=f_title, fill=(4, 120, 87))
    draw.text((845, 175), "主任医师复核片子", font=f_sub, fill=(55, 65, 81))
    draw.text((845, 208), "确认补充诊断意见", font=f_sub, fill=(55, 65, 81))
    # 绿色通过按钮模拟
    draw.rounded_rectangle([845, 250, 1055, 290], radius=8, fill=(16, 185, 129))
    draw.text((865, 258), "[ 审核通过并归档 ]", font=f_badge, fill=(255, 255, 255))

    # 连线 -> 终审入库
    draw.line([1080, 210, 1170, 210], fill=(5, 150, 105), width=4)
    draw.polygon([(1170, 202), (1185, 210), (1170, 218)], fill=(5, 150, 105))

    # 节点 4: 终审闭环与打印
    draw.rounded_rectangle([1185, 70, 1445, 350], radius=14, fill=(240, 253, 250), outline=(13, 148, 136), width=3)
    draw.text((1205, 95), "闭环归档与放行", font=f_title, fill=(15, 118, 110))
    # 成果 1
    draw.rounded_rectangle([1205, 145, 1425, 225], radius=10, fill=(255, 255, 255), outline=(20, 184, 166), width=2)
    draw.text((1220, 160), "移入: 完成审核目录", font=f_badge, fill=(15, 118, 110))
    draw.text((1220, 190), "同步归档至按患者归档", font=f_sub, fill=(107, 114, 128))
    # 成果 2
    draw.rounded_rectangle([1205, 245, 1425, 325], radius=10, fill=(255, 255, 255), outline=(20, 184, 166), width=2)
    draw.text((1220, 260), "报告输出: 特殊列表", font=f_badge, fill=(15, 118, 110))
    draw.text((1220, 290), "解锁第一页安全打印", font=f_sub, fill=(107, 114, 128))

    im.save(out_path, quality=95)

def generate_diagram_page_layout(out_path):
    """生成版面控制与首页单页打印示意图"""
    w, h = 1400, 400
    im = PILImage.new("RGB", (w, h), (248, 250, 252))
    draw = ImageDraw.Draw(im)

    f_title = get_pil_font(24, bold=True)
    f_sub = get_pil_font(18, bold=False)
    f_badge = get_pil_font(18, bold=True)
    f_tag = get_pil_font(16, bold=True)

    # 绘制模拟第一页
    p1_x, p1_y, p1_w, p1_h = 100, 40, 450, 320
    draw.rounded_rectangle([p1_x+4, p1_y+4, p1_x+p1_w+4, p1_y+p1_h+4], radius=12, fill=(226, 232, 240))
    draw.rounded_rectangle([p1_x, p1_y, p1_x+p1_w, p1_y+p1_h], radius=12, fill=(255, 255, 255), outline=(59, 130, 246), width=2)
    # 标题栏
    draw.rectangle([p1_x, p1_y, p1_x+p1_w, p1_y+36], fill=(59, 130, 246))
    draw.text((p1_x + 15, p1_y + 8), "【第 1 页】终审法定报告单 (必打页)", font=f_badge, fill=(255, 255, 255))

    # 内容块
    draw.rectangle([p1_x+20, p1_y+50, p1_x+p1_w-20, p1_y+85], fill=(241, 245, 249), outline=(203, 213, 225))
    draw.text((p1_x+30, p1_y+58), "表头: 病理号 / 姓名 / 性别 / 年龄 / 送检部位", font=f_sub, fill=(71, 85, 105))

    draw.rectangle([p1_x+20, p1_y+95, p1_x+p1_w-20, p1_y+165], fill=(238, 242, 255), outline=(199, 210, 254))
    draw.text((p1_x+30, p1_y+105), "显微图像排版区 (1-9图自动行数补偿)", font=f_badge, fill=(67, 56, 202))
    draw.text((p1_x+30, p1_y+135), "图片下方精确对齐部位标注 (1.部位名)", font=f_sub, fill=(100, 116, 139))

    draw.rectangle([p1_x+20, p1_y+175, p1_x+p1_w-20, p1_y+245], fill=(240, 253, 244), outline=(187, 247, 208))
    draw.text((p1_x+30, p1_y+185), "镜下描述与病理诊断 (固定15磅行距)", font=f_badge, fill=(22, 101, 52))
    draw.text((p1_x+30, p1_y+215), "自动填充空行, 防止尾部内容溢出到下页", font=f_sub, fill=(100, 116, 139))

    draw.rectangle([p1_x+20, p1_y+255, p1_x+p1_w-20, p1_y+295], fill=(254, 243, 199), outline=(252, 211, 77))
    draw.text((p1_x+30, p1_y+265), "署名栏: 报告医生 / 审核医生 / 报告日期", font=f_badge, fill=(180, 83, 9))

    # 中间打印流向
    draw.line([600, 200, 720, 200], fill=(37, 99, 235), width=6)
    draw.polygon([(720, 190), (740, 200), (720, 210)], fill=(37, 99, 235))
    draw.rounded_rectangle([590, 130, 730, 175], radius=8, fill=(37, 99, 235))
    draw.text((605, 142), "一键极速打印", font=f_badge, fill=(255, 255, 255))
    draw.text((600, 220), "只打印第1页", font=f_title, fill=(220, 38, 38))
    draw.text((585, 255), "自动忽略第2页附页", font=f_sub, fill=(100, 116, 139))

    # 绘制模拟第二页
    p2_x, p2_y, p2_w, p2_h = 780, 40, 450, 320
    draw.rounded_rectangle([p2_x+4, p2_y+4, p2_x+p2_w+4, p2_y+p2_h+4], radius=12, fill=(226, 232, 240))
    draw.rounded_rectangle([p2_x, p2_y, p2_x+p2_w, p2_y+p2_h], radius=12, fill=(255, 255, 255), outline=(148, 163, 184), width=2)
    # 标题栏
    draw.rectangle([p2_x, p2_y, p2_x+p2_w, p2_y+36], fill=(148, 163, 184))
    draw.text((p2_x + 15, p2_y + 8), "【第 2 页】临床病史附页 (电子存档专用)", font=f_badge, fill=(255, 255, 255))

    # 内容块
    draw.rectangle([p2_x+20, p2_y+60, p2_x+p2_w-20, p2_y+200], fill=(248, 250, 252), outline=(226, 232, 240))
    draw.text((p2_x+30, p2_y+75), "详细病史资料与临床特征摘要:", font=f_badge, fill=(71, 85, 105))
    draw.text((p2_x+30, p2_y+115), "• 电子档案完整保留，永久可查", font=f_sub, fill=(100, 116, 139))
    draw.text((p2_x+30, p2_y+150), "• 纸质打印默认跳过，杜绝纸张浪费", font=f_sub, fill=(100, 116, 139))

    # 红色打叉标签
    draw.rounded_rectangle([p2_x+80, p2_y+240, p2_x+370, p2_y+290], radius=10, fill=(254, 226, 226), outline=(239, 68, 68), width=2)
    draw.text((p2_x+105, p2_y+252), "免打印附页 · 自动节约用纸", font=f_title, fill=(220, 38, 38))

    im.save(out_path, quality=95)


# -------------------------------------------------------------
# ReportLab 带页眉页脚 Canvas
# -------------------------------------------------------------
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
        self.setFont("YaHei", 9)
        self.setFillColor(colors.HexColor("#64748b"))

        # 页眉 (仅第 2 页及之后显示)
        if self._pageNumber > 1:
            self.drawString(54, 800, "皮肤病理报告系统 — 快速入门操作手册 (脱敏图文版)")
            self.setStrokeColor(colors.HexColor("#e2e8f0"))
            self.setLineWidth(0.5)
            self.line(54, 792, 541, 792)

        # 页脚
        page_str = f"第 {self._pageNumber} 页 / 共 {page_count} 页"
        self.drawRightString(541, 36, page_str)
        self.drawString(54, 36, "内部操作规程 · 严守质控安全 · 绿色离线部署")
        self.setStrokeColor(colors.HexColor("#e2e8f0"))
        self.setLineWidth(0.5)
        self.line(54, 48, 541, 48)

        self.restoreState()


# -------------------------------------------------------------
# 主编排逻辑
# -------------------------------------------------------------
def build_manual_pdf(pdf_path):
    # 先生成高清插图
    work_dir = os.path.dirname(pdf_path)
    img_dir = os.path.join(work_dir, ".temp_imgs")
    os.makedirs(img_dir, exist_ok=True)

    diag1 = os.path.join(img_dir, "diag1_workflow.png")
    diag2 = os.path.join(img_dir, "diag2_special.png")
    diag3 = os.path.join(img_dir, "diag3_layout.png")

    generate_diagram_workflow(diag1)
    generate_diagram_special_flow(diag2)
    generate_diagram_page_layout(diag3)

    # 建立 PDF 文档 (A4 边距 54 pt ≈ 1.9 cm)
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        leftMargin=54,
        rightMargin=54,
        topMargin=54,
        bottomMargin=54
    )

    styles = getSampleStyleSheet()

    # 自定义样式
    s_title = ParagraphStyle(
        "DocTitle",
        fontName="YaHeiBold",
        fontSize=24,
        leading=32,
        textColor=colors.HexColor("#0f172a"),
        alignment=1, # 居中
        spaceAfter=8
    )

    s_subtitle = ParagraphStyle(
        "DocSubtitle",
        fontName="YaHei",
        fontSize=12,
        leading=18,
        textColor=colors.HexColor("#475569"),
        alignment=1,
        spaceAfter=15
    )

    s_badge_bar = ParagraphStyle(
        "BadgeBar",
        fontName="YaHei",
        fontSize=9,
        leading=14,
        textColor=colors.HexColor("#0284c7"),
        alignment=1,
        spaceAfter=20
    )

    s_h1 = ParagraphStyle(
        "Heading1_Custom",
        fontName="YaHeiBold",
        fontSize=15,
        leading=22,
        textColor=colors.HexColor("#1e3a8a"),
        spaceBefore=14,
        spaceAfter=8,
        keepWithNext=True
    )

    s_h2 = ParagraphStyle(
        "Heading2_Custom",
        fontName="YaHeiBold",
        fontSize=12,
        leading=18,
        textColor=colors.HexColor("#0369a1"),
        spaceBefore=10,
        spaceAfter=5,
        keepWithNext=True
    )

    s_body = ParagraphStyle(
        "Body_Custom",
        fontName="YaHei",
        fontSize=9.5,
        leading=15,
        textColor=colors.HexColor("#334155"),
        spaceAfter=6
    )

    s_body_bold = ParagraphStyle(
        "BodyBold_Custom",
        fontName="YaHeiBold",
        fontSize=9.5,
        leading=15,
        textColor=colors.HexColor("#1e293b"),
        spaceBefore=6,
        spaceAfter=2,
        keepWithNext=True
    )

    s_tip = ParagraphStyle(
        "Tip_Custom",
        fontName="YaHei",
        fontSize=9,
        leading=14,
        textColor=colors.HexColor("#0f766e")
    )

    s_warn = ParagraphStyle(
        "Warn_Custom",
        fontName="YaHeiBold",
        fontSize=9,
        leading=14,
        textColor=colors.HexColor("#b91c1c")
    )

    s_table_text = ParagraphStyle(
        "TableText",
        fontName="YaHei",
        fontSize=8.5,
        leading=12,
        textColor=colors.HexColor("#334155")
    )

    s_table_hdr = ParagraphStyle(
        "TableHdr",
        fontName="YaHeiBold",
        fontSize=9,
        leading=13,
        textColor=colors.HexColor("#ffffff")
    )

    story = []

    # ---------------- 封面 / 标题区 ----------------
    story.append(Paragraph("皮肤病理报告系统", s_title))
    story.append(Paragraph("快速入门操作手册 (图文教程 · v1.2 脱敏版)", s_subtitle))
    story.append(Paragraph("【适用对象：皮肤科医师 / 报告录入员 / 质控主任 / 规培进修生】&nbsp;&nbsp;|&nbsp;&nbsp;【全流程单页精准打印 · 绿色免安装】", s_badge_bar))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#3b82f6"), spaceAfter=15))

    # 提示框：开箱即用
    callout_data = [[
        Paragraph("<b>★ 新手快速导读：</b> 本手册采用图文通俗讲解。日常出具报告仅需掌握 <b>“录入 ➔ 摄片 ➔ 排版 ➔ 审核 ➔ 打印”</b> 五个核心动作。遇到 <b>“待审核”</b> 或 <b>“免疫组化”</b> 疑难病例时，系统会自动启动质控安全锁，请务必经由审核医生复核放行后再行出单。", s_tip)
    ]]
    callout_table = Table(callout_data, colWidths=[487])
    callout_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#f0fdfa")),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#99f6e4")),
        ('LEFTPADDING', (0,0), (-1,-1), 12),
        ('RIGHTPADDING', (0,0), (-1,-1), 12),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(callout_table)
    story.append(Spacer(1, 12))

    # ---------------- 第一部分：系统启动与主界面 ----------------
    story.append(Paragraph("一、极速起步：系统一键启动与初始化", s_h1))
    story.append(Paragraph("本系统为完全独立的<b>绿色免安装软件包</b>，出厂自带运行时，不写入系统注册表，即解压即用：", s_body))

    start_steps = [
        "<b>1. 双击启动：</b> 在软件根目录下直接双击 <code>启动系统.bat</code>（或双击 <code>启动系统(无黑框后台版).vbs</code>）。系统自动调起默认浏览器打开操作主页（<code>http://127.0.0.1:8765</code>）。",
        "<b>2. 初次设置向导：</b> 首次打开网页时，根据弹窗提示确认 <b>工作目录</b> 路径（建议设在 D 盘或大容量数据盘），并输入您的 <b>报告医生姓名</b> 与 <b>审核医生姓名</b>（后续可在系统设置中随时修改）。",
        "<b>3. 优雅退出：</b> 结束全天工作后，双击根目录下的 <code>停止系统.bat</code> 即可安全平稳关闭后台常驻服务。"
    ]
    for step in start_steps:
        story.append(Paragraph(step, s_body))

    story.append(Spacer(1, 10))

    # ---------------- 第二部分：日常病理报告五步标准工作流 ----------------
    story.append(Paragraph("二、日常病理报告五步标准工作流 (图解)", s_h1))
    story.append(Paragraph("科室日常标本处理严格遵循标准化流水线作业，全过程清晰流转如下图所示：", s_body))

    # 插入流程图 1
    story.append(Image(diag1, width=487, height=115))
    story.append(Spacer(1, 10))

    # 详细步骤拆解
    story.append(Paragraph("步骤 1：患者信息登记与模板初始化 (【患者信息录入】工作区)", s_h2))
    story.append(Paragraph("• <b>录入表单：</b> 准确填入患者姓名、性别、年龄、病理号、门诊/住院号、送检科室、送检医生、临床诊断与临床病史。", s_body))
    story.append(Paragraph("• <b>病理号前缀规范：</b> 普通门诊不带前缀（如 <code>2026110301</code>）；住院病房请选 <code>B</code>（如 <code>B2026110301</code>）；分院区选 <code>L</code>；会诊选 <code>H</code>；特殊标本选 <code>T</code>。<b>不同前缀当日自增序号独立递增，互不干扰</b>。", s_body))
    story.append(Paragraph("• <b>规划图片数与部位联动：</b> 估算该标本需要拍摄的照片数量（1至9张）。选定后，下方会自动生成相应数量的“送检部位”输入框。在生成报告时，每个部位将以 <code>1.部位名</code>、<code>2.部位名</code> 规范对齐于对应图片正下方。", s_body))
    story.append(Paragraph("• <b>一键生成：</b> 点击【录入并生成报告】。系统自动在 <code>统一归档/病理报告录入/</code> 生成初稿 Word，并自动在 <code>统一归档/病理照片/</code> 为该患者建立专属摄片目录（<code>病理号_姓名</code>）。", s_body))

    story.append(Spacer(1, 8))
    story.append(Paragraph("步骤 2：显微镜摄片与图片存入", s_h2))
    story.append(Paragraph("• 在显微镜摄片电脑上完成切片拍摄后，将照片直接保存或拷贝至对应患者的 <code>统一归档/病理照片/病理号_姓名/</code> 文件夹中。", s_body))
    story.append(Paragraph("• <b>文件命名建议：</b> 推荐采用 <code>1_基底细胞癌40X.jpg</code> 或 <code>序号+诊断</code> 格式，系统将自动根据文件名中的关键词智能推荐病理描述词库。", s_body))

    story.append(Spacer(1, 8))
    story.append(Paragraph("步骤 3：报告制作、调图与排版 (【报告制作】工作区)", s_h2))
    story.append(Paragraph("• <b>进入编辑：</b> 点击左侧工作列表中的患者卡片（默认列出今日患者，可自由切换日期查询）。", s_body))
    story.append(Paragraph("• <b>临床照片快速调阅：</b> 界面设有【临床照片查询】按钮。点击后，系统自动在 <code>统一归档/临床照片/</code> 检索该患者姓名的皮损大体照片文件夹，若存在则直接唤起资源管理器打开，便于对照临床表现。", s_body))
    story.append(Paragraph("• <b>图片点选装配：</b> 界面以清晰缩略图展示已摄照片，支持一键点选替换 Word 模板中的预留占位框。如拍摄张数少于预设位，排版引擎会自动剥离多余空白框，无需担心出现黑色空白矩形。", s_body))
    story.append(Paragraph("• <b>诊断词库智能推荐：</b> 右侧词库栏自动关联推荐常见病描述（如皮脂腺痣、脂溢性角化、银屑病等），点击【一键填入】即可将标准镜下描述与诊断直接填入文本框。", s_body))
    story.append(Paragraph("• <b>智能排版保存：</b> 点击【保存报告】。系统自动按行距15磅和图片规格重算行高补偿，确保<b>第一页完整收尾于医生签名栏</b>，并存入 <code>统一归档/病理报告输出/</code>。", s_body))

    story.append(Spacer(1, 10))

    # ---------------- 第三部分：特殊疑难质控与归档流转 ----------------
    story.append(Paragraph("步骤 4：质控审核机制与特殊病例流转 (【报告审核】工作区)", s_h2))
    story.append(Paragraph("为杜绝未经上级确认的疑难报告或未出免疫组化结果的半成品报告误发患者，系统设置了严密的质控安全锁：", s_body))

    warn_box = [[
        Paragraph("<b>！质控红线原则：</b><br/>"
                  "1. 凡勾选 <b>[待审核]</b> 或 <b>[免疫组化]</b> 的报告，在制作台<b>绝对禁止直接归档</b>（点击归档将弹出警示拦截）。<br/>"
                  "2. 保存时，系统自动将报告派发到 <code>特殊患者归档/待审核</code> 或 <code>特殊患者归档/免疫组化</code> 待查队列中。<br/>"
                  "3. 必须由审核医师在【报告审核】工作区完成复核并点击【审核通过并归档】，系统才会将其转移至 <code>特殊患者归档/完成审核</code> 并同步入库 <code>按患者归档</code>，解除打印锁定！", s_warn)
    ]]
    w_table = Table(warn_box, colWidths=[487])
    w_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#fff1f2")),
        ('BOX', (0,0), (-1,-1), 1.5, colors.HexColor("#f43f5e")),
        ('LEFTPADDING', (0,0), (-1,-1), 12),
        ('RIGHTPADDING', (0,0), (-1,-1), 12),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
    ]))

    # 插入特殊病例流程图
    story.append(KeepTogether([
        Image(diag2, width=487, height=136),
        Spacer(1, 8),
        w_table
    ]))

    story.append(Spacer(1, 10))
    story.append(Paragraph("步骤 5：报告安全输出与批量单页打印 (【报告输出】工作区)", s_h2))
    story.append(Paragraph("【报告输出】专为报告分发台设计，实行严格的<b>双列表物理隔离</b>机制：", s_body))
    story.append(Paragraph("• <b>常规已归档列表：</b> 仅列出普通已归档、且不属于待审核和免疫组化的合规报告。", s_body))
    story.append(Paragraph("• <b>特殊完成审核列表：</b> 仅列出虽然勾选了待审核/免疫组化，但<b>已经上级医师复核签署完毕</b>的报告。", s_body))
    story.append(Paragraph("• <b>一键批量打印今日归档：</b> 点击顶部按钮，系统自动调起打印机，逐份输出今日报告，极速高效。", s_body))

    story.append(Spacer(1, 10))
    story.append(Paragraph("三、排版控制与“单页打印保护”揭秘", s_h1))
    story.append(Paragraph("日常打印中，病史资料常导致报告溢出到第 2 页，浪费大量纸张。本系统研发了独创的<b>单页精准打印保护机制</b>：", s_body))

    layout_table_data = [
        [Paragraph("图片规格", s_table_hdr), Paragraph("图片布局", s_table_hdr), Paragraph("镜下描述可用行数", s_table_hdr), Paragraph("病理诊断可用行数", s_table_hdr), Paragraph("第一页收尾位置", s_table_hdr)],
        [Paragraph("1 张图 (单图)", s_table_text), Paragraph("居中大图 (1行)", s_table_text), Paragraph("14 行 (含标题)", s_table_text), Paragraph("10 行 (含标题)", s_table_text), Paragraph("精确收于报告日期", s_table_text)],
        [Paragraph("2 张图 (双图)", s_table_text), Paragraph("左右并排 (1行)", s_table_text), Paragraph("14 行 (含标题)", s_table_text), Paragraph("8 行 (含标题)", s_table_text), Paragraph("精确收于报告日期", s_table_text)],
        [Paragraph("3 张图 (三图)", s_table_text), Paragraph("紧凑陈列 (1行)", s_table_text), Paragraph("14 行 (含标题)", s_table_text), Paragraph("10 行 (含标题)", s_table_text), Paragraph("精确收于报告日期", s_table_text)],
        [Paragraph("4 - 6 张图", s_table_text), Paragraph("双行阵列 (2行)", s_table_text), Paragraph("10 行 (含标题)", s_table_text), Paragraph("6 行 (含标题)", s_table_text), Paragraph("压缩空行，首页保全", s_table_text)],
        [Paragraph("7 - 9 张图", s_table_text), Paragraph("九宫格阵列 (3行)", s_table_text), Paragraph("4 行 (含标题)", s_table_text), Paragraph("3 行 (含标题)", s_table_text), Paragraph("极度紧凑，绝不超页", s_table_text)]
    ]
    lt_flow = Table(layout_table_data, colWidths=[90, 95, 105, 105, 92])
    lt_flow.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1e3a8a")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.HexColor("#f8fafc"), colors.HexColor("#ffffff")]),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))

    # 插入版面图与表格
    story.append(KeepTogether([
        Image(diag3, width=487, height=139),
        Spacer(1, 8),
        lt_flow
    ]))

    story.append(Spacer(1, 10))

    # ---------------- 第四部分：配套工具与高频操作 ----------------
    story.append(Paragraph("四、实用配套工具：签名单生成与综合查询", s_h1))

    story.append(Paragraph("1. 报告领取签领单自动化生成 (【报告领取签名单】工作区)", s_h2))
    story.append(Paragraph("• <b>自由筛选：</b> 可自由设定出具报告的起始日期与截止日期；支持多选前缀（如同时勾选门诊与病房 B）。", s_body))
    story.append(Paragraph("• <b>自动制表：</b> 系统即刻汇总明细列表，点击【生成签领单并打印】，自动生成规范带留样签名字段的 Word 文档，并即时送达打印机，交接清晰有据可查。", s_body))

    story.append(Spacer(1, 8))
    story.append(Paragraph("2. 病例综合查询与穿透式维护 (【病例综合查询】工作区)", s_h2))
    story.append(Paragraph("• <b>全要素组合检索：</b> 支持模糊搜索姓名、病理号、住院号、门诊号；可按“未归档 / 已归档 / 需要重新归档 / 待审核 / 完成审核”状态一键组合过滤。", s_body))
    story.append(Paragraph("• <b>穿透式文件夹定位：</b> 对查询出的任意病例，支持在卡片中直接一键唤起资源管理器，直达其【病理报告输出】、【病理照片文件夹】或【按患者归档文件夹】。", s_body))
    story.append(Paragraph("• <b>信息修正与二次修改保护：</b>", s_body))
    story.append(Paragraph("&nbsp;&nbsp;a) 若仅录入信息打错（如年龄、病理号微调），可直接点击【修改信息】即刻修正生效。<br/>"
                           "&nbsp;&nbsp;b) 若需修改显微照片或镜下诊断，点击【进入制作】重新排版保存。系统会自动将归档状态变更为 <b>“需要重新归档”</b>，避免静默篡改历史已封存档案。经人工核对无误后点击重新归档，方可覆盖归档病历，全方位保护医疗文书严肃性！", s_body))

    story.append(Spacer(1, 12))
    # ---------------- 第五部分：常见问题与排错指南 ----------------
    story.append(Paragraph("五、新手常见问题解答 (FAQ)", s_h1))

    faq_items = [
        ("Q1：点击“临床照片查询”时提示找不到文件夹？",
         "答：系统默认根据“患者姓名”在 <code>统一归档/临床照片/</code> 下查找同名子文件夹。如果临床摄片人员尚未将照片导入，或文件夹名称含有错别字，系统会自动打开临床照片总目录，您可以手动检索或将大体照片放入该目录即可。"),

        ("Q2：为什么在报告制作台点击“归档”按钮弹出了红色警告？",
         "答：这是系统的质控保护机制。该病例勾选了“待审核”或“免疫组化”，属于特殊疑难病例，必须由审核医生在【报告审核】工作区完成复核并点击“审核通过并归档”，普通制作界面禁止越权直接归档。"),

        ("Q3：打印机打出来的报告真的只有第一页吗？第二页病史去哪了？",
         "答：是的！系统底层通过专用单页打印指令，硬件强制仅输出第 1 页。第二页的病史资料完整保存在生成的 Word 文件内部，供电子档案终身查阅追溯，绝不浪费一张纸。"),

        ("Q4：如果医生不小心把报告提前归档了，还能撤销吗？",
         "答：可以。在【报告输出】或【病例综合查询】工作区中，每个病例均配有【撤销归档】按钮。点击后，系统会安全移出归档目录中的文件，并将数据库状态恢复为未归档，允许医生重新制作排版。"),

        ("Q5：病房（B）和普通门诊同一天都录入第01号，病理号会冲突吗？",
         "答：绝对不会！系统对不同前缀（常规、B、L、H、T）实施完全独立的自增流水号机制。<code>2026110301</code> 与 <code>B2026110301</code> 能够完美并存，物理文件夹也天然隔离，请放心使用。")
    ]

    for q, a in faq_items:
        story.append(Paragraph(f"<b>{q}</b>", s_body_bold))
        story.append(Paragraph(a, s_body))
        story.append(Spacer(1, 4))

    story.append(Spacer(1, 15))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#e2e8f0"), spaceAfter=10))

    # 底部脱敏声明与签署栏
    footer_table_data = [
        [
            Paragraph("<b>编制单位：</b> 皮肤病理系统研发团队<br/><b>技术支持：</b> 本地离线部署支持组", s_table_text),
            Paragraph("<b>培训脱敏声明：</b><br/>本文档内所载人名（张医生、李主任、张建国等）及编号均为模拟教学数据，无任何真实个人或机构隐私信息。", s_table_text),
            Paragraph("<b>版本标识：</b> v1.2.0-Release<br/><b>文档归档：</b> <code>工作目录/文档/</code>", s_table_text)
        ]
    ]
    ft_table = Table(footer_table_data, colWidths=[150, 210, 127])
    ft_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#f8fafc")),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#cbd5e1")),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(ft_table)

    # 构建文档
    doc.build(story, canvasmaker=NumberedCanvas)
    print(f"[OK] PDF 生成成功: {pdf_path}")

    # 清理临时图片
    try:
        for f in [diag1, diag2, diag3]:
            if os.path.exists(f):
                os.remove(f)
        if os.path.exists(img_dir):
            os.rmdir(img_dir)
    except Exception as e:
        print(f"清理临时图片提示: {e}")

if __name__ == "__main__":
    target_pdf = r"D:\AIagent\project\dermatology-system\工作目录\文档\皮肤病理报告系统_快速入门操作手册.pdf"
    build_manual_pdf(target_pdf)
