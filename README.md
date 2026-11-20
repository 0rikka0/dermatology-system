<div align="center">

# 🔬 Dermatology Pathology Reporting System
### 皮肤病理报告全流程系统 (v6.0)

*面向皮肤科病理报告辅助生成、多图自适应排版、外部 Word 智能双向同步、质控审核、独立免疫组化与单页打印的本地轻量级开源工作站*

<p align="center">
  <img src="https://img.shields.io/badge/Version-v6.0-2563eb.svg" alt="Version v6.0" />
  <img src="https://img.shields.io/badge/Python-3.8%2B-blue.svg?logo=python&logoColor=white" alt="Python Version" />
  <img src="https://img.shields.io/badge/Platform-Windows%207%20%7C%2010%20%7C%2011-0078D6.svg?logo=windows&logoColor=white" alt="Platform" />
  <img src="https://img.shields.io/badge/Architecture-Local%20Native%20WebUI-orange.svg" alt="Architecture" />
  <img src="https://img.shields.io/badge/Runtime-Portable%20%2F%20Zero--Install-success.svg" alt="Zero Install" />
  <img src="https://img.shields.io/badge/License-MIT-yellow.svg" alt="License" />
</p>

<p align="center">
  <a href="#-系统介绍">系统介绍</a> •
  <a href="#-核心特性">核心特性</a> •
  <a href="#-系统架构与技术栈">系统架构</a> •
  <a href="#-快速开始">快速开始</a> •
  <a href="#-仓库目录结构">目录结构</a> •
  <a href="#-系统详细文档">文档中心</a> •
  <a href="CHANGELOG.md">更新日志</a> •
  <a href="#-数据脱敏合规声明">脱敏声明</a>
</p>

</div>

---

## 📖 系统介绍

**皮肤病理报告系统（v6.0）** 针对皮肤病理从临床取材送检、患者登记与独立流水取号（支持五类标准前缀：无前缀=普通门诊报告、`B`=病房报告、`L`=分院区门诊报告、`H`=会诊报告、`T`=其他特殊报告）、显微镜摄片命名识别、词库描述与多图排版、疑难/重切病例上级复核、独立免疫组化（IHC）第二份报告出具，到外部 Word 协同修改回写、Word COM 仅打第 1 页与患者领取签名单生成等全业务链路定制研发。

系统采用前后端解耦的 **本地原生轻量 WebUI 架构（8 大工作区）**，零外部 CDN 与云端依赖，全部业务数据严格保存于本地工作目录，兼顾极致响应速度、排版精度、数据防损与医疗隐私隔离安全。

---

## ✨ 核心特性

- 🔄 **外部 Word 修改检测、三层隔离双向同步与富文本保留（v6.0 新增）**
  - 系统每次生成或复制报告自动计算路径 **SHA-256 指纹** 存入 `doc_sync_meta`（兼容历史空指纹零误报，自动过滤 `~$` 锁文件）；检测到外部 Word 修改后在界面展示 `⚠️ Word已外改` 徽章，并在查看/编辑/保存/覆盖前触发 `HTTP 409` 拦截提示，由医生确认 **正向导入外部修改** 或 **将外部文件移入回收站备份后覆盖**；
  - 支持 `tables[0]` 12 项字段、`tables[1]` 多位点送检部位、多行正文及签名栏全量逆向解析，配合 **就地局部更新引擎** 原地保留用户在 Word 中添加的文字颜色、高亮样式与审阅批注；
  - 严格隔离 **录入层**、**草稿层（输出/待审核/重切/完成审核）**、**成品层（普通/组化已归档）** 三大层级，组化与普通报告支持公共基础信息跨文档双向同步，并以病理号不匹配拦截错份同步。
- 🛡️ **防静默永删、覆写前回收备份、20 份 WAL 快照与主键唯一归档（v6.0 新增）**
  - 内置 `Send2Trash` + 20s 参数化 PowerShell 回收站双通道；当网络共享盘等环境不支持回收站时 **严禁静默永久删除**，改为抛出明确警报确认框等待人工确认；
  - 全部 Word/Excel 生成、签名单导出、普通/免疫组化归档及重命名移动前统一调用 **先回收旧文件再写入**；每日首次启动及删患前执行 SQLite WAL `backup()` 快照并在 `数据库/backups/` 滚动保留最新 **20 份**；
  - 基于主键 `pathology_no` 唯一匹配管理归档报告（无论立即或延迟重归档、改名/改号），确保 `已归档报告/` 永不遗留旧名重复文件且绝对不触碰 `临床照片/`。
- 🔒 **OS 级无感防多开与后台稳定守护**
  - 集成 `app/instance_lock.py`，结合 **独占端口（`allow_reuse_address=False` + `SO_EXCLUSIVEADDRUSE`）**、**Win32 内核命名互斥体（`CreateMutexW`）**、**`.server_runtime.json` 真实端口追踪** 与 **绕过系统代理 `/api/status` 特征探活** 四步组合防线；
  - 无黑框 `pythonw.exe` 后台运行期间关闭浏览器再次启动时，毫秒级识别既有后台、直接唤起浏览器进入页面并静默退出，绝不重复开启后台进程；支持 `python -m app.instance_lock --stop` 安全关停。
- 📐 **1–9 图自适应排版、签名防溢出与“单页打印保护”**
  - 原生支持 **1 至 9 张** 显微病理照片智能阵列排版，图注与送检部位严格下沉对齐；
  - 锁定 15 磅固定行距并动态适配建议行数（单图 15/10、双图 14/10、三图 15/10、4-6 图 10/6、7-9 图 5/3），签名行采用 Word 左/右制表位及字号微缩自适应（`12pt` 至 `9pt`），确保医生姓名较长时仍稳定收尾于第 1 页；
  - 后台调用 Word/WPS COM 强制 **仅打印法定第 1 页**，第 2 页病史仅留存电子档案。
- 🧪 **独立免疫组化（IHC）双报告与安全删除机制**
  - 仅标注 `需免疫组化` 的患者原初诊报告按常规流程直接制作、批量归档与打印；
  - 在独立的【免疫组化】模块基于已归档原报告创建第二份副本（复制原报告与病理照片至 `特殊患者归档/免疫组化/`，**不复制临床照片**，且新建前禁止静默覆写已有草稿）；
  - 支持完全独立编辑、保存、归档至 `已归档报告/<原文件名>_免疫组化报告.docx`，或通过红色【删除免疫组化报告】将其移入回收站并清空组化标记恢复普通状态（绝不误删原报告）。
- ⚖️ **待审核/重切等价协同、双草稿同步与审核首页打印**
  - 【报告审核】专责处理 `待审核` 与 `需重切` 病例（二者逻辑完全等价），支持边保存边点击【一键打印第一页】；
  - 分流时不复制临床照片，审核区【保存修改并同步草稿】自动回同步新增照片并按草稿同级保真分发同步 `未归档报告/病理报告输出/` 与 `特殊患者归档/`；修改已归档病例自动置 `archived=2`，须回【报告制作】一键归档刷新。
- 🖨 **今日一键批量归档、三列表隔离与跨周末签名单**
  - 【报告制作】支持【一键归档今日已保存未归档报告】（含普通及纯免疫组化，自动跳过待审核/重切）；
  - 【报告输出】物理隔离展示 **常规已归档**、**特殊审核完成**、**免疫组化已归档** 三栏列表，支持单独/批量单页打印及独立撤销；
  - 【报告领取签名单】同步生成 Excel 与无空行 Word 签名单，自动推算参考手术日期（周一提交自动跨越周末展开覆盖周五至周日）；【病历综合查询】支持报告日期与送检日期双区间过滤（默认全部时间）。

---

## 🏗 系统架构与技术栈

```text
+--------------------------------------------------------------------------+
|                          本地 WebUI (客户端呈现层)                        |
|   HTML5 / CSS3 (8 模块响应式布局) + 原生 JavaScript (无外部 CDN 依赖)      |
+--------------------------------------------------------------------------+
                                     │ (HTTP REST API / 127.0.0.1:8765~8774)
                                     ▼
+--------------------------------------------------------------------------+
|                 单实例治理与后端服务层 (Python 3.8+)                      |
|   instance_lock.py (Win32 Mutex + 状态追踪 + 绕过代理探活 + --stop 退出)  |
|   ThreadingHTTPServer (SO_EXCLUSIVEADDRUSE + pythonw 空流防护) + REST API |
+--------------------------------------------------------------------------+
        │                        │                          │
        ▼                        ▼                          ▼
+--------------------+  +------------------------+  +-----------------------+
|   数据持久层       |  |      文档排版引擎       |  |     系统级交互接口     |
| SQLite 3 (WAL)     |  |  python-docx / OpenXML |  | Send2Trash / PS 回收站 |
| 20份备份 + 指纹列  |  |  逆向解析 / 样式保留   |  | COM Word 单页硬件打印 |
| patients + ihc_reports| | 制表位签名防溢出对齐 | |                       |
+--------------------+  +------------------------+  +-----------------------+
```

---

## 📁 仓库目录结构

```plaintext
dermatology-system/
├── app/                         # 后端核心业务逻辑
│   ├── config.py                # 系统配置与初始化管理
│   ├── database.py              # SQLite 双表模型 (含 doc_sync_meta 指纹) 与 20 份 WAL 快照轮转
│   ├── doc_generator.py         # Word 报告生成、1-9 图补偿、就地保留样式/批注与制表位签名对齐
│   ├── doc_preview.py           # A4 高保真预览引擎与全量结构化逆向解析 (extract_docx_details)
│   ├── instance_lock.py         # OS 级命名互斥体、运行时状态追踪、探活与优雅停止
│   ├── path_manager.py          # 防静默永删回收、覆写前备份、三层指纹分检与主键唯一归档
│   ├── server.py                # 本地 HTTP 服务 (SO_EXCLUSIVEADDRUSE)、409 冲突拦截与 REST API
│   ├── signing_sheet_generator.py # 覆写保护版 Excel + 无空行 Word 签领单（跨周末手术日期计算）
│   └── template_manager.py      # 249 篇病理词库解析与诊断自动合成
├── web/                         # 前端原生 Web 资源
│   ├── index.html               # 八模块交互主界面
│   ├── css/style.css            # 医疗质感高对比度三栏输出、外改徽章与路径栏样式
│   └── js/app.js                # 全业务模块交互、外部 Word 冲突确认弹窗与 API 通信
├── 模板/                        # 标准排版骨架 (1图、2图、3-9图模板与 placeholder.jpeg)
├── 文字模板/                    # 专业病理描述词库 (普通模板.docx + 免疫组化模板.docx)
├── 文档/                        # 系统全景架构、逻辑冲突梳理、SOP 图文操作手册 PDF 与脱敏日志
├── 示例文件/                    # 脱敏虚拟患者成品示例报告与签名单示例
├── 工作目录/                    # 标准运行目录占位骨架 (数据库/backups、已归档/未归档/特殊归档等)
├── scripts/                     # 运维辅助与构建脚本
│   ├── build_standalone_package.py # 一键打包绿色免安装发布版（含 Send2Trash、防多开 .vbs）
│   └── generate_manual_pdf.py   # 编译图文标注版 SOP 操作手册 PDF
├── tests/                       # 自动化功能测试套件
├── main.py                      # 系统主程序入口（集成互斥体防多开与无感唤醒）
├── 启动系统.bat                 # Windows 一键启动脚本
├── CHANGELOG.md                 # 版本演进与更新日志 (v1.0 ~ v6.0)
├── requirements.txt             # 源码运行所需第三方依赖
├── .gitignore                   # Git 忽略规则
└── LICENSE                      # 开源许可协议 (MIT)
```

---

## 🚀 快速开始

### 方式一：独立便携版（推荐，适合科室直接使用）
1. 克隆或获取源码后可执行 `python scripts/build_standalone_package.py` 构建免安装便携包，或下载对应发布压缩包；
2. 解压至本地磁盘后，可将包内的 `启动系统(无黑框后台版).vbs` 右键**“发送到 -> 桌面快捷方式”**作为日常唯一入口图标（支持防多开：后台运行中再次双击仅自动打开浏览器，不会多开后台进程）；
3. 也可直接双击 `启动系统.bat` 启动并在默认浏览器打开 `http://127.0.0.1:8765`；需要彻底关停服务时运行 `停止系统.bat` 或执行 `python -m app.instance_lock --stop`。

### 方式二：从源码运行（适合开发者与二次定制）

#### 1. 环境准备
- 操作系统：Windows 7 SP1 / Windows 10 / Windows 11 (64位)
- Python 环境：Python 3.8 及以上版本

#### 2. 克隆仓库与安装依赖
```bash
git clone https://github.com/0rikka0/dermatology-system.git
cd dermatology-system
pip install -r requirements.txt
```

#### 3. 启动与安全停止
```bash
python main.py
# 若需在后台模式下安全停止运行中的本地服务：
python -m app.instance_lock --stop
```

---

## 📖 系统详细文档

系统各业务流程的标准操作规程（SOP）、架构全景、冲突消解机制与历史版本记录已载入以下文档：

- **[皮肤病理报告系统_快速入门操作手册.pdf](文档/皮肤病理报告系统_快速入门操作手册.pdf)**：通俗易懂的 **SOP-01 至 SOP-06 标准操作指南**（聚焦“需要做什么 / 怎么做 / 不能做什么”与五类病理号前缀速查、外改/防永删处理），附全流程示意图与 WebUI 实景按钮标注截图。
- **[皮肤病理报告系统_系统功能与架构全景手册.md](文档/皮肤病理报告系统_系统功能与架构全景手册.md)**：v6.0 全系统技术架构、双表数据字典（含 `doc_sync_meta` 与 20 份快照）、REST API 契约、外部 Word 三层同步、防永删护栏与单实例锁权威参考。
- **[系统逻辑设计与冲突梳理说明.md](文档/系统逻辑设计与冲突梳理说明.md)**：三轨生命周期与 18 大核心业务逻辑/并发冲突的详细消解设计。
- **[更新日志 (CHANGELOG.md)](CHANGELOG.md)**：自 `v1.0` 至 `v6.0` 的完整版本演进记录。
- **[源码脱敏处理记录与审核日志.md](文档/源码脱敏处理记录与审核日志.md)**：v6.0 开源版本全要素脱敏合规审核记录。

---

## 🛡 数据脱敏合规声明

本项目源码、文档、截图及示例数据在公开发布前，已根据相关医疗信息安全规范执行了严格的**全要素脱敏处理**：
- **医疗机构与分院**：真实医院与特定分院区名称均已替换为泛化代称（“某三甲医院皮肤科”、“分院区 / 分院”）；
- **医护人员**：真实医生姓名均已替换为标准代称（“张医生”、“李主任”、“王医生”）；
- **患者隐私**：所有示例数据及教程截图中的患者姓名、门诊号、病理号及送检部位均为系统合成模拟数据；
- **元数据擦除**：Office 文档的核心元数据（创建者、最后修改者、本地绝对路径）均已彻底清除重写。

---

## ⚠️ 免责声明与风险告知 / Disclaimer and Risk Notice

> **请在使用本软件前仔细阅读本声明。您下载、安装、运行或以任何方式继续使用本软件，即视为您已阅读、理解并完全接受本声明的全部内容。**
> **Please read carefully before using this software. By downloading, installing, running, or continuing to use this software, you acknowledge that you have read, understood, and agreed to this disclaimer in full.**

### 🇨🇳 中文版

**1. 项目性质**
本项目（以下简称"本软件"）是一个完全开源的研究、学习与测试项目，**不是**经药品监督管理部门或其他监管机构注册、备案、认证或批准的医疗器械、诊断工具或医疗信息系统。本软件**未在成熟的真实医疗环境中经过充分测试和临床验证**，可能存在功能缺陷、逻辑错误、安全漏洞、兼容性问题及未知风险。

**2. 禁止用于临床诊疗**
本软件**不得**作为临床诊断、治疗决策、病理判读、报告签发或任何涉及患者生命健康的医疗活动的依据或工具。本软件产生的任何内容（包括但不限于病理报告、数据记录、统计结果和 AI 辅助输出）均**仅供参考**，不能替代执业医师和专业人员的独立判断。任何将本软件用于真实医疗环境的行为，均由使用者自行决定并**自行承担全部风险和责任**。

**3. 按"现状"提供，不作担保**
本软件按"现状"和"可用"状态提供，开发者**不作任何明示或默示的担保**，包括但不限于适销性、特定用途适用性、准确性、完整性、可靠性、持续可用性和不侵权的担保。

**4. 风险与责任免除**
在适用法律允许的最大范围内，对于因使用或无法使用本软件而产生的任何直接、间接、附带、特殊、惩罚性或后果性损失，软件的开发者、贡献者、维护者及其关联方（以下统称"开发者"）**均不承担任何责任**，包括但不限于：

1. 因输入错误、录入错误、字段错位、模板错误或操作不当造成的信息错误；
2. 文件、病历、图像、报告或其他数据的丢失、损坏、被篡改、泄露或无法恢复；
3. 因软件缺陷、算法偏差或系统故障导致的误诊、漏诊、延误诊疗或错误报告；
4. 病理报告、医疗文书或诊疗结果引发的医疗纠纷、投诉、诉讼、索赔、行政处罚或刑事责任；
5. 违反医疗器械监管、病历管理、个人信息与数据安全、网络安全等法律法规所产生的后果；
6. 因第三方组件、模型、服务、网络环境或硬件所造成的任何损失；
7. 因业务中断、利润损失、声誉损害或其他经济损失。

**5. 使用者的责任**
使用者应自行负责：(a) 评估本软件是否适合其使用目的；(b) 遵守所在地适用的法律法规（包括医疗、数据保护和网络安全）；(c) 对数据进行加密、定期备份和脱敏处理；(d) 取得患者知情同意及必要的授权与伦理审批；(e) 对软件输出进行专业复核。**严禁**在未做完整合规评估和验证的情况下处理真实患者数据。

**6. 赔偿**
若因使用者使用或滥用本软件致使开发者遭受第三方索赔、行政调查或损失，使用者同意在法律允许的范围内对开发者进行赔偿并使其免受损害。

**7. 法律适用的限制**
本声明在适用法律允许的最大范围内有效。若本声明的任何条款依法被认定为无效或不可执行，不影响其他条款的效力。法律强制性规定不可排除的责任，不在免除范围内。

**8. 声明的变更与接受**
开发者有权随时更新本声明而无需另行通知。**继续使用本软件即表示您知晓并接受上述全部风险，并同意本声明的约束。** 如您不同意本声明的任何内容，请立即停止使用并删除本软件。

*本软件的使用同时受所附开源许可证约束；如本声明与许可证存在冲突，以对开发者责任限制更广泛者为准（在法律允许范围内）。*

### 🇬🇧 English Version

**1. Nature of the Project.** This is a fully open-source research, educational, and testing project. It is **NOT** a registered, cleared, certified, or approved medical device, diagnostic tool, or clinical information system. It has **NOT been adequately tested or validated in real clinical environments** and may contain defects, logic errors, security vulnerabilities, and unknown risks.

**2. Not for Clinical Use.** This software must **NOT** be used for clinical diagnosis, treatment decisions, pathology interpretation, issuing reports, or any activity affecting patient health or safety. All outputs are **for reference only** and do not replace the independent judgment of licensed professionals. Any use in a real medical setting is at the user's **sole risk and responsibility**.

**3. "AS IS", No Warranty.** The software is provided "AS IS" and "AS AVAILABLE", without warranties of any kind, express or implied, including merchantability, fitness for a particular purpose, accuracy, completeness, reliability, and non-infringement.

**4. Limitation of Liability.** To the maximum extent permitted by law, the developers, contributors, and maintainers shall **not be liable** for any direct, indirect, incidental, special, punitive, or consequential damages arising from the use of or inability to use this software, including without limitation: (a) input or data-entry errors; (b) loss, corruption, tampering, leakage, or unrecoverability of files, records, images, reports, or data; (c) misdiagnosis, missed diagnosis, delayed treatment, or erroneous reports caused by defects or algorithmic errors; (d) disputes, complaints, lawsuits, claims, penalties, or criminal liability arising from pathology reports or medical documents; (e) violations of medical-device, medical-records, privacy, data-security, or cybersecurity laws; (f) failures of third-party components, models, services, networks, or hardware; (g) business interruption or loss of profits or reputation.

**5. User Responsibilities.** Users are solely responsible for assessing suitability, complying with applicable laws (including healthcare, data-protection, and cybersecurity), encrypting, de-identifying, and backing up data, obtaining patient consent and ethical approvals, and professionally reviewing all outputs. Do **not** process real patient data without full compliance assessment and validation.

**6. Indemnification.** To the extent permitted by law, you agree to indemnify and hold harmless the developers from any third-party claims, investigations, or losses arising from your use or misuse of this software.

**7. Severability.** If any provision is held unenforceable, the remainder stays in effect. Liability that cannot be excluded under mandatory law is not excluded.

**8. Changes and Acceptance.** The developers may update this disclaimer at any time. **Continued use of this software constitutes your acknowledgment and acceptance of all the risks described above.** If you do not agree, stop using and delete the software immediately.

---

## 📜 开源协议

本项目采用 [MIT 许可证](LICENSE) 开源。欢迎自由使用、学习交流与二次开发。
