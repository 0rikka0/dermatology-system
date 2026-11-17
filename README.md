<div align="center">

# 🔬 Dermatology Pathology Reporting System
### 皮肤病理报告全流程系统

*面向临床皮肤科医师与病理技术员的现代化、轻量级、开箱即用型病理报告制作与全流程质控工作站*

<p align="center">
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
  <a href="#-数据脱敏合规声明">脱敏声明</a>
</p>

</div>

---

## 📖 系统介绍

**皮肤病理报告系统** 是专为医院皮肤科临床取材送检、患者信息录入、显微镜观察摄片、病理医师描述及诊断、上级医师质控审核（待审核/重切）、独立免疫组化第二份报告出具、Word COM 单页打印发放及纸质签收留档等全业务流程量身定制的高效工作站。

系统采用前后端解耦的 **本地原生轻量 WebUI 架构（8 大工作区）**，无需连接外部互联网，全部数据严格保留在本地工作目录中，兼具极致响应速度与医疗数据私密安全性。

---

## ✨ 核心特性

- 📐 **多图自适应排版与“单页打印保护”**
  - 原生支持 **1 至 9 张** 显微病理照片智能阵列排版，图注与送检部位严格下沉对齐；
  - 严格锁定行距 15 磅，并根据图片数量动态适配镜下描述与病理诊断建议行数（单图 15/10 行，双图 14/10 行，三图 15/10 行，4-6 图 10/6 行，7-9 图 5/3 行）；
  - 独创 **单页打印硬件保护机制**：第一页精确收尾于医生签名栏，后台调用 Word/WPS COM 强制仅打印法定第 1 页，第二页病史永久留存电子档案，杜绝多页废纸浪费。
- 🧪 **独立免疫组化（IHC）双报告工作区**
  - 仅标注 `需免疫组化` 的患者原初诊报告按常规流程直接制作、打印与一键归档；
  - 在独立的【免疫组化】模块基于已归档原报告创建第二份副本（复制原归档报告与病理照片至 `特殊患者归档/免疫组化/`，**不复制临床照片**），提供复制后病理照片目录打开与路径一键复制功能；
  - 支持完全独立地编辑、保存（不动 `patients` 主表及原报告文件）并归档至 `已归档报告/<原文件名>_免疫组化报告.docx`，同一患者两份报告并存互不覆盖、独立输出与撤销。
- ⚖️ **待审核与重切等价协同、双草稿同步与防误改保护**
  - 【报告审核】专责处理 `待审核` 与 `重切` 病例（二者逻辑完全等价；兼有免疫组化时双标各自执行、原报告按待审流程归档）；
  - 触发特殊分流时**不复制临床照片**（仅复制病理照片目录），审核区【临床照片查询】仍打开 `未归档报告/临床照片/`，而【病理照片目录】打开与复制路径指向 `特殊患者归档/待审核(或重切)/.../病理照片/`；
  - 审核保存与制作保存**双向完全同步** `未归档报告/病理报告输出/` 与 `特殊患者归档/` 两份草稿及新增照片，且**不提前修改 `已归档报告/`**；修改已归档病例自动置 `archived=2`（需重新归档），审核完成或修订后统一前往【报告制作】一键归档入库。
- 🛡️ **WebUI 启动合规风控与文件占锁防护**
  - 每次打开 WebUI 自动弹出不可跳过的免责声明告知窗，退出时平稳释放本地进程；
  - 写入 Word/Excel 前检测独占锁，若被 Word/WPS 打开则返回清晰中文提示；全流程删除均通过 Windows 回收站软删除；SQLite 启用 WAL 模式并发防锁。
- 🔢 **多院区与病房独立流水号管理**
  - 支持普通门诊（无前缀）、病房标本（`B`）、分院区标本（`L`）、疑难会诊（`H`）、科研特检（`T`）；`H`/`T` 前缀支持自定义送检单位与科室，年龄支持 `岁/月/日` 及 `自由输入` 兼容儿童；
  - 各前缀分类享有完全独立的当日流水计数器，互不干扰、互不冲号。
- 🖼 **临床照片穿透调阅与 249 篇智能诊断词库**
  - 录入、制作、审核与组化工作区均支持一键检索 `未归档报告/临床照片/`；
  - 内置 249 篇专科文字模板（普通 191 篇 + 免疫组化 58 篇），自动识别显微镜图片命名特征（单图自动去序号、多图保留序号）并智能推荐填入。
- 🖨 **三列表安全隔离输出与签领单自动化**
  - 【报告输出】采用 **常规已归档**（含纯组化患者原报告）、**特殊审核完成** 与 **免疫组化已归档** 三列表物理隔离输出，支持单独或批量调用 Word COM 仅打第 1 页及独立撤销归档/审核；
  - 【报告领取签名单】支持按送检/报告日期与多前缀筛选，生成 Excel 与无空行 Word 签名单，Word 中醒目标注参考**手术日期**（前一工作日）并支持 Word 直接打印。

---

## 🏗 系统架构与技术栈

```
+--------------------------------------------------------------------------+
|                          本地 WebUI (客户端呈现层)                        |
|   HTML5 / CSS3 (8 模块响应式布局) + 原生 JavaScript (无外部 CDN 依赖)      |
+--------------------------------------------------------------------------+
                                     │ (HTTP REST API / 127.0.0.1:8765)
                                     ▼
+--------------------------------------------------------------------------+
|                        后端服务层 (Python 3.8+)                           |
|   内置标准库 ThreadingHTTPServer 高性能并发微服务 + RESTful API 路由分发    |
+--------------------------------------------------------------------------+
        │                        │                          │
        ▼                        ▼                          ▼
+--------------------+  +------------------------+  +-----------------------+
|   数据持久层       |  |      文档排版引擎       |  |     系统级交互接口     |
| SQLite 3 (WAL)     |  |  python-docx / OpenXML |  | Windows Shell / 回收站|
| patients +         |  |  固定15磅行距精确计算   |  | COM Word 单页硬件打印 |
| ihc_reports 双表   |  |                        |  |                       |
+--------------------+  +------------------------+  +-----------------------+
```

---

## 📁 仓库目录结构

```plaintext
dermatology-system/
├── app/                         # 后端核心业务逻辑
│   ├── config.py                # 系统配置与初始化管理
│   ├── database.py              # SQLite 双表模型 (patients + ihc_reports) 与 CRUD
│   ├── doc_generator.py         # Word 报告生成与智能行数补偿排版
│   ├── doc_preview.py           # A4 高保真预览引擎（含普通/免疫组化双模式）
│   ├── path_manager.py          # 目录管理、双草稿同步、独立 IHC 归档与占锁检测
│   ├── server.py                # 本地 HTTP 服务与 REST API（含 /api/ihc/* 与 /api/reviews/*/photos）
│   ├── signing_sheet_generator.py # Excel + 无空行 Word 签领单生成
│   └── template_manager.py      # 249 篇病理词库解析与诊断自动合成
├── web/                         # 前端原生 Web 资源
│   ├── index.html               # 八模块交互主界面
│   ├── css/style.css            # 医疗质感高对比度三栏输出与路径栏样式
│   └── js/app.js                # 全业务模块交互与 API 通信
├── 模板/                        # 标准排版骨架 (1图、2图、3-9图模板与 placeholder.jpeg)
├── 文字模板/                    # 专业病理描述词库 (普通模板.docx + 免疫组化模板.docx)
├── 文档/                        # 系统全景架构、逻辑冲突梳理、SOP 图文操作手册 PDF 与脱敏日志
├── 示例文件/                    # 脱敏虚拟患者成品示例报告与签名单示例
├── 工作目录/                    # 标准运行目录占位骨架 (已归档报告/未归档报告/特殊患者归档等)
├── scripts/                     # 运维辅助与构建脚本
│   ├── build_standalone_package.py # 一键打包绿色免安装发布版
│   └── generate_manual_pdf.py   # 编译图文版 SOP 操作手册 PDF
├── tests/                       # 自动化功能测试套件
├── main.py                      # 系统主程序入口
├── 启动系统.bat                 # Windows 一键启动脚本
├── requirements.txt             # 源码运行所需第三方依赖
├── .gitignore                   # Git 忽略规则
└── LICENSE                      # 开源许可协议 (MIT)
```

---

## 🚀 快速开始

### 方式一：独立便携版（推荐，适合临床科室直接使用）
1. 从 Releases 下载 `皮肤病理报告系统-绿色免安装版.zip`（或 Windows 7 专版 `win7.zip`）；
2. 解压至本地任意磁盘（如 `D:\` 或 `E:\` 根目录）；
3. 双击 `启动系统.bat`（或无黑框后台版 `.vbs`），系统自动启动本地服务并在默认浏览器中打开 `http://127.0.0.1:8765`；
4. 首次使用时按向导确认工作目录及医生姓名，即可开始作业（详细步骤请直接查阅 `文档/皮肤病理报告系统_快速入门操作手册.pdf` 内附 SOP）。

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

#### 3. 启动应用
```bash
python main.py
```

---

## 📖 系统详细文档

系统各业务流程的使用方式、标准操作规程（SOP）、质控规范与排版规则已详细载入以下文档中，请查阅 `文档/`：

- **[皮肤病理报告系统_快速入门操作手册.pdf](文档/皮肤病理报告系统_快速入门操作手册.pdf)**：含工作流图解与 **SOP-01 至 SOP-06 标准操作规程** 的 5 页图文手册，适合医师与技术员即学即用。
- **[皮肤病理报告系统_系统功能与架构全景手册.md](文档/皮肤病理报告系统_系统功能与架构全景手册.md)**：v1.4.0 全功能规范、双表数据字典、三轨生命周期与排版规则权威参考。
- **[系统逻辑设计与冲突梳理说明.md](文档/系统逻辑设计与冲突梳理说明.md)**：涵盖 10 大核心潜在业务冲突与防呆防覆盖机制的深度设计。
- **[源码脱敏处理记录与审核日志.md](文档/源码脱敏处理记录与审核日志.md)**：开源版本全要素脱敏合规审核记录。

---

## 🛡 数据脱敏合规声明

本项目源码、文档及示例数据在公开发布前，已根据相关医疗信息安全规范执行了严格的**全要素脱敏处理**：
- **医疗机构与分院**：真实医院与特定分院区名称均已替换为泛化代称（“某三甲医院皮肤科”、“分院区”）；
- **医护人员**：真实医生姓名均已替换为标准代称（“张医生”、“李主任”、“王医生”）；
- **患者隐私**：所有示例数据中的患者姓名、门诊号、病理号及送检部位均为系统合成模拟数据；
- **元数据擦除**：Office 文档的核心元数据（创建者、最后修改者、本地绝对路径）均已彻底清除重写。

---

## ⚠️ 免责声明与风险告知 / Disclaimer and Risk Notice

> **请在使用本软件前仔细阅读本声明。您下载、安装、运行或以任何方式继续使用本软件，即视为您已阅读、理解并完全接受本声明的全部内容。**
> **Please read carefully before using this software. By downloading, installing, running, or continuing to use it, you acknowledge that you have read, understood, and agreed to this disclaimer in full.**

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
