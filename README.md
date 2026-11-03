<div align="center">

# 🔬 Dermatology Pathology Reporting System
### 皮肤病理报告全流程系统

*面向临床皮肤科医师与病理技术员的现代化、轻量级、开箱即用型病理报告制作与全流程质控工作站*

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.8%2B-blue.svg?logo=python&logoColor=white" alt="Python Version" />
  <img src="https://img.shields.io/badge/Platform-Windows%2010%20%7C%2011-0078D6.svg?logo=windows&logoColor=white" alt="Platform" />
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

**皮肤病理报告系统** 是专为医院皮肤科临床取材送检、患者信息录入、显微镜观察摄片、病理医师描述及诊断、特殊免疫组化标记、上级医师质控审核、报告打印发放及纸质签收留档等全业务流程量身定制的高效工作站。

系统采用前后端解耦的 **本地原生轻量 WebUI 架构**，无需连接外部互联网，数据完全保留在本地工作目录中，兼具极致的响应速度与医疗数据私密安全性。

---

## ✨ 核心特性

- 📐 **多图自适应排版与“单页打印保护”**
  - 原生支持 **1 至 9 张** 显微病理照片智能阵列排版，图注与送检部位严格下沉对齐；
  - 严格锁定行距 15 磅，并根据图片数量动态适配镜下描述与病理诊断建议行数（单图 15/10行，双图 14/10行，三图 15/10行，4-6图 10/6行，7-9图 5/3行）；
  - 独创 **单页打印硬件保护机制**：根据不同图片规格智能计算行高补偿，确保第一页收尾于医生签名栏。打印机硬件强制仅输出法定第 1 页，第二页病史作为电子档案永久存档，彻底杜绝多页废纸浪费。
- 🛡️ **WebUI 启动合规风控与免责声明**
  - 每次打开软件启动 WebUI 时，界面自动弹出不可跳过的免责与风险告知窗口；
  - 提供【我已知晓并同意】与【退出】双选项，操作者知晓并同意后方可继续使用，点击退出时系统安全平稳释放后台常驻进程。
- 🔒 **严密的质控安全锁与审核闭环**
  - 勾选 `待审核`（疑难病例）或 `免疫组化`（特染标记）时，系统启动质控安全锁，**强制拦截制作台直接归档**；
  - 自动分发派送至专职质控队列，由审核医师复核签署后方可移入 `完成审核` 并放行打印。
- 🔢 **多院区与病房独立流水号管理**
  - 支持普通门诊（无前缀）、病房标本（`B`）、分院区标本（`L`）、疑难会诊（`H`）、科研特检（`T`）；
  - 各前缀分类享有完全独立的当日流水计数器，互不干扰、互不冲号。
- 🖼 **临床照片一键穿透调阅与诊断词库推荐**
  - 一键穿透检索患者大体皮损照片文件夹，即时调阅 Windows 资源管理器对比观察；
  - 内置丰富病理描述词库，智能识别显微镜图片命名中的疾病特征，置顶推荐标准镜下描述并支持一键填入。
- 🖨 **双列表物理隔离输出与签领单自动化**
  - 报告输出区采用“常规报告”与“完成审核特殊报告”物理隔离双列表，彻底杜绝未审报告误发；
  - 独立的签领单生成工作区，支持按时间段和多前缀灵活筛选，一键导出规范 Excel/Word 签名单并自动唤起打印。
- 🔍 **穿透式综合查询与“需要重新归档”保护**
  - 全要素多重组合检索与状态回退，支持穿透定位输出文件、照片及归档文件夹；
  - 针对已归档报告的二次修改实施 `需要重新归档` 保护，严防静默篡改历史已封存档案。

---

## 🏗 系统架构与技术栈

```
+--------------------------------------------------------------------------+
|                          本地 WebUI (客户端呈现层)                        |
|   HTML5 / CSS3 (现代响应式布局、护眼配色) + 原生 JavaScript (无外部CDN依赖)    |
+--------------------------------------------------------------------------+
                                     │ (HTTP REST API / 127.0.0.1:8765)
                                     ▼
+--------------------------------------------------------------------------+
|                        后端服务层 (Python 3.8+)                           |
|   内置标准库 ThreadingHTTPServer 高性能并发微服务 + RESTful API 路由分发    |
+--------------------------------------------------------------------------+
        │                        │                          │
        ▼                        ▼                          ▼
+---------------+    +------------------------+    +-----------------------+
|  数据持久层   |    |      文档排版引擎       |    |     系统级交互接口     |
| SQLite 3 本地 |    |  python-docx / OpenXML |    | Windows Shell API     |
| 关系型数据库  |    |  固定15磅行距精确计算   |    | COM Word 单页硬件打印 |
+---------------+    +------------------------+    +-----------------------+
```

---

## 📁 仓库目录结构

```plaintext
dermatology-system/
├── app/                         # 后端核心业务逻辑
│   ├── config.py                # 系统配置与初始化管理
│   ├── database.py              # SQLite 数据库模型与 CRUD
│   ├── doc_generator.py         # Word 报告生成与智能行数补偿排版
│   ├── doc_preview.py           # 报告高保真预览与打印接口
│   ├── path_manager.py          # 规范目录管理、文件生命周期流转
│   ├── server.py                # 本地 HTTP 服务与 REST API
│   ├── signing_sheet_generator.py # 报告领取签领单自动化生成
│   └── template_manager.py      # 词库与模板解析器
├── web/                         # 前端 Web 资源
│   ├── index.html               # 交互主界面
│   ├── css/style.css            # 医疗质感高对比度样式
│   └── js/app.js                # 前端交互与 API 通信
├── 模板/                        # 基础排版骨架 (1图、2图、3-9图标准模板)
├── 文字模板/                    # 专业病理描述词库 (普通病理 + 免疫组化)
├── 文档/                        # 系统详细设计说明、全景手册与 PDF 图文教程
├── 示例文件/                    # 标准脱敏交付成品报告范例与签名单示例
├── 工作目录/                    # 系统标准运行根目录 (已归档报告/未归档报告/特殊患者归档等)
├── scripts/                     # 运维辅助与构建脚本
│   ├── build_standalone_package.py # 一键打包绿色免安装发布版
│   └── generate_manual_pdf.py   # 编译图文版 PDF 操作手册
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
1. 从 Releases 中下载最新发布的 `皮肤病理报告系统-绿色免安装版.zip`；
2. 解压至本地任意磁盘（如 `D:\` 或 `E:\` 根目录）；
3. 双击根目录下的 `启动系统.bat`（或 `启动系统(无黑框后台版).vbs`），系统将自动启动微服务并在默认浏览器中打开 `http://127.0.0.1:8765`；
4. 首次使用时根据弹窗指引确认工作目录及医生姓名，即可立即开展日常病理报告录入与出具工作。

### 方式二：从源码运行（适合开发者与二次定制）

#### 1. 环境准备
- 操作系统：Windows 10 / 11 (64位)
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
控制台将输出启动信息，并在 1 秒后自动调起系统默认浏览器访问 `http://127.0.0.1:8765`。

---

## 📖 系统详细文档

系统各业务流程的使用方式、质控规范和排版规则已详细载入以下文档中，请查阅 `文档/`：

- **[系统逻辑设计与冲突梳理说明.md](文档/系统逻辑设计与冲突梳理说明.md)**：涵盖 7 大核心潜在业务冲突与防呆防覆盖机制的深度设计。
- **[皮肤病理报告系统_系统功能与架构全景手册.md](文档/皮肤病理报告系统_系统功能与架构全景手册.md)**：系统全功能规范、数据库字典、状态机流转与排版规则权威参考。
- **[皮肤病理报告系统_快速入门操作手册.pdf](文档/皮肤病理报告系统_快速入门操作手册.pdf)**：内置流程图与版面示意图的 4 页图文入门培训教程，适合规培生与进修医生即学即用。
- **[源码脱敏处理记录与审核日志.md](文档/源码脱敏处理记录与审核日志.md)**：开源版本全要素脱敏合规审核记录。

---

## 🛡 数据脱敏合规声明

本项目源码、文档及示例数据在公开发布前，已根据相关医疗信息安全规范执行了严格的**全要素脱敏处理**：
- **医疗机构**：真实医院及院区名称均已替换为泛化代称（“某三甲医院皮肤科”）；
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
