# 软件著作权申报材料准备通用工具包 (software-copyright-skills)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Version: V1.0-Final](https://img.shields.io/badge/Version-V1.0--Final-blue.svg)](https://github.com/shitanlylee0516/software-copyright-skills)

本项目是一个开源、通用的中国计算机软件著作权登记材料准备与排版辅助工具（Skill），当前为 **V1.0 最终版**。旨在帮助开发者、科研人员和技术团队标准化整理申报材料草案、排查常见敏感文件、审计代码物理行数，并生成符合规范格式的完整版 Word (`.docx`) 用户操作说明书。

开源主页：[https://github.com/shitanlylee0516/software-copyright-skills](https://github.com/shitanlylee0516/software-copyright-skills)

> [!IMPORTANT]
> **免责与官方地位声明**：
> 1. 本项目为开源技术辅助工具，非官方发布，**不代表中国版权保护中心（CPCC）或任何官方机构**，不作登记成功保证。
> 2. 本工具包依据 MIT 许可证免费开源，仅代表工具本身无授权许可费，不代表官方软件著作权登记规费的免除或保证；当前官方登记是否收费或具体收费标准须以官方最新规定为准。
> 3. 本工具仅拦截高风险文件名与敏感路径组件，**无法对源代码全文内容进行深层商业机密或硬编码凭证的完全检测**，申报人仍须进行严格的手工机密审查。
> 4. 所有涉及著作权主体、统一社会信用代码、开发完成时间及首次发表时间等法定事实，若未获申报人充分授权与事实核实，一律置为空（`null`）或标记为“待确认”，严禁伪造事实。

---

## 目录结构

```text
software-copyright-skills/
├── LICENSE                                     # MIT 开源许可证
├── README.md                                    # 项目主说明文档
├── .gitignore                                  # Git 忽略配置
├── requirements-optional.txt                   # 可选扩展依赖（排版工具）
├── skills/
│   └── software-copyright-materials/
│       ├── SKILL.md                            # Skill 主定义与指引
│       ├── agents/
│       │   └── openai.yaml                     # OpenAI Agent 配置文件
│       ├── references/
│       │   ├── requirements.md                 # 著作权申报材料标准规范指南
│       │   ├── manuals.md                      # 操作手册排版与 Word 替换工作流
│       │   └── source-review.md                # 源码审查准则与物理行数审计
│       ├── scripts/
│       │   ├── collect_source.py               # 源码审查采集与清单清点工具
│       │   └── build_manual.py                 # 操作手册 Word (.docx) 生成工具
│       └── examples/
│           ├── application.example.json        # 申报元数据虚拟示例
│           └── screenshots.example.json        # 截图占位槽位配置示例
└── tests/                                      # 自动化测试用例（独立验收用）
```

---

## 安装与配置

### 1. 安装 Skill

要将本 Skill 导入您的 Codex 或智能助手环境中，只需将 `skills/software-copyright-materials` 文件夹复制到相应的技能存储目录中（请确保不覆盖已有同名自定义技能）：

```bash
# 复制示例（根据实际技能目录调整目标路径）
cp -r skills/software-copyright-materials /path/to/your/agent/skills/
```

### 2. 环境与依赖要求

- **Python 3.10+**
- 基础源码审查工具 `collect_source.py` 仅使用 Python 标准库，**无需安装任何额外依赖**。
- Word 操作手册生成工具 `build_manual.py` 需要安装可选排版依赖：

```bash
pip install -r requirements-optional.txt
```

### 3. 测试运行

在安装上述可选依赖后，可在项目根目录下运行独立自动化测试套件：

```bash
python -m unittest discover -s tests
```

---

## 核心工具使用说明

### 一、源代码自审与行数清点 (`collect_source.py`)

用于在正式整理前自审源程序文件、拦截 `.env` / 私钥 / 证书等敏感文件、清点物理行数并生成自审包。

> [!NOTE]
> 该工具生成的产物仅作为内部自审和行数统计用途，**并非直接用于版权局提交的最终法定鉴存文档**。

```bash
python skills/software-copyright-materials/scripts/collect_source.py \
  --root /path/to/project_root \
  --manifest /path/to/manifest.json \
  --out /path/to/review_dir
```

- `--root`：源代码根目录。
- `--manifest`：JSON 清单文件，必须仅包含唯一键 `files`，格式形如 `{"files": ["src/main.py", "src/util.py"]}`。严禁包含冒号（`:`）、绝对路径、路径遍历（`..`）或通配符。
- `--out`：目标输出目录。为保护现场，如果该目录已存在，工具将拒绝覆盖并安全退出。

### 二、用户操作说明书 Word 完整版生成 (`build_manual.py`)

基于 Markdown 文本与元数据配置，自动生成符合标准排版规范的完整版 Word (`.docx`) 软件说明书。

```bash
python skills/software-copyright-materials/scripts/build_manual.py \
  --input /path/to/manual.md \
  --metadata /path/to/application.json \
  --screenshots /path/to/screenshots.json \
  --output /path/to/manual.docx
```

#### 推荐排版主工作流（完整版交付）：
1. **生成完整版说明书**：使用上述命令生成带有标准页边距、宋体版式、黑色页眉（1x2 无框表格严格靠右对齐）、动态页码（`PAGE` / `NUMPAGES`）及截图占位槽位的完整版 `.docx` 文档。
2. **替换截图**：在 Microsoft Word 或 WPS 中打开文档，右键单击附录中的占位图片，选择**「更改图片」**（Change Picture），替换为真实的软件全屏或主窗口运行截图。请保持截图原始宽高比，并逐图检查裁切与排版适配，避免直接删除占位图重新插入。
3. **逐页检查与导出 PDF**：检查文字描述与界面一致性后，人工逐页核验，导出为高质量 PDF 文件供最终审查与打印。

---

## 局限性与使用边界

1. **无本地渲染引擎环境限制**：自动化脚本主要负责构建符合 OpenXML 规范的 Word 结构和嵌入占位图。在缺少 Microsoft Office 渲染引擎的环境中，未执行视觉渲染并不代表最终排版验收，必须通过 Word 打开并导出 PDF 进行人工逐页检查。
2. **独立 PDF 差异**：直接由 Markdown 渲染生成的独立 PDF 仅为可选辅助方案，其与 Word 排版引擎在分页断点、边距折算上存在差异，不能作为 Word 最终排版样式的等价证明。

---

## 许可证

本项目依据 [MIT 许可证](LICENSE) 开源。版权所有 (c) 2026 shitanlylee0516。
