---
name: software-copyright-materials
description: 用于协助准备与初审中国计算机软件著作权登记材料（源代码自审抽查、用户操作手册 Word 排版模板及元数据清单编排），不提供官方注册或法律保证。
---

# 软件著作权申报材料准备 Skill

本 Skill 旨在协助开发者与企业合规整理中国计算机软件著作权登记申报所需的初审材料，包括源程序鉴存前的自审清单生成、物理行数统计、敏感凭据排查，以及用户操作手册的规范 Word (.docx) 模板编排。

> [!IMPORTANT]
> **法律与官方地位声明**：
> 1. 本 Skill 及相关脚本仅作为内部材料草拟、排版和格式自审的技术辅助工具，**不代表任何官方版权登记机构（如中国版权保护中心 CPCC）**，亦不构成任何法律意见、注册成功保证或法定义务免除。
> 2. 所有涉及权利主体名称、证件信息、开发完成日期、首次发表日期等法定事实，若未获申报人充分授权与事实核验，必须严格保持为空（`null`）或标记为「待确认」。
> 3. 严禁在本流程中上传、泄露或提交任何未经许可的私有代码、密钥证书、生产环境账号或商业秘密。本工具仅拦截高风险文件名与路径，无法全面扫描代码内容内部的机密凭据，仍须申报人进行严格的人工敏感信息排查。

## 核心功能与工作流

### 1. 源代码自审与清单审计 (`scripts/collect_source.py`)

提供确定性源程序文件清单抽检、编码验证（UTF-8/无 NUL）、敏感配置拦截、物理行数统计及审查合规包生成。

- 详细规范请参阅 [source-review.md](references/source-review.md)。
- **注意**：工具生成的产物仅供申报前内部自审与行数统计，**非最终法定源程序鉴存提交件**。

使用示例（在 Skill 所在目录下运行或指定相对路径）：
```bash
python scripts/collect_source.py \
  --root /path/to/project \
  --manifest /path/to/manifest.json \
  --out /path/to/review_output
```

### 2. 用户操作手册排版生成 (`scripts/build_manual.py`)

根据 Markdown 结构化内容与元数据，生成满足标准 A4、宋体排版、页眉页脚（Word PAGE/NUMPAGES 域）、明晰表格以及预置真实图片占位槽位的 `.docx` 母版。

- 推荐工作流请参阅 [manuals.md](references/manuals.md)。
- **操作要点**：在 Word 中打开生成的母版，右键单击占位图并选择**「更改图片」**（Change Picture），替换为实际软件运行界面全屏或主窗口截图。替换时务必保持原始宽高比，并逐图检查裁切与排版适配，避免直接删除占位图。
- 整体申请材料规范请参阅 [requirements.md](references/requirements.md)。

使用示例（在 Skill 所在目录下运行或指定相对路径）：
```bash
python scripts/build_manual.py \
  --input /path/to/manual.md \
  --metadata /path/to/meta.json \
  --screenshots /path/to/shots.json \
  --output /path/to/manual.docx
```

## 依赖说明

基础代码收集工具使用 Python 3.10+ 标准库，无需额外安装依赖；Word 操作手册生成工具依赖以下可选排版库，在独立安装本 Skill 时可直接安装：
- `python-docx==1.2.0`
- `markdown-it-py==3.0.0`
- `Pillow>=10,<13`

安装命令：
```bash
pip install python-docx==1.2.0 markdown-it-py==3.0.0 "Pillow>=10,<13"
```
