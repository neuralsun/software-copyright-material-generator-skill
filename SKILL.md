---
name: generate-cn-software-copyright-kit
description: "Analyze a software source tree and generate or audit the three Chinese software-copyright materials: an evidence-backed software design and user manual with screenshots and diagrams, an application-form draft with verified technical facts and uninvented legal fields, and a general-deposit source-code document with reproducible first/last-page selection and real Office pagination. Use for 软件著作权、软著申请、说明书、申请表、源代码交存、前后各30页、60页代码、补正检查、代码雷同检查, or cross-material consistency review from an existing codebase."
---

# 生成中国软件著作权三件套

把源码视为事实来源，把用户确认的名称、版本、日期和权利信息视为法律事实来源。生成可复核的三份 DOCX：软件设计与使用说明书、软件著作权登记申请表底稿、源程序一般交存文档。不得把“文档已生成”表述为“必然通过登记”。

## 开始前

完整阅读以下四份规则后再行动：

- `references/config-schema.md`：统一配置、中间文件和各生成器输入结构。
- `references/evidence-and-writing-rules.md`：功能证据等级、运行验证、截图和写作边界。
- `references/material-details.md`：三份材料的版式、字段、选码和内容细节；需要定位申请表模板单元格时读取 `references/application-template-map.json`。
- `references/qa-gates.md`：阻断提交的终检条件和报告要求。

同时核对任务当日接收平台、办理地区和当前官方表单要求。只采用可追溯的官方规则；无法核对时继续生成“待复核底稿”，但把 `official_rules_verified` 设为 `false`，不得宣称可直接提交。平台在线表单或当期官方模板优先于本 Skill 的通用申请表模板。

下面命令中的 `scripts/`、`assets/` 和 `references/` 都相对于本 Skill 目录解析。若当前工作目录是用户项目，使用这些资源的绝对路径，不要误用项目中同名目录。

运行前确认 Python 3.10+、`python-docx`、Pillow、lxml 和 PyMuPDF 可用；版本范围见 `requirements.txt`。真实分页另需 Windows 上可自动化的 WPS Writer 或 Microsoft Word。缺少排版引擎时可以生成草稿，但不能通过最终分页门禁。

## 建立工作区与唯一事实源

保持项目源码只读。把过程文件放进独立工作目录，把最终目录限制为三份 DOCX。建议结构：

```text
copyright-work/
  kit_config.json
  project_facts.json
  source_manifest.json
  feature_evidence.json
  figures_manifest.json
  diagrams.json
  manual_content.json
  deposit_config.json
  deposit_manifest.json
  qa_report.json
  qa_report.md
  figures/
  previews/
final/
  <软件全称>_软件设计与使用说明书.docx
  <软件全称>_软件著作权登记申请表.docx
  <软件全称>-代码(一般交存版).docx
```

先创建 `kit_config.json`，以后所有名称、版本、完成日期、程序量、技术环境和权利字段都从这里派生。法律主体、证件、地址、联系人、权利取得方式、完成日期和发表情况只能由用户或其已有正式材料明确提供；缺失时保留空白并阻断“可提交”状态，不得猜写。

## 扫描并冻结源码

运行：

```powershell
python scripts/scan_project.py --project <源码根目录> --config <kit_config.json> --output <project_facts.json> --manifest <source_manifest.json>
```

先检查排除目录、第一方源码范围和语义顺序，再接受清单。不得纳入依赖、构建产物、压缩文件、缓存、生成代码、素材、数据库、上传文件或旧软著材料。若显式源码缺失、文件无法严格解码、存在替换字符、硬编码凭据、旧软件名或版本冲突，先处理或让用户确认，不能越过错误继续冻结。

冻结后不得修改源码。任何文件哈希变化都必须重新扫描、重建三份材料并重做分页。程序量统一采用冻结清单中的非空物理行数，不使用 Word 段落数、压缩包大小或估算值。

## 建立功能—代码证据矩阵

逐模块阅读入口、界面、路由/API、业务服务、持久化或外部调用，编写 `feature_evidence.json`。按证据等级 0—4 判定功能；只有 3 级及以上才能作为已实现用户功能，只有 4 级才能作为说明书核心截图功能。模型、表结构、依赖、按钮、模拟数据、README、未调用函数或构建产物均不能单独证明功能完成。

在隔离环境中构建和运行系统，执行主用户路径与异常路径的冒烟测试。不得连接生产环境、发送真实消息、创建真实订单或更改真实数据。无法安全运行时写明限制，把相应功能降级，不得伪造运行结果。

## 生成并审查图片

优先截取实际运行界面。每张截图都要能对应功能 ID，并人工或 OCR 检查软件名称、版本、个人信息、测试账号、密钥、磁盘路径、错误堆栈、浏览器警告和拼接痕迹。允许裁掉无关窗口边框或做隐私遮盖，但不得拼接成不存在的界面，不得用设计稿冒充运行截图。

缺少真实截图时可在草稿中使用“后期补图”占位，但最终 `submission_ready` 必须为 `false`。架构图和流程图只能使用已证实的节点、调用和数据流，配置 `diagrams.json` 后运行：

```powershell
python scripts/render_diagrams.py --input <diagrams.json> --output-dir <figures目录>
```

## 生成说明书

从证据矩阵和截图清单编写 `manual_content.json`。一级结构采用“软件设计说明”和“软件使用说明”，正文尽量使用连贯、信息密度高的长段落，避免口号、同义反复、机械枚举和套话。每个操作步骤必须与真实界面和实际数据流一致；未实现的预约、支付、审核、删除、导出、消息推送等能力不得因存在名称相近的表或按钮而写入。

运行：

```powershell
python scripts/build_manual.py --input <manual_content.json> --output <说明书.docx>
```

草稿阶段才可设置 `document.allow_missing_images=true`。正式生成必须关闭它，并确保所有核心功能有证据、所有图题准确、图片清晰、目录和页码可刷新。

## 生成申请表底稿

把冻结程序量和经证据约束的主要功能技术说明写回 `kit_config.json`。功能技术说明去除空白后控制在 500—1000 字；不得加入说明书未证明的功能。运行：

```powershell
python scripts/build_application_form.py --config <kit_config.json> --template assets/application-form-template.docx --output <申请表.docx> --report <申请表报告.json> --require-ready
```

只有著作权人、申请人和代理人的身份/联系字段不全时，去掉 `--require-ready` 才能生成相应位置留白且 `ready=false` 的底稿。软件全称、版本、完成日期、原创性质、开发方式、权利范围、环境、语言、程序量和 500—1000 字功能技术说明属于生成前置事实；任一缺失时生成器会停止，应向用户补问，不能猜写。若用户暂时无法确认，只能在工作目录提供未填写的 `assets/application-form-template.docx` 并明确它不是已生成申请表，不得改名后混入最终三件套目录。本模板用于准备底稿；当期接收平台要求在线生成或使用新版表单时，必须把已核实内容转填到官方载体，不得改变官方表格结构。

## 生成源程序一般交存文档

从冻结清单派生 `deposit_config.json`，其中以 `source_manifest` 指向清单，并复用相同的软件名称和版本。先干跑并检查风险，再生成：

```powershell
python scripts/build_code_deposit.py --manifest <deposit_config.json> --dry-run --report <代码审计.json> --provenance <deposit_manifest.json>
python scripts/build_code_deposit.py --manifest <deposit_config.json> --output <代码.docx> --report <代码审计.json> --provenance <deposit_manifest.json>
```

源码不足 3000 个非空行时必须全量交存，绝不复制凑页；等于 3000 行时全量交存；超过 3000 行时按默认 50 行/页取前 1500 行和后 1500 行。保持同一冻结清单的连续程序顺序，不能为了避开雷同而删除合法业务代码。登录 HTML/CSS 等通用表现层可以作为相似度风险候选人工判断，但认证、授权和会话业务代码不能仅因常见而删去。

按需与用户提供的旧代码 DOCX、源码目录或旧 manifest 比较：

```powershell
python scripts/audit_similarity.py --manifest <deposit_config.json> --scope deposit --reference <旧材料或源码> --output <相似度报告.json>
```

本地相似度报告只说明已提供语料之间的连续块情况，不能代表登记机构未知库的比对结果。

## 用 Office 锁定真实分页

三个生成器写入的是动态域和显式分页结构。必须使用 Word 或 WPS 重新分页、更新正文/页眉/页脚域并导出 PDF：

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
& scripts/refresh_office.ps1 -Paths @("<说明书.docx>", "<申请表.docx>", "<代码.docx>") -PdfDir "<previews目录>" -Office Auto
```

记录脚本返回的每份真实页数，再修正 DOCX 页数属性和 `NUMPAGES` 缓存：

```powershell
python scripts/finalize_docx.py --path <说明书.docx> --pages <实测页数> --path <申请表.docx> --pages <实测页数> --path <代码.docx> --pages <实测页数> --report <分页报告.json>
```

逐页检查刚导出的 PDF，确认长代码行没有造成额外页、图表未截断、标题未孤立、目录正确且页码连续。若因此修改文档，重新执行 Office 刷新、导出和实测；最后一次 Office 保存之后再运行 `finalize_docx.py`，并且不要再用 Office 保存成品，以免 WPS 重写错误的 `NUMPAGES` 缓存。Office 自动化不可用时只能标记“分页未验证”，不得声称三件套已经可提交。

## 执行包级终检

运行统一审计：

```powershell
python scripts/audit_materials.py --config <kit_config.json> --facts <project_facts.json> --manifest <source_manifest.json> --deposit-manifest <deposit_manifest.json> --evidence <feature_evidence.json> --manual-content <manual_content.json> --manual <说明书.docx> --application <申请表.docx> --code <代码.docx> --pdf-dir <previews目录> --report-json <qa_report.json> --report-md <qa_report.md>
```

读取报告而不是只看退出码。阻断项必须归零；警告必须逐项分类并留下人工复核结论。再人工检查 DOCX 修订、批注、隐藏文字、外链、旧名称、旧元数据、敏感信息、截图清晰度、正式签章字段和当期官方表单。最终目录只放三份名称明确的 DOCX，过程 JSON、PDF 预览和报告留在工作目录。

仅当 `qa_report.json` 为 `submission_ready=true`，用户确认全部法律事实、截图和签章安排，且当期官方规则已核对时，才能表述为“可进入人工提交复核”。
