# Generate CN Software Copyright Kit

从已有系统源码生成并审计中国软件著作权登记常用的三类材料：

1. 软件设计与使用说明书；
2. 软件著作权登记申请表底稿；
3. 源程序一般交存文档。

本项目既可以作为 Codex Skill 使用，也可以直接调用其中的 Python 和 PowerShell 脚本。它强调“源码证据、法律事实、真实分页、跨材料一致性”四条主线，适合从 Web、桌面、CLI、移动端、插件或服务端项目中整理软著材料。

> [!IMPORTANT]
> 本项目生成的是提交前复核材料和申请表底稿，不提供法律意见，也不保证登记通过。正式办理前必须核对任务当日接收平台、办理地区和官方表单要求。平台在线表单、补正通知和窗口要求始终优先于本仓库的通用规则与模板。

## 为什么需要这个项目

软著材料的问题通常不在 Word 能否生成，而在三个文件是否来自同一事实源。常见风险包括：说明书写入代码并未实现的功能，申请表程序量与源码不一致，截图仍显示旧软件名，少于 3000 行的源码被复制凑成 60 页，Word 缓存页数与真实渲染页数不同，以及申请人、完成日期或权利范围被擅自补写。

本项目把这些风险转化为可复现的工作流和阻断门禁：先冻结第一方源码，再建立功能—代码证据矩阵，随后生成三份 DOCX，最后通过 Word/WPS、PDF 和 DOCX 包级检查给出 `submission_ready` 结论。

## 主要能力

| 能力 | 说明 |
|---|---|
| 第一方源码扫描 | 排除依赖、构建产物、缓存、媒体、旧材料和数据库等非交存内容 |
| 可复现源码快照 | 记录文件顺序、编码、SHA-256、物理行数、非空行数和整包哈希 |
| 名称与版本检查 | 检查旧软件名、多个项目版本、带类型注解的版本常量和跨材料冲突 |
| 乱码与敏感值检查 | 识别 Unicode 替换字符、常见中文乱码、私钥、默认密码、邀请码和令牌候选 |
| 功能证据分级 | 以 0—4 级区分设计描述、数据结构、后端能力、完整调用链和运行截图 |
| 说明书生成 | 生成 A4 封面、动态目录、设计章、使用章、图片、表格和动态页码 |
| 架构图与流程图 | 根据显式 JSON 节点和边生成中文 PNG，不凭空推断系统能力 |
| 申请表底稿 | 基于 37 行参考模板填入已确认事实，缺失身份字段时留白并报告 |
| 一般交存代码 | 小于 3000 行全量交存；超过 3000 行取冻结顺序的前后各 1500 行 |
| 代码相似度审计 | 检查内部重复，并可与旧 DOCX、源码目录或旧 manifest 比较连续代码块 |
| Office 真实分页 | 使用 WPS Writer 或 Microsoft Word 更新域、重新分页并导出 PDF |
| 三件套统一终检 | 检查名称、版本、日期、程序量、证据、元数据、页数、PDF 和代码 provenance |

## 设计原则

- **源码是功能事实来源。** 模型、表结构、依赖、按钮、README、mock 数据和未调用函数不能单独证明功能完成。
- **用户确认是法律事实来源。** 著作权人、证件、地址、完成日期、开发方式、权利范围和发表情况不得推断。
- **一个配置驱动三份材料。** 软件名称、版本、日期、技术环境和程序量不能分别维护。
- **真实排版引擎决定页数。** `python-docx` 只能生成结构，最终页数必须由 Word/WPS 和 PDF 验证。
- **草稿与可提交状态分离。** 文件成功生成不等于可以提交，阻断项未清零时必须保持 `submission_ready=false`。

## 仓库结构

```text
generate-cn-software-copyright-kit/
├─ SKILL.md
├─ README.md
├─ requirements.txt
├─ agents/
│  └─ openai.yaml
├─ assets/
│  └─ application-form-template.docx
├─ references/
│  ├─ application-template-map.json
│  ├─ config-schema.md
│  ├─ evidence-and-writing-rules.md
│  ├─ material-details.md
│  └─ qa-gates.md
└─ scripts/
   ├─ scan_project.py
   ├─ render_diagrams.py
   ├─ build_manual.py
   ├─ build_application_form.py
   ├─ build_code_deposit.py
   ├─ audit_similarity.py
   ├─ refresh_office.ps1
   ├─ finalize_docx.py
   └─ audit_materials.py
```

详细配置和材料规则请直接阅读：

- [配置与中间文件规范](references/config-schema.md)
- [功能证据、运行验证与写作规则](references/evidence-and-writing-rules.md)
- [三份材料的内容与版式细节](references/material-details.md)
- [三件套终检门禁](references/qa-gates.md)

## 环境要求

基础环境：

- Python 3.10 或更高版本；
- `python-docx`、Pillow、lxml、PyMuPDF；
- 能够读取待处理项目源码的本地环境。

真实分页还需要：

- Windows；
- WPS Writer 或 Microsoft Word，且支持 COM 自动化。

说明书、申请表、代码 DOCX 和图片可在没有 Office 的环境中生成，但此时只能作为草稿，不能通过最终分页门禁。

## 安装

### 作为普通仓库使用

```bash
git clone <your-repository-url>
cd generate-cn-software-copyright-kit
python -m venv .venv
```

Windows PowerShell：

```powershell
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Linux/macOS：

```bash
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Linux/macOS 可以完成源码扫描、图片生成和 DOCX 初稿生成，但 `refresh_office.ps1` 仅适用于 Windows。若没有等效的真实排版验证，不要把结果标为可提交。

### 作为 Codex Skill 使用

把整个仓库目录复制到 Codex 的 Skills 目录，目录名保持为：

```text
generate-cn-software-copyright-kit
```

也可以在对话中直接提供本仓库路径。调用示例：

```text
使用 $generate-cn-software-copyright-kit 分析当前系统源码，生成软著说明书、申请表底稿和一般交存代码，并完成跨材料一致性检查。输出到指定目录，不要猜写申请人和权利信息。
```

## 快速开始

推荐让 Codex 按 [SKILL.md](SKILL.md) 执行完整流程。以下命令用于理解和手动调试；正式任务不要跳过证据、Office 分页和统一终检。

### 1. 准备工作目录

建议把过程文件和最终文件分开：

```text
copyright-work/
├─ kit_config.json
├─ project_facts.json
├─ source_manifest.json
├─ feature_evidence.json
├─ figures_manifest.json
├─ diagrams.json
├─ manual_content.json
├─ deposit_config.json
├─ deposit_manifest.json
├─ qa_report.json
├─ qa_report.md
├─ figures/
└─ previews/

final/
├─ <软件全称>_软件设计与使用说明书.docx
├─ <软件全称>_软件著作权登记申请表.docx
└─ <软件全称>-代码(一般交存版).docx
```

草稿阶段不要把空白模板、占位截图或未分页文件混入 `final/`。

### 2. 创建唯一事实源

先创建 `kit_config.json`。完整字段见 [config-schema.md](references/config-schema.md)，下面只展示轮廓：

```json
{
  "schema_version": 1,
  "software": {
    "full_name": "待确认的软件全称",
    "version": "V1.0",
    "completion_date": "YYYY-MM-DD",
    "original": true,
    "development_mode": "independent",
    "rights_scope": {"mode": "all", "details": ""},
    "hardware_environment": "实际验证的硬件环境",
    "software_environment": "实际验证的软件环境",
    "programming_languages": "语言及版本",
    "source_line_count": 0,
    "function_and_technical_features": "完成证据审查后填写500—1000字"
  },
  "copyright_owners": [],
  "applicant": {},
  "agent": {},
  "deposit": {
    "mode": "general",
    "lines_per_page": 50,
    "front_pages": 30,
    "back_pages": 30
  },
  "confirmations": {
    "software_identity_confirmed": false,
    "completion_date_confirmed": false,
    "rights_confirmed": false,
    "source_scope_confirmed": false,
    "screenshots_reviewed": false,
    "official_rules_verified": false
  }
}
```

不要在示例字段中填写虚构的公司、姓名、身份证号、地址或完成日期。身份与联系字段缺失时可以保留空白；软件全称、版本、完成日期、原创性质、开发方式、权利范围、环境、语言、程序量和功能技术说明则是生成已填写申请表的前置事实。

### 3. 扫描并冻结源码

```powershell
python scripts/scan_project.py `
  --project "D:\path\to\source" `
  --config "copyright-work\kit_config.json" `
  --output "copyright-work\project_facts.json" `
  --manifest "copyright-work\source_manifest.json"
```

扫描返回非零退出码时，先查看 `project_facts.json`。版本冲突、无法解码、真实乱码、私钥、固定密码、邀请码、旧软件名或显式源码缺失不能直接忽略。

冻结后不要修改源码。任何文件变化都应重新扫描，并重建全部下游文件。

### 4. 建立功能证据和图片

阅读界面、路由/API、业务服务、数据访问和外部调用，编写 `feature_evidence.json`。只有证据等级 3 及以上的功能才能作为已实现功能写入说明书，只有等级 4 的功能才能作为带真实运行截图的核心功能。

生成架构图和流程图：

```powershell
python scripts/render_diagrams.py `
  --input "copyright-work\diagrams.json" `
  --output-dir "copyright-work\figures"
```

生成图只用于解释经过验证的架构和流程，不能替代真实运行截图。

### 5. 生成说明书

```powershell
python scripts/build_manual.py `
  --input "copyright-work\manual_content.json" `
  --output "copyright-work\<软件全称>_软件设计与使用说明书.docx"
```

草稿可以使用 `document.allow_missing_images=true` 输出“后期补图”占位。正式版必须关闭该选项，并重新生成、分页和审计。

### 6. 生成申请表底稿

```powershell
python scripts/build_application_form.py `
  --config "copyright-work\kit_config.json" `
  --template "assets\application-form-template.docx" `
  --output "copyright-work\<软件全称>_软件著作权登记申请表.docx" `
  --report "copyright-work\application_report.json" `
  --require-ready
```

`--require-ready` 会在著作权人、申请人或代理人的必要身份/联系字段缺失时返回非零退出码。去掉该参数可以生成身份字段留白且 `ready=false` 的底稿。

若完成日期、开发方式、权利范围、环境、语言、程序量或功能技术说明等前置事实缺失，生成器会直接停止。此时应向申请人补问，而不是自动填入今天的日期或示例值。内置 DOCX 是通用底稿，不是接收平台永久有效的官方表单。

### 7. 生成一般交存代码

先创建引用冻结清单的 `deposit_config.json`：

```json
{
  "source_manifest": "source_manifest.json",
  "software": {
    "name": "与kit_config完全相同的软件全称",
    "version": "V1.0"
  },
  "deposit": {
    "lines_per_page": 50,
    "front_pages": 30,
    "back_pages": 30
  },
  "audit": {
    "repeat_widths": [5, 9, 10, 20],
    "block_on_sensitive": true
  }
}
```

先执行 dry-run：

```powershell
python scripts/build_code_deposit.py `
  --manifest "copyright-work\deposit_config.json" `
  --dry-run `
  --report "copyright-work\code_audit.json" `
  --provenance "copyright-work\deposit_manifest.json"
```

确认没有阻断项后再生成 DOCX：

```powershell
python scripts/build_code_deposit.py `
  --manifest "copyright-work\deposit_config.json" `
  --output "copyright-work\<软件全称>-代码(一般交存版).docx" `
  --report "copyright-work\code_audit.json" `
  --provenance "copyright-work\deposit_manifest.json"
```

### 8. 可选的相似度比较

```powershell
python scripts/audit_similarity.py `
  --manifest "copyright-work\deposit_config.json" `
  --scope deposit `
  --reference "D:\old-materials\old-code.docx" `
  --output "copyright-work\similarity_report.json"
```

`--reference` 可以重复使用，也可以传入旧源码文件、源码目录或旧 manifest。本地检查无法访问登记机构的未知代码库，因此不能给出“绝对无雷同”保证。

### 9. 使用 Word/WPS 锁定真实分页

在 Windows PowerShell 中执行：

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

& scripts/refresh_office.ps1 `
  -Paths @(
    "copyright-work\<软件全称>_软件设计与使用说明书.docx",
    "copyright-work\<软件全称>_软件著作权登记申请表.docx",
    "copyright-work\<软件全称>-代码(一般交存版).docx"
  ) `
  -PdfDir "copyright-work\previews" `
  -Office Auto
```

脚本会更新正文、页眉和页脚域，重复分页直到页数稳定，并导出 PDF。逐页检查截图、表格、目录、标题位置、代码软换行和页码后，把脚本返回的实测页数写入：

```powershell
python scripts/finalize_docx.py `
  --path "copyright-work\<软件全称>_软件设计与使用说明书.docx" --pages <实测页数> `
  --path "copyright-work\<软件全称>_软件著作权登记申请表.docx" --pages <实测页数> `
  --path "copyright-work\<软件全称>-代码(一般交存版).docx" --pages <实测页数> `
  --report "copyright-work\pagination_report.json"
```

`finalize_docx.py` 应在最后一次 Office 保存后执行。完成缓存修正后不要再次用 Office 保存成品，否则 WPS 可能重写错误的 `NUMPAGES` 缓存。

### 10. 执行三件套统一终检

```powershell
python scripts/audit_materials.py `
  --config "copyright-work\kit_config.json" `
  --facts "copyright-work\project_facts.json" `
  --manifest "copyright-work\source_manifest.json" `
  --deposit-manifest "copyright-work\deposit_manifest.json" `
  --evidence "copyright-work\feature_evidence.json" `
  --figures "copyright-work\figures_manifest.json" `
  --manual-content "copyright-work\manual_content.json" `
  --manual "copyright-work\<软件全称>_软件设计与使用说明书.docx" `
  --application "copyright-work\<软件全称>_软件著作权登记申请表.docx" `
  --code "copyright-work\<软件全称>-代码(一般交存版).docx" `
  --pdf-dir "copyright-work\previews" `
  --report-json "copyright-work\qa_report.json" `
  --report-md "copyright-work\qa_report.md"
```

只有在 `qa_report.json` 中满足以下条件，才可以表述为“可进入人工提交复核”：

```json
{
  "submission_ready": true,
  "errors": []
}
```

即使自动门禁全部通过，仍需申请人确认法律事实、签章、证照、官方在线表单和图片内容。

## 一般交存选码规则

默认采用常见的一般交存基线：

| 冻结源码非空行数 | 处理方式 |
|---:|---|
| 小于 3000 行 | 全量交存，不复制、不循环、不填充 |
| 等于 3000 行 | 全量交存 |
| 大于 3000 行 | 冻结顺序的前 1500 行 + 后 1500 行 |

默认每页目标为 50 个源码段落、前后各 30 页。长代码行可能在 Word 中产生软换行，因此仍须检查 PDF 中的实际视觉行和页数。

上述数值是常用工作基线，不应视为永久不变的法规。每次使用时都要核对当期官方要求，并在配置和 QA 报告中记录核对来源与日期。

## 功能证据等级

| 等级 | 典型证据 | 材料中允许的表述 |
|---:|---|---|
| 0 | 只有名称、注释、README 或依赖 | 不申报 |
| 1 | 只有模型、表、Schema、种子或未调用函数 | 通常不写，最多说明预留结构 |
| 2 | 后端函数或 API 存在，但用户路径不完整 | 谨慎说明内部能力及限制 |
| 3 | 交互入口、调用边界、业务处理和结果链完整 | 可以写为已实现功能 |
| 4 | 3 级证据 + 隔离运行验证 + 真实截图 | 可以作为说明书核心操作功能 |

非 Web 项目可以通过 `required_layers` 声明真实调用层，例如桌面应用使用 `ui + domain`，服务库使用 `api + domain`。不要为了通过检查而伪造不存在的 UI 或 HTTP API。

## 会阻断提交的典型问题

- 软件全称、版本、完成日期或权利事实未确认；
- 申请人、著作权人、联系人或签章字段缺失；
- 当前官方规则和表单未核对；
- 源码冻结后发生变化；
- 项目存在多个互相冲突的软件版本；
- 交存内容包含依赖、构建产物、缓存或生成代码；
- 源码或材料中存在乱码、私钥、默认密码、令牌或固定邀请码；
- 少于 3000 行时复制代码凑成 60 页；
- 说明书声称的功能没有完整代码证据；
- 核心功能没有通过运行验证或缺少真实截图；
- 截图含旧名称、个人信息、凭据、路径、错误或拼接痕迹；
- 申请表功能技术说明不在 500—1000 个非空白字符范围内；
- DOCX 含批注、修订、隐藏文本、外链或旧元数据；
- Word/WPS 未完成真实分页，或 PDF 页数与 DOCX 缓存不一致。

完整清单见 [qa-gates.md](references/qa-gates.md)。

## 隐私与安全

- 不要把真实身份证、手机号、病历、订单、住址、人脸或账户信息放进截图和示例配置。
- 不要在 issue、日志、QA 报告或测试夹具中提交 API key、数据库连接串、密码和令牌。
- 在隔离数据库和本地回环地址中运行项目，避免调用生产系统、真实支付、短信、邮件或外部业务接口。
- 报告只记录敏感值类型和位置，具体值应显示为 `<redacted>`。
- 在把源码或生成材料上传 GitHub 前，再做一次仓库级 secret scan 和个人信息检查。

## 当前限制

- 内置申请表是通用填报底稿，不保证与未来接收平台表单完全一致。
- 当前申请表生成器主要面向原创软件、一般交存和最多两名原始著作权人。
- 图片像素检查不能自动证明截图确实来自目标系统，仍需人工或 OCR 复核。
- 功能证据脚本能验证路径、层级和行号，但不能完全替代对业务语义的源码审读。
- 本地相似度检查只能覆盖用户提供的材料，无法访问登记机构或第三方的未知代码库。
- 没有 Word/WPS 时无法确认最终 DOCX 的真实页数与域缓存。
- 自动检查通过不代表权属证明、证件、签章和申请事实已经获得法律确认。

## 开源前检查

在首次推送 GitHub 之前建议完成：

- [ ] 确认仓库不含真实项目源码、客户数据、截图和软著申请人的身份资料；
- [ ] 对 Git 历史执行 secret scan，而不仅是检查当前工作区；
- [x] 已添加 MIT `LICENSE`；
- [ ] 明确维护者、贡献方式和安全问题报告渠道；
- [ ] 配置 CI，至少运行 Python 语法检查、Ruff 和 Skill 结构校验；
- [ ] 在干净环境中验证 `requirements.txt`；
- [ ] 明确本项目不是官方登记平台，也不提供登记通过保证。

## 贡献

欢迎提交能够提高可复现性、材料一致性和风险识别准确度的改进。提交 Pull Request 时请：

1. 不上传真实申请材料、证照、凭据和个人信息；
2. 为规则变更提供匿名化的最小测试样例；
3. 说明对说明书、申请表、代码交存或 QA 门禁的影响；
4. 运行 Ruff、Python 语法检查和 Skill 结构校验；
5. 不以降低检测率为目的加入绕过名称、版本、敏感值或相似度检查的逻辑。

安全漏洞、隐私泄露或凭据暴露问题不宜在公开 issue 中粘贴原始值。开源前请配置独立的安全报告方式。

## 许可证

本项目采用 [MIT License](LICENSE)。使用、修改和分发时请保留许可证及版权声明。贡献者仍需确认其提交内容有权以 MIT 许可证公开，尤其不要提交来源不明的官方表单、第三方模板、真实申请材料或受限制代码。

## 免责声明

本项目及其生成物仅用于材料整理、技术审计和提交前复核，不构成法律意见、权属证明或行政结果承诺。软件著作权登记规则、平台字段、签章方式和鉴别材料要求可能调整；使用者应自行核对官方信息，并对提交内容的真实性、完整性和合法性负责。
