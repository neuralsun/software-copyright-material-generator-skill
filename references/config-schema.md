# 配置与中间文件规范

## 目录

1. 唯一事实源
2. `kit_config.json`
3. 源码扫描配置
4. `feature_evidence.json`
5. `figures_manifest.json`
6. `diagrams.json`
7. `manual_content.json`
8. `deposit_config.json`
9. 文件派生与一致性规则

## 1. 唯一事实源

使用 UTF-8 JSON。`kit_config.json` 是名称、版本、日期、技术环境、权利事实和申请人字段的唯一可编辑源；`project_facts.json` 与 `source_manifest.json` 由扫描器生成，不手工改程序量、哈希或文件列表。`manual_content.json`、`deposit_config.json` 和申请表内容必须由两者派生。

相同事实只维护一份：

| 事实 | 权威来源 |
|---|---|
| 软件全称、简称、登记版本 | `kit_config.json/software`，用户确认 |
| 完成日期、开发方式、权利范围、发表情况 | `kit_config.json/software`，用户确认 |
| 著作权人、申请人、联系人、代理人 | `kit_config.json`，用户或正式证照提供 |
| 第一方源码文件、顺序、编码、哈希、程序量 | `source_manifest.json` |
| 功能是否实现 | `feature_evidence.json` + 代码与运行证据 |
| 截图是否真实、是否复核 | `figures_manifest.json` |
| 实际页数 | Word/WPS 返回结果和导出 PDF |

## 2. `kit_config.json`

以下为通用骨架。尖括号内容必须替换；未知主体身份和联系字段用空字符串，不得用“待定”“某公司”或虚构号码占位。软件全称、版本、完成日期、原创性质、开发方式、权利范围、技术环境、语言、程序量和功能技术说明是生成申请表的前置事实，不能留空后指望生成器代填。

```json
{
  "schema_version": 1,
  "software": {
    "full_name": "<登记软件全称>",
    "short_name": "",
    "version": "V1.0",
    "completion_date": "YYYY-MM-DD",
    "original": true,
    "development_mode": "independent",
    "rights_scope": {"mode": "all", "details": ""},
    "classification_no": "",
    "hardware_environment": "<处理器、内存、存储、网络等>",
    "software_environment": "<操作系统、运行时、数据库、浏览器等及版本>",
    "programming_languages": "<语言及版本>",
    "source_line_count": 0,
    "function_and_technical_features": "<去空白后500—1000字，待证据矩阵完成后填写>",
    "first_publication_date": "",
    "first_publication_country": "",
    "first_publication_city": ""
  },
  "copyright_owners": [
    {
      "name": "",
      "credential_no": "",
      "nationality": "",
      "address": ""
    }
  ],
  "applicant": {
    "name": "",
    "address": "",
    "postal_code": "",
    "contact": "",
    "phone": "",
    "mobile": "",
    "email": "",
    "fax": ""
  },
  "agent": {},
  "deposit": {
    "mode": "general",
    "lines_per_page": 50,
    "front_pages": 30,
    "back_pages": 30,
    "long_line_warning": 120
  },
  "source": {
    "source_order": [],
    "exclude_globs": [],
    "forbidden_terms": []
  },
  "manual": {
    "target_page_range": [8, 30],
    "minimum_images": 2
  },
  "application": {"expected_pages": 3},
  "confirmations": {
    "software_identity_confirmed": false,
    "completion_date_confirmed": false,
    "rights_confirmed": false,
    "source_scope_confirmed": false,
    "screenshots_reviewed": false,
    "official_rules_verified": false
  },
  "official_rules": {
    "receiving_platform": "",
    "jurisdiction": "",
    "checked_at": "",
    "official_source": ""
  }
}
```

枚举与约束：

- `version` 使用 `V1.0`、`V1.0.0` 等明确格式；三份材料、源代码版本常量和截图必须一致。
- `development_mode` 支持 `independent/独立开发`、`collaborative/合作开发`、`commissioned/委托开发`、`assigned/下达任务开发`。
- `rights_scope.mode` 支持 `all/全部` 或 `partial/部分权利`；部分权利必须写 `details`。
- 当前申请表生成器只处理原创软件与一般交存；不满足时停止并改用适合的正式流程，不要强行套模板。
- `first_publication_date` 为空时国家、城市也必须为空；填写日期时两者必须同时填写。
- 申请表模板最多容纳两名原始著作权人；更多主体必须使用当期官方载体处理。
- 生成申请表前，把 `source_manifest.json/program_nonblank_lines` 原样写入 `software.source_line_count`。
- `function_and_technical_features` 去除全部空白后必须为 500—1000 字。
- `agent` 完全为空表示无代理；只要填写任一代理字段，就必须补齐名称、地址、邮编、联系人、手机和邮箱。

法律字段最低完整性：每个 `copyright_owners[]` 填名称、证件号码、国籍/地区和地址；`applicant` 填名称、地址、邮编、联系人、手机和邮箱。电话、传真可按实际情况留空。申请人签章、签署日期及证照真伪仍须人工或平台核验。

## 3. 源码扫描配置

`source` 只在确有需要时设置。不要用空数组覆盖扫描器的默认排除目录。可用字段：

- `include_extensions`：要扫描的源码扩展名；通常省略以使用默认值。
- `exclude_parts`：完整替换默认排除目录，风险高；只有审查过全部值时才设置。
- `exclude_globs`：追加项目特定排除模式。
- `include_globs`：非空时只保留命中的文件。
- `source_order`：项目根目录下的相对 POSIX 路径，作为程序开头顺序；未列文件按语义顺序追加。
- `forbidden_terms`：旧软件名、旧简称或不得残留的项目标识。

扫描器默认排除 `.git`、虚拟环境、`node_modules`、`vendor`、`dist`、`build`、缓存、覆盖率目录、上传/媒体/材料目录等，并排除压缩后 JS/CSS、source map 和锁文件。扫描完成后仍须人工核对第一方范围。

## 4. `feature_evidence.json`

每个功能使用稳定 ID。顶层可使用 `features` 数组；终检也兼容以 ID 为键的对象映射。

```json
{
  "schema_version": 1,
  "software": {"full_name": "<同kit_config>", "version": "V1.0"},
  "features": [
    {
      "id": "F01",
      "name": "<功能名称>",
      "claim": "<允许写进说明书的精确表述>",
      "evidence_level": 4,
      "status": "implemented",
      "required_layers": ["ui", "api", "domain"],
      "user_path": ["<进入入口>", "<提交操作>", "<结果反馈>"],
      "code_refs": [
        {"layer": "ui", "path": "frontend/src/...", "line": 1, "symbol": "..."},
        {"layer": "api", "path": "backend/...", "line": 1, "symbol": "..."},
        {"layer": "service", "path": "backend/...", "line": 1, "symbol": "..."},
        {"layer": "persistence", "path": "backend/...", "line": 1, "symbol": "..."}
      ],
      "runtime_verified": true,
      "runtime_notes": "<构建/冒烟测试命令和结果>",
      "screenshots": ["figures/S01.png"],
      "limitations": ""
    }
  ],
  "excluded_claims": [
    {"name": "<未实现或无法证明的功能>", "reason": "<证据断点>"}
  ]
}
```

`layer` 使用 `ui`、`api`、`business`、`service`、`persistence` 等明确值。所有路径必须存在于冻结 manifest，行号不得越界。Web 功能默认要求 UI、API 和一个业务端层。桌面、CLI、插件或库应显式填写符合真实调用边界的 `required_layers`，例如 `["ui", "domain"]` 或 `["api", "domain"]`，不能为通过检查而伪造不存在的层。4 级还必须 `runtime_verified=true` 且至少有一张真实存在的截图。

## 5. `figures_manifest.json`

```json
{
  "schema_version": 1,
  "figures": [
    {
      "id": "S01",
      "path": "figures/S01.png",
      "kind": "runtime_screenshot",
      "feature_ids": ["F01"],
      "caption": "<与界面内容一致的图题>",
      "captured_at": "YYYY-MM-DDTHH:MM:SS+08:00",
      "source": "local isolated runtime",
      "review": {
        "software_name_correct": true,
        "old_name_absent": true,
        "pii_absent_or_redacted": true,
        "credential_absent": true,
        "error_absent": true,
        "splice_absent": true,
        "manually_reviewed": true
      }
    }
  ]
}
```

架构图和流程图使用 `kind=generated_diagram`，并列出其证据功能 ID。缺图占位使用 `kind=placeholder`，它永远不能通过最终门禁。

## 6. `diagrams.json`

```json
{
  "architecture": {
    "title": "系统总体架构",
    "output": "architecture.png",
    "layers": [
      {"id": "ui", "label": "交互层", "title": "用户界面", "nodes": [{"id": "portal", "label": "业务入口"}]},
      {"id": "service", "label": "业务层", "title": "业务服务", "nodes": [{"id": "core", "label": "核心处理"}]}
    ],
    "edges": [{"from": "portal", "to": "core", "bidirectional": true}]
  },
  "flow": {
    "title": "核心业务流程",
    "output": "workflow.png",
    "lanes": [{"id": "user", "label": "用户"}, {"id": "system", "label": "系统"}],
    "nodes": [
      {"id": "start", "lane": "user", "column": 0, "label": "进入功能", "shape": "terminal"},
      {"id": "submit", "lane": "user", "column": 1, "label": "提交信息"},
      {"id": "process", "lane": "system", "column": 2, "label": "执行业务处理"},
      {"id": "end", "lane": "user", "column": 3, "label": "查看结果", "shape": "terminal"}
    ],
    "edges": [{"from": "start", "to": "submit"}, {"from": "submit", "to": "process"}, {"from": "process", "to": "end"}]
  }
}
```

## 7. `manual_content.json`

生成器接受段落、标题、图片和表格块。为便于终检，每个描述功能的块或其父级容器都附加 `feature_ids`；生成器会忽略该审计字段，终检会读取。

```json
{
  "software": {"name": "<同kit_config.full_name>", "version": "V1.0", "owner": "<著作权人>"},
  "document": {
    "title_suffix": "软件设计与使用说明书",
    "body_page_start": 1,
    "allow_missing_images": false,
    "allow_missing_captions": false,
    "toc": {"title": "目录", "min_level": 1, "max_level": 3},
    "footer": {"prefix": "第 ", "suffix": " 页", "show_section_total": true}
  },
  "chapters": [
    {
      "title": "1 软件设计说明",
      "sections": [
        {
          "title": "1.1 软件概述",
          "blocks": [{"type": "paragraph", "text": "<连贯长段落>"}]
        },
        {
          "title": "1.2 总体架构",
          "feature_ids": ["F01"],
          "blocks": [
            {"type": "paragraph", "text": "<只描述已证实的架构和数据流>"},
            {"type": "figure", "path": "figures/architecture.png", "caption": "系统总体架构", "width_cm": 15.0}
          ]
        }
      ]
    },
    {
      "title": "2 软件使用说明",
      "sections": [
        {
          "title": "2.1 核心功能操作",
          "core_feature_ids": ["F01"],
          "blocks": [
            {"type": "paragraph", "text": "<入口、输入、处理、输出和异常反馈>"},
            {"type": "figure", "path": "figures/S01.png", "caption": "核心功能运行界面", "width_cm": 15.0}
          ]
        }
      ]
    }
  ]
}
```

使用 `feature_ids` 标记一般已实现功能，使用 `core_feature_ids` 标记配有运行截图或作为核心操作说明的功能；后者必须达到证据等级 4。块类型：`paragraph` 支持 `text`、`alignment`、`bold`、`italic`；`heading` 支持 `text`、`level`；`figure` 支持 `path`、`caption`、`width_cm`；`table` 支持 `caption`、`headers`、`rows`、`widths_cm`。图片路径相对 JSON 文件或 `--base-dir` 解析。目录和页码是动态域，必须经 Office 刷新。

## 8. `deposit_config.json`

```json
{
  "source_manifest": "source_manifest.json",
  "software": {"name": "<同kit_config.full_name>", "version": "V1.0"},
  "source": {"tab_size": 4, "max_file_bytes": 20971520},
  "deposit": {"lines_per_page": 50, "front_pages": 30, "back_pages": 30},
  "audit": {
    "repeat_widths": [5, 9, 10, 20],
    "long_line_warning": 120,
    "block_on_sensitive": true,
    "block_on_high_login_risk": false,
    "forbidden_selected_paths": []
  },
  "similarity": {"references": []}
}
```

`source_manifest` 相对 `deposit_config.json` 解析。不要复制 manifest 的文件数组后再手改；生成器会复验文件 SHA-256、物理行数、非空行数和整包快照。`forbidden_selected_paths` 只用于阻止误纳入明确不属于第一方程序的路径，不得把它当作规避审查的删码工具。

相似度参考可以是旧 DOCX、单个源码、源码目录或旧 `source_manifest.json`：

```json
{"path": "../old/source.docx", "label": "旧项目代码", "type": "docx"}
```

## 9. 文件派生与一致性规则

按以下顺序派生，禁止逆向手改：

```text
kit_config.json + 源码
  -> project_facts.json + source_manifest.json
  -> feature_evidence.json + figures_manifest.json
  -> diagrams.json + manual_content.json + deposit_config.json
  -> 三份DOCX + deposit_manifest.json
  -> Office实测页数/PDF
  -> qa_report.json + qa_report.md
```

改名称、版本、完成日期、源码范围、功能边界或截图后，从受影响的上游节点重新生成全部下游文件。最终三份材料不得通过搜索替换各自单独修补，否则会失去可追溯一致性。
