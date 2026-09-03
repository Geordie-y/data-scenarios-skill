# 内容 JSON 契约

顶层字段为 `document_mode`、`allow_assumptions`、`sources`、`industries`、`limits`。默认正式模式 `formal`；本模板的补全式编制使用 `allow_assumptions: true`。仅在用户明确要求核对稿或禁止推演时使用 `review`。

## 段落 Block

`{"text":"通过接口同步业务数据，形成运行看板。","kind":"proposal","sources":["S1"]}`

`fact` 是来源支持的事实，必须关联证据；`proposal` 是方案设计；`assumption` 是允许自拟的参数或假设；`missing` 只允许出现在 review 模式。所有类型仅供内部核对，生成器不添加“建议/假设/待补”前缀。

每个 Block 是一个自然段，text 不含换行。正式 text 应已写为成稿，不允许资料页码评述和缺失清单。事实仍需出处，不能将自拟内容标为 fact。

## 结构

- `sources` 数组：每项有 `id`、`file`、`locator`、`excerpt` 字符串。用于独立编制说明，不显示在正文。
- `limits` 字符串数组：自拟内容依据、来源冲突、真实数据可用性及机构身份等核验事项。
- `industries` 数组：每项有 `title` 和非空 `scenes` 数组。标题不手填编号。
- 每个 scene 有 `name`，以下单个 Block：`description`、`provider`、`payer`；以下非空 Block 数组：`background`、`solution`、`data`、`beneficiaries`；以及 `revenue`、`cases`。
- `provider` 与 `payer` 同时用于表格和正文，保持短而明确，细节放在 data 和 beneficiaries 中。

## 收益对象 revenue

包含 Block `billing`（计费方式）、`subject`（收入测算主体名称）、`costs`（成本构成或核算方式），字符串 `period`、`formula`、`result_name`、`unit`，以及 parameters 数组。

参数结构：`{"id":"N","label":"付费机构数","value":3,"unit":"家","kind":"assumption","sources":[]}`。

所有 formal 参数有数值，不能用 null 占位。fact 参数需证据；assumption 参数需 allow_assumptions=true。每组自拟值在 limits 记录为何选用、哪些内容未作市场验证。使用合适规模的规划参数，不自拟实际交易事实。

formula 只允许参数、数字、加减乘除及括号，例如 `N * P`；参数 ID 为字母开头的字母数字下划线。用 Decimal 计算。正文自动显示测算口径、参数及代入公式；不逐个打印“假设值”。单位必须自行审核，例如 N 为家、P 为万元/家/年，结果为万元/年。

## 案例

正式材料必须有案例内容，真实项目与典型应用分开建模。

真实项目：`{"type":"actual","title":"有证据的项目名称","detail":{"text":"项目已证实的应用内容。","kind":"fact","sources":["S2"]}}`。

自拟示例：`{"type":"example","title":"区域机构运营质检","detail":{"text":"某区域运营机构汇集工单记录，系统按业务规则生成异常清单，运营人员据此核对记录并安排回查。","kind":"proposal","sources":["S1"]}}`。

生成器给 example 标题加“应用示例：”，同时进入表格和正文。这是应用方式演示，不是虚构落地案例。不能为了消除标签改填 actual。

## 验证

`generate_docx.py --check-only` 校验完整结构、正式文风、来源和测算；`check_docx.py` 校验实际成品中的正式文风、章节、摘要及版式。它们不代替事实核对和逐页视觉检查。正文原样展示写好的 text，不做掩盖不确定性的机械删词。
