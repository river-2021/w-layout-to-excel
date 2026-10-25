# w-layout-to-excel 改进实施计划（实战经验沉淀）

依据：2026-10-19 完整实战（36 记录识别 → 低置信度确认 → 剔除 → E/F 分区 + aisle_sides 奇偶编号 + mirror 布局 → layout.xlsx）。用户已确认全部执行。

## 1. generate.py 五项改动

### 1.1 `--drop` 参数（A2）
- `ap.add_argument("--drop", default="", help="剔除的记录/货架 id，逗号分隔（前缀匹配）")`
- 复用匹配语义：`rid == p or rec == p or rid.startswith(p + "-")`（与 zones 前缀匹配一致）
- 模块级 `drop_matches(rack, patterns)` helper（overlay.py 复用）
- main 中过滤 racks，stderr 打印 `[generate] --drop 已剔除 N 排（模式: ...）`；剔除后为空则报错

### 1.2 `code.continuous_segments`（A3，默认 true）
- `assign_aisle_sides(racks_ordered, zone_of, continuous=True)` 增加 with 参数
- true（现行为）：同列跨段连续编号（col_bays 合并排序后 i=1..n → 2i-1/2i）
- false：每段（rack）独立从 i=1 起编
- generate_codes 透传：`continuous = code.get("continuous_segments", True)`

### 1.3 C5：aisle_sides 遇 3+ 并列列明确报错
- assign_aisle_sides 中 `len(col_list) > 2` 时收集违规记录，循环结束后：
  `SystemExit("[generate] 错误: numbering_mode=aisle_sides 要求每条记录最多 2 条并列列（背靠背一对）；记录 X/N 有 3 列。请在 layout.json 拆分该记录或改用 rack_group_scope 计数")`
- 不再静默 KeyError

### 1.4 C6：mirror 模式横排守卫
- mirror_specs 开头：`if not v_major: SystemExit("[generate] 错误: orientation.mode=mirror 目前只支持竖排(V)货架为主的仓库；横排仓库请用 classic/transpose")`

### 1.5 C7：未分区警告
- generate_codes 内（zone_of 构建后）：rules.zones 非空时收集 `zone_of(rid) is None` 的 record，stderr 打印
  `[generate] 警告: 以下货架未匹配任何分区（zone 段为空）: R99, ...`
- 需 `import sys`

## 2. 新增 scripts/overlay.py（A1）

```
python overlay.py <layout.json> --image <png> -o check.html [--rules rules.json] [--drop R36,...]
```

- `sys.path` 注入脚本目录，`import generate` 复用 order_racks / generate_codes / build_zone_lookup / drop_matches（零重复逻辑）
- 无 rules：记录框分色（蓝=正常 / 红虚线=confidence<0.5）+ 图层开关（记录框/格子/排除区域）
- 有 rules：先跑完整编号，再渲染——记录框按区分色（调色板：#2196F3/#FF9800/#4CAF50/#9C27B0/#00BCD4...），每格显示 rack_group 段值（小字号，从 code 按 segments+sep 切出），hover title 显示完整码与归属；编号图层默认开
- HTML 自包含（数据内嵌 JSON），图片用绝对路径；图例+图层勾选；打印 stdout 摘要 {ok, output, records, cells, zones}
- 仅标准库 + 复用 generate（openpyxl 随 generate 导入，无新依赖）

## 3. 测试（全部在 /var/folders/.../opencode/layout-test）

1. py_compile 两脚本
2. **case 可复现性**：normalize(新 gpt6_input.json) → generate `--drop R05,R07,R08,R09,R11,R12,R13,R15,R16,R17,R19,R36 -r rules.json -o case.xlsx` → 与现有 layout.xlsx 逐格对比（行数/列数/抽查码/格数=848）
3. continuous_segments=false：R01 左列应每段从 02 重启（10 格段 02-20、6 格段 02-12、02-12）
4. C5：构造 3 列记录的小输入 → 明确报错
5. C6：构造 H-major layout + mirror → 明确报错
6. C7：rules 去掉 R35 → stderr 警告列出 R35
7. overlay.py 无 rules / 有 rules 各生成一页并 open 给用户看

## 4. 刷新回归基线（B4）

- `reference/gpt6_input.json` ← 当前 layout_now.json（36 记录原始输入）
- 新增 `reference/gpt6_case_rules.json` = 实战最终 rules（E/F 分区、aisle_sides、mirror、1 层、continuous_segments: true 显式写入）+ 剔除清单注释性 README 不做，命令写进 SKILL.md
- 生成 `reference/gpt6_case.xlsx`（normalize + --drop + case_rules，校验 848 格）
- 用新输入重生成 `reference/gpt6_sample.xlsx`（classic 默认规则，保持通用基线）

## 5. SKILL.md 文档更新（D8-D10）

- 「用途」：补 mirror 模式一句（一条货架线=一列、与图同向）
- 「目录」：加 overlay.py、gpt6_case_rules.json / gpt6_case.xlsx
- 第 2 步判定：改为"用 overlay.py 生成核对页（红框=低置信），连同 notes/needs_review 给用户选择"
- 第 3 步问答：
  - 区域指派改为对照 overlay 页
  - 新增 aisle_sides 分支问题：奇偶规则确认、单列归属（区首奇/其余偶）、左边缘是否算通道、跨段续号（continuous_segments）
  - 剔除类操作推荐 `generate.py --drop`（不必改 layout.json）
- 第 4 步：加"生成前可用 overlay.py --rules 做编号预览"；命令示例带 --drop
- rules.json 参考：加 `"numbering_mode"`、`"continuous_segments"` 键
- Troubleshooting 新增：
  - 3+ 并列列报错的处理（拆 record 或换计数模式）
  - mirror 横排不支持 → classic/transpose
  - 未分区警告含义
  - R35 类离网格列在 mirror 下被多数行锚点吸收、可能自成一空行（设计行为）
