---
name: w-layout-to-excel
description: 把 GPT 识别的仓库布局图货架 JSON（扁平 cells 或 x_ranges/y_ranges 分组格式）转成数字化 Excel 布局：一行=一条货架排、格子=货架组、通道留 1 空行/列、背靠背两排相邻不留白、统一浅灰底纹、格内写货位码（区-通道-货架组-层号-货位序号，段结构/计数范围/蛇形可配置）。用户给出布局图识别 JSON / 提到"布局转 Excel"、"货架识别"、"导出货位 Excel"时触发。
---

# w-layout-to-excel

把「GPT 识别的仓库平面图货架 JSON」变成「可编辑、带货位码的 Excel 布局」。

## 用途

- 输入：仓库平面图 + GPT 输出的货架识别 JSON（用户自己跑 GPT；没有就用提示词模板）
- 输出：`.xlsx` —— 格子 = 货架组，通道留空行/列，**背靠背两排相邻不留白**，统一浅灰底纹（`style.fill`，默认 `D9D9D9`），格内 = 该格子全部层的货位码（换行分隔）。三种布局：`classic`（一行=一排货架）、`transpose`（一列=一排）、**`mirror`（一条货架线=一列，与图纸上下左右同向，横向通道=列内空行）**
- 术语：**排** = 一条货架线（`rack`）；**货架组** = 排上的一个格子（`rack_group`/格子）；**block** = 背靠背等紧邻排的分组（Excel 中相邻、不留通道）

## 目录

```
scripts/
  normalize.py         ① GPT JSON -> layout.json（格式校验/背靠背拆分/低置信度透传）
  generate.py          ④ layout.json + rules.json -> .xlsx（--drop 可剔除记录）
  overlay.py           ②③ 核对页：识别结果/编号预览渲染到底图 HTML（默认自动打开，--no-open 关闭）
  requirements.txt     仅 openpyxl
reference/
  gpt6_prompt.md       给 GPT 的提示词模板（分组格式为主，扁平为备选）
  gpt6_input.json      真实识别输入样例（36 记录，x_ranges/y_ranges 分组格式）
  gpt6_sample_rules.json / gpt6_sample.xlsx   通用默认规则回归基线
  gpt6_case_rules.json / gpt6_case.xlsx       真实案例基线（E/F 分区 + aisle_sides 奇偶编号 + mirror + --drop 绿色阵列）
  sample_layout.json / sample_rules.json / sample.xlsx   内部格式最小样例
```

## 工作流程

### 0. 依赖

首次运行安装（放入脚本目录的 `_venv` 或项目 venv）：
```
python -m venv <workdir>/_venv
<workdir>/_venv/bin/pip install -r <skill>/scripts/requirements.txt
```
脚本统一用该 venv 的 python 调用。安装失败时提示用户允许 pip install。

### 1. GPT 识别（用户侧）

用户已有识别 JSON → 直接进入第 2 步。没有就把 `reference/gpt6_prompt.md` 的模板提示词给用户，让用户附图发给 GPT，拿回 JSON。

### 2. 规范化（normalize.py）

```
python <skill>/scripts/normalize.py <gpt6.json> -o <workdir>/layout.json
```

- 自动识别两种输入格式：**`rack_positions[].x_ranges/y_ranges` 分组**（默认推荐，token 省约 3 倍，大图不截断）与 **扁平 `racks[].cells`**（缺格场景/小图备选，无满格假设）。格式取舍见 `reference/gpt6_prompt.md` 备注
- 背靠背：同记录内深度方向小间隙的并列列 → 拆成多个物理排、归同一 `block`
- 排长方向的大间隙（横向通道）→ 拆成不同 block
- 明确报错（定位到记录/字段）：JSON 非法、缺 id / x_ranges / y_ranges / cells、区间起点>=终点、坐标非法、无货架
- **低置信度透传**：`confidence`（数字或"较低"等文字）、`review_reason`、`last_y_range_truncated`、`excluded_regions`、尺寸异常格子 → 写入 `layout.json` 的 `notes` / `needs_review`，不静默丢弃

**判定（必须）**：normalize 成功后，**无论 `confidence` 高低、`needs_review` 是否为空，都必须生成核对页并让用户确认**。脚本默认自动打开浏览器：
```
python <skill>/scripts/overlay.py <workdir>/layout.json --image <原图.png> -o <workdir>/check.html
```
页面里红虚线框=低置信记录、蓝框=正常、黄=GPT 排除区域，可勾选图层。**生成后必须暂停**，把页面（连同 `notes`）给用户看并明确询问：a) 接受当前结果继续；b) 修正 JSON（或换图重跑 GPT）后重跑 normalize；c) 剔除问题记录——用第 4 步 `--drop 记录id` 即可，不必改 layout.json。用户确认前不得进入第 3 步。无界面/CI 环境加 `--no-open`（不自动开浏览器），并自行说明页面路径。`--strict` 可把全部警告当错误。

### 3. 询问用户（→ rules.json）

按顺序问（选择 + 允许手输，给默认值）：

1. **区域与前缀**：展示货架记录 ID 列表（`layout.json` 的 `racks[].record`，或直接看核对页），问"哪些记录属于哪个区、各区前缀"。没分区就单区（`"rack_ids": ["*"]`）。**前缀匹配**：`"R01"` 覆盖其拆分排 `R01-1`…`R01-N`。未匹配任何区的记录 generate 会打警告，须补齐或确认。
2. **货位码段结构**：可用段 `zone`（区前缀）/ `aisle`（通道号）/ `rack`（区内排序号）/ `rack_group`（货架组=格子号）/ `level`（层号）/ `position`（层内货位序号）。逐个确认启用段、顺序、分隔符（默认 `-`）、补零位数（默认 aisle 2 / rack_group 2 / 其他 0）。
3. **货架组计数范围**：两种模式二选一——
   - 常规 `rack_group_scope` = `aisle`（通道/块内，默认）/ `zone`（区内连续）/ `global`（全仓连续）/ `rack`（排内）；
   - **`numbering_mode: "aisle_sides"`（通道左右奇偶）**：通道=相邻货架记录间的真实空隙（区内从 01 编）；左侧排用奇数 01,03…、右侧排用偶数 02,04…（从上到下 2n-1/2n）；背靠背左列挂左通道（偶）、右列挂右通道（奇）；区首单列记奇、其余单列记偶；左边缘空地是否算通道 01（成对开头的区默认算）。**追加确认**：跨段续号 `continuous_segments`（默认 true=被横向通道切开的段连续编号；false=每段从 01 重新起编）。参考 `reference/gpt6_case_rules.json`。
4. **货位序号**：一格一层内有几个货位（`position_count`，默认 1，不需要就去掉 `position` 段）。
5. **层数**：默认全局 `3`；可按排 `levels.per_rack`。
6. **蛇形**：默认启用、相邻排交替方向；确认起点角（默认 bottom_left）与首排方向（默认 right）。`aisle_sides` 模式下不适用（自动关闭）。
7. **方向确认**：三种布局——`classic`（默认，排=行）；`transpose`（排=列，横向通道切段会拆成独立列组）；**`mirror`（推荐核对方向用，与图上下左右完全对应：一条货架线=一列、背靠背相邻、记录间空 1 列、横向通道=列内空行）**。`orientation.mode` 控制。

把答案写成 `rules.json`（对照 `reference/gpt6_sample_rules.json`）。

### 4. 生成 Excel（generate.py）

```
python <skill>/scripts/generate.py <workdir>/layout.json -r <workdir>/rules.json -o <workdir>/layout.xlsx [--drop R36,R05]
```
- `--drop`：按记录 id 前缀匹配剔除（如低置信的绿色阵列/R36），不动 layout.json，重跑零成本
- 生成前**必须**做**编号预览**（脚本默认自动打开）：`overlay.py <layout.json> --image <原图.png> --rules rules.json [--drop ...] -o check.html`——每格显示货架组号、外框按区分色、hover 看完整货位码，在图上直接核对编号。**必须暂停并让用户确认编号与分区**后才 generate；改过 `rules.json`/`--drop` 重跑后必须重新预览确认。无界面/CI 环境加 `--no-open`
- 收尾：打开 Excel 给用户核对；有问题回到第 3 步改 `rules.json`（或第 2 步改 JSON/加 `--drop`）重跑——normalize 与 generate 可单独重跑

## rules.json 参考

```json
{
  "zones": [{"prefix": "A", "rack_ids": ["*"]}],
  "code": {
    "segments": ["zone", "aisle", "rack_group", "level"],
    "sep": "-",
    "pads": {"zone": 0, "aisle": 2, "rack_group": 2, "level": 0},
    "rack_group_scope": "aisle",
    "position_count": 1,
    "serpentine": {"enabled": true, "start_corner": "bottom_left", "first_dir": "right"}
  },
  "levels": {"mode": "global", "count": 3},
  "orientation": {"mode": "classic"},
  "style": {"fill": "D9D9D9", "stroke": "808080"},
  "sheet_name": "layout",
  "cell_size": {"col_width": 14}
}
```
`orientation.mode`：`classic`（默认）/ `transpose`（=旧行为，`transpose:true` 仍兼容）/ `mirror`（镜像图纸，仅支持竖排 V 为主的仓库）。货位码另支持：
- `numbering_mode: "aisle_sides"`：通道=相邻货架记录间的真实空隙（区内从 01 编），左侧排奇数 01,03…、右侧排偶数 02,04…（从上到下 2n-1/2n），背靠背左列挂左通道、右列挂右通道，区首单列记奇、其余单列记偶（要求每条记录 ≤2 并列列，否则明确报错）。完整示例见 `reference/gpt6_case_rules.json`
- `continuous_segments`（默认 true）：同列被横向通道切开的段是否连续编号；false = 每段重新从 01 起编

## layout.json 格式（normalize 产物，可手改后直接跑 generate.py）

```json
{
  "confidence": 0.45, "needs_review": ["R24"], "notes": ["..."],
  "racks": [
    {"id": "R01-1", "record": "R01", "dir": "V", "block": "R01-s1",
     "confidence": 1.0, "notes": [],
     "bays": [{"x0": 87, "y0": 241, "x1": 98, "y1": 266}]}
  ]
}
```
手改时：调 `block` 可合并/拆分通道留白；删除问题 `racks[]` 条目可剔除误识别；`fill`/`stroke`（0-1 RGB 数组）可选，缺省用 rules 的 `style`。

## Troubleshooting

- **GPT JSON 缺字段/坐标非法** → normalize 明确报错到"记录 x 字段 y"，把报错给用户，改 JSON 重跑，不要硬编。
- **分组格式的满格假设**：x_ranges×y_ranges 会把缺格补齐。图上明显缺格时建议用 `reference/gpt6_prompt.md` 的扁平格式重跑 GPT。
- **低置信度条目**（绿色阵列、边缘错位列等）→ 走第 2 步判定流程：核对页确认后用 `--drop` 剔除或接受。
- **背靠背误判**（间隙阈值不符）→ 手改 `layout.json` 的 `block` 字段后只重跑 generate.py。
- **aisle_sides 报"记录 X(N列) 超出"** → GPT 把 3 条以上并列列分进了同一记录；在 layout.json 把多余列拆成独立 record，或改用 rack_group_scope 计数。
- **mirror 报"只支持竖排(V)"** → 横排仓库用 classic/transpose（mirror 的行锚点算法按竖排设计）。
- **警告"未匹配任何分区"** → rules.zones 的 rack_ids 漏了这些记录，货位码 zone 段会是空；补 zone 或确认可接受。
- **mirror 下个别列与主网格错位**（如 R35 类不被横向通道切断的列）→ 其格子会被就近吸附到全局行锚点，通道处表现为该列连续、其他列空行，这是如实呈现；编号不受影响。
