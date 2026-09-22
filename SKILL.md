# SKILL：东南大学 × 国网电科院（南瑞）汇报 PPT 模板

面向**操作这套模板的 agent**。读完你应该能独立改文字、加页、插图片/公式/表格、加章节、改配色，并且不踩已知的雷。

**本目录只含 pptx 相关内容**（beamer 版不在范围内，已剔除）。

---

## 0. 何时使用 / 先读什么

**使用时机**：用户要求修改这套汇报模板、往里面加内容、调整版式、改配色，或把这套模板套用到新的一次汇报。

**动手前必须做三件事**：

1. 读 `design/tokens.json` —— 所有颜色/字号的唯一真源。**任何色值都从它取，不要凭记忆写 hex。**
2. 读 `build/build_pptx.py` 顶部的 `DECK` 字典 —— 汇报的文案（标题、汇报人、全部章节名）都在这里。
3. 读 `build/slides.py` 的 `SLIDE_PLAN` —— 内容页的结构。它随附的是一份
   18 页 / 5 章的功能完备示例（每种能力各出现一次），直接照抄改即可。

**最小自足集**：本目录复制到任何位置都能构建。已实测从零重建通过。

```
SKILL.md                  ← 本文档
design/
  tokens.json             ← 单一真源：颜色 / 字号 / 栅格
  check_tokens.py         ← 对比度校验器（改色后必跑）
assets/
  seu-logo.png            ← 东大完整校标（校徽 + 书法校名）330×128
  seu-seal.png            ← 仅校徽方形（内容页右上用）578×578
  nari-logo.png           ← 南瑞完整 lockup，透明底，纯平 #00706B  1088×206
  figures/                ← 插图 + manifest.json（像素尺寸，供等比放置）
  formulas/               ← 公式/伪代码的透明 PNG + manifest.json
build/
  build_pptx.py           ← 主构建器（A/B/C 三阶段）
  _pptx_kit.py            ← OOXML 底层封装（形状 / 版式 / 部件注册）
  slides.py               ← **内容页**：随附的 18 页功能完备示例
  make_formulas.py        ← 公式 / 伪代码渲染（LaTeX → 透明 PNG）
  prep_figures.py         ← 从源目录收集插图并统一命名（可选）
  extract_math.py         ← 从 LaTeX 源抽取公式块（可选）
  prep_refs.py            ← 从源 PDF 抽取参考文献（可选）
  _com.py                 ← PowerPoint COM 公共层（等退出 + 重试，不杀进程）
  _audit.py               ← 交付前机器审查（公式倍率 / 出血 / 可读性 / 空标题）
  _render.py / _pptopen.py / _contact.py   ← 验证工具
out/                      ← 产物
错例/                      ← 反例对照图 + README（给格式审查用，见 §7.3）
```

`prep_figures.py` / `extract_math.py` / `prep_refs.py` 三个"取材"脚本**都是可选的**：
它们把既有材料（插图目录、LaTeX 源、PDF）转成管线能吃的中间件。不用就留空目录，
`make_formulas.py` 与 `slides.py` 都能在没有它们产物的状态下工作。见 §5.2b。

**随附示例**：`slides.py` 是 18 页 / 5 章的功能完备示例 —— 管线支持的每种能力
（双栏、图文、卡片、表格、公式块、公式墙、伪代码、关键数字、小结、参考文献）
都出现至少一次，所以它同时是 cookbook。换成你自己的内容就是改这个文件。

---

## 1. Do / Do Not

### Do

| | |
|---|---|
| **Do** | 改配色改字号 → 只改 `design/tokens.json`，然后重跑构建 |
| **Do** | 文字一律画在**页**上（用 `slides.py` 里的 `tb()` / `K.textbox()`） |
| **Do** | 绿色进度段画在**版式**上，金色子进度条画在**页**上（见 §4） |
| **Do** | 改完**一定**渲染成 PNG 并**真的看图**（§7） |
| **Do** | 公式先渲染成 PNG 再插入（§5.3），不要用 PowerPoint 公式编辑器 |
| **Do** | 验证脚本 import `build/_com.py`，不要自己 `Dispatch()` |
| **Do** | 色值用 `K.c(tok, "primary")` 这类语义 key 取 |
| **Do** | 新建正文用 `slides.py` 的现成辅助函数：`tb` / `bullets` / `add`，插图用 `img` + `cap`，表格用 `tbl`，公式用 `eqimg` / `eqblock` / `eqgrid`，算法用 `alg`（§5） |
| **Do** | 公式**必须**与同页其他公式共用一个倍率 `EQ_K` —— 见 §5.3.1，这是返工最多的地方 |

### Do Not

| | |
|---|---|
| **Don't** | ❌ **绝不 `taskkill /F /IM POWERPNT.EXE`** —— 会关掉用户正开着的所有 PowerPoint 窗口，且不留崩溃痕迹，表现为"莫名闪退"。用 `_com.py` |
| **Don't** | ❌ 不要把文字画在**版式**上 —— 版式文字在普通视图里点不中、改不了，用户会卡住 |
| **Don't** | ❌ 不要手抄 hex 到代码里 —— 一定从 `tokens.json` 派生 |
| **Don't** | ❌ 不要在手工 `SlideLayoutPart.load()` 出来的 part 上调用 `get_or_add_image_part()` —— 会产生重名 media 部件并损坏文件（用 `preload_images`） |
| **Don't** | ❌ 不要把新内容直接写进 `out/*.pptx` —— 产物是生成的，改了会被下次构建覆盖 |
| **Don't** | ❌ 不要用 `<a:rot>` 子元素做旋转（会被静默忽略，形状保持轴对齐） |
| **Don't** | ❌ 不要用 `sldLayoutId id="2147483648"`（PowerPoint 会拒绝打开整个文件） |
| **Don't** | ❌ 不要凭 `grep` 字面量判断能力是否存在 —— 找能力，不找名字 |
| **Don't** | ❌ 不要在没有渲染验证的情况下宣称改好了 |

---

## 2. 标准流程（按顺序做）

### 第 0 步 — 自我定位

```bash
cd build
python _pptopen.py ../out/seu_nari_template.pptx      # 产物现在能被打开吗
python _render.py  ../out/seu_nari_template.pptx ../_preview/deck
python _contact.py ../_preview/deck ../_preview/sheet.png 3 560
```

看联系表，确认你理解当前状态。**不要跳过这一步**，否则你不知道自己的改动是否破坏了什么。

### 第 1 步 — 改文案

汇报的标题、汇报人、导师、各章节名都在 `build/build_pptx.py` 的 `DECK` 字典里：

```python
DECK = {
    "author_inst": "张珂 · 东南大学电气工程学院",     # 页脚署名
    "cover_kicker": "研究生第一学年研究进展汇报",
    "cover_title": "面向时序预测的轻量化模型研究",
    "cover_en": "Lightweight Modeling for Time-Series Forecasting",
    "cover_meta": [("汇报人", "张珂"), ("导师", "×××　教授"), ...],
    "end_title": "感谢聆听", "end_sub": "敬请批评指正", "end_date": "2026 年 9 月",
    "sections":    ["研究背景", "相关工作", ...],     # 章节中文名（数量随意，见下）
    "sections_en": ["BACKGROUND", "RELATED WORK", ...],  # 英文名，两条列表等长
}
```

**模板代码里不写死任何文案** —— 封面、致谢、页眉页脚、目录条目全部从这里取。

改 `sections` / `sections_en` 就够了 —— 目录条目、章节过渡页、内容页眉标三处都会跟着变。
**章节数由列表长度决定**（`NSEC = len(DECK["sections"])`），版式、进度段数、刻度全自动跟随，
不需要改任何常量。

### 第 2 步 — 改内容（加/删/改页）

见 §5 的任务手册。新页加进 `slides.py` 的 `SLIDE_PLAN`，每行是：

```python
(版式名, 标题占位符文字, 眉标章节号 或 None, 正文函数 或 None)
```

### 第 3 步 — 构建

```bash
cd build
python prep_figures.py      # 仅当换了插图：从源目录收集并统一命名 → assets/figures/
python extract_math.py      # 仅当公式来自 LaTeX 源：抽取公式块 → _math_spec.json
python prep_refs.py         # 仅当文献来自源 PDF：抽取条目 → _refs.json
python make_formulas.py     # 渲染公式/伪代码 PNG（含内联 FORMULAS / ALGOS）
python build_pptx.py        # 全量：主题 + 母版 + 全部版式 + 全部内容页
```

后三步各自可跳过：没跑过取材脚本时，`make_formulas.py` 只渲染内联公式，
`slides.py` 用内联的文献列表。

> `.potx` 已移出流程：实测各 agent 都直接用 python 产出 `.pptx`，没人用
> 双击新建的模板文件。真需要时用 PowerPoint 的 文件 → 另存为 → PowerPoint 模板，
> 或从同级 `PPT模板/` 项目取 `make_potx.py`。

分阶段调试用（出问题时能立刻定位是哪一层）：

```bash
python build_pptx.py --phase A   # 只主题+母版
python build_pptx.py --phase B   # +全部版式（封面/目录/过渡 NN/内容 NN/致谢）
python build_pptx.py             # = ABC
```

### 第 4 步 — 验证（**必做，不可省**）

```bash
cd build
python _pptopen.py ../out/seu_nari_template.pptx
python _render.py  ../out/seu_nari_template.pptx ../_preview/deck
python _contact.py ../_preview/deck ../_preview/sheet.png 3 560
python ../design/check_tokens.py            # 只在改过 tokens.json 时需要
```

然后**打开联系表逐项看**（§7 有检查清单）。自动化只能告诉你"文件合法"，**不能告诉你"长得对"**。

---

## 3. 分层架构（理解这个才不会改错地方）

| 层 | 承载 | 为什么在这一层 |
|---|---|---|
| **母版** | 白底、页脚细线、页脚署名、页码域 | 所有页共享，改一次全改 |
| **版式** | 进度条**绿色段**、**东大校徽（仅 `内容 NN`）**、章节刻度、标题占位符、金色标题细线、几何菱形 | 每章一套；填充是版式属性 → 章内插页不会错乱 |
| **页** | **所有文字** + 正文内容 + **金色子进度条** | 作者日常编辑的地方 |

**硬规则**：版式上**只放图形与占位符，不放文字**。理由见 Do Not。

**版式数量随章节数变化**，共 `3 + 2 × NSEC` 个（封面 + 目录 + 每章两个 + 致谢，
`NSEC = len(DECK["sections"])`）：

```
封面 · 目录 · 章节过渡 01…NN · 内容 01…NN · 致谢
```

随附示例是 5 章 → 13 个版式。

命名规律：**`内容 NN` = 第 NN 章**。套用它，进度条前 NN 段自动变绿。

`python-pptx` 1.0.2 的能力边界（决定了为什么必须直接写 OOXML）：
- `SlideLayouts` **没有** `add_slide_layout()`
- `LayoutShapes` / `MasterShapes` **没有** `add_shape()` / `add_picture()` / `add_textbox()`
- 无法新增母版

可用接口：`element` 可 lxml 操作、`SlideLayoutPart.load(partname, content_type, package, blob)`、`part.relate_to(...)`、`preload_images()`。

---

## 4. 进度条（两级，务必理解再改）

| 级别 | 外观 | 放在哪 | 含义 |
|---|---|---|---|
| **章节级** | 前 NN 段填绿（总段数 = 章节数 `NSEC`） | **版式** | 你在第几章 |
| **章内页级** | 当前段下方一条金线，长 `48px × k/N` | **页** | 你在本章第几页 |

`k` = 本章第几页，`N` = 本章总页数。例：某章 5 页 → 金线 9.6 / 19.2 / 28.8 / 38.4 / 48 px。

**代码位置**：
- 绿段：`build_pptx.py` 的 `bar_sp(tok, k_filled)`（版式级）
- 金线：`slides.py` 的 `gold_sub_bar(slide, tok, section_no, k, n)`（页级），`k/N` 由 `build_slides()` 按 `SLIDE_PLAN` 自动分组算出

**几何常量**（`build_pptx.py` 顶部）：`BAR_TOP=93.5`、`BAR_W=48`、`BAR_H=9`、`BAR_PITCH=55`、`M=53`。

**维护成本（必须告知用户）**：绿段插页不会坏；**金线插页需要手动拖右边缘**，因为长度取决于"本章共几页"。若用户不想要这个维护，把每页金线都拖满 48px，即退化成纯章节级、零维护。

**改段数/加章节**：

1. 只改 `DECK["sections"]` / `["sections_en"]` —— 段数由 `NSEC = len(DECK["sections"])`
   派生，绿色进度段、章节刻度、以及 `章节过渡 NN` / `内容 NN` 两个版式循环全部跟着走。
2. 内容里引用 `内容 0N` 版式的地方随之增删。

> 早期版本把段数硬编码在 4 处 `range(1, 7)` 里（`bar_sp`、`marks_sp`、`build_layouts` ×2），
> 漏改一处会**静默少画一段绿条**或直接 `KeyError`。已改为从 `DECK` 派生。

---

## 5. 任务手册（Cookbook）

所有坐标单位是 **px**（1280×720 画布，1px = 1/96 英寸，96px = 1 英寸）。用 `K.px()` 转 EMU。

### 5.1 布局常量速查

`build_pptx.py`：

```python
M           = 53      # 左右页边距
BODY_W      = 1174    # 正文区宽（= 1280 - 2*53）
BAR_TOP     = 93.5    # 进度条顶（与 52px 页眉行垂直居中）
BAR_W       = 48      # 进度段宽
BAR_H       = 9       # 进度段高
BAR_PITCH   = 55      # 进度段间距（BAR_W + 7 gap）
SEAL_CM     = 2.91    # 校徽边长（厘米，用户指定）
SEAL_PX     = round(K.cm(SEAL_CM))   # = 110 px，内容页右上角
FOOT_RULE_Y = 660     # 页脚细线
TITLE_Y     = 166     # 标题占位符顶（高 50）
RULE_Y      = 228     # 金色标题细线（宽 192，高 2）
CONTENT_Y   = 254     # 正文区顶部（slides.py 也有一份，同值）
BODY_H      = FOOT_RULE_Y - CONTENT_Y     # = 406px ★ 每页的排版预算
```

> **`BODY_H = 406px` 是支配性常量。** 每一页的排版预算都是它：
> 正文区 1174 × 406，任何内容都必须在 `CONTENT_Y=254` 与 `FOOT_RULE_Y=660`
> 之间排完。放不下时的退让顺序见 §5.8。

`slides.py`：

```python
EYEBROW_Y   = 144     # 章节眉标
CONTENT_Y   = 254     # 正文区顶部
FOOT_RULE_Y = 660     # 页脚细线
BODY_H      = FOOT_RULE_Y - CONTENT_Y     # = 406px ★ 每页的排版预算
BODY_W      = 1174    # 与上面同值
COL_W       = 572.5   # 双栏栏宽 (1174-29)/2
COL2_X      = 654.5   # 右栏左边界 53 + 572.5 + 29
CARD_W      = 378     # 三卡卡宽 (1174-40)/3
EQ_K        = 0.40    # ★ deck 标准公式倍率（≈16pt 数学字号）—— 见 §5.3.1
ALG_K       = EQ_K    # 算法清单的同一上限：算法与公式出自同一条渲染管线
NO_STYLE    = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"   # "No Style, No Grid"
```

> 公式放置一律走 `EQ_K` 倍率制（不是把每个框各自缩放到最大），理由见 §5.3.1。

字体面（`LATIN` / `EA`）定义在 `build_pptx.py`，**从 `tokens.json` 的 `type.latin-stack[0]` / `type.cjk-stack[0]` 派生**；`slides.py` 直接 `from build_pptx import LATIN, EA`，不重复定义。想换字体改 `tokens.json` 即可。

### 5.2 加一页内容

在 `build/slides.py` 加一个 `c_*` 函数，再注册进 `SLIDE_PLAN`：

```python
def c_my_page(slide, tok):
    tb(slide, tok, M, CONTENT_Y, BODY_W, 40,
       [{"runs": [("正文第一段。", K.c(tok, "ink"), False)],
         "size": 16, "latin": LATIN, "ea": EA, "line_spacing": 1.55}],
       name="正文")

SLIDE_PLAN = [
    ...
    ("内容 02", "我的新页标题", 2, c_my_page),   # 第 2 章
]
```

**要点**：
- 版式名用 `内容 NN`，`NN` 决定进度条绿到第几段
- 第三个字段（章节号）用来生成"NN · 章节名"眉标；不需要就填 `None`
- 标题**不要**自己画 —— 靠 `SLIDE_PLAN` 的第二个字段写进标题占位符
- 标题占位符是**空的**（PowerPoint 正确行为），所以**一定要填第二个字段**，否则页面看着像没写完

**段落字典支持的 key**：

```python
{"runs": [(文字, "RRGGBB", 是否粗体[, 上下标]), ...],   # 必填；可多个 run 混排
 "size": 16,                                  # pt
 "latin": LATIN, "ea": EA,                    # 西文 / 中日韩字体，两个都要给
 "align": "l" | "ctr" | "r",
 "line_spacing": 1.55,                        # 倍数
 "spc": 180,                                  # 字距，1/100 pt
 "spc_before": 10, "spc_after": 8,            # 段前/段后，pt
 "marL": 171450, "indent": -171450,           # 悬挂缩进（EMU）
 "bullet": K.c(tok, "primary")}               # 绿色方块项目符号；配 marL/indent 用
```

`runs` 的每一项还可以是 **4 元组**，第 4 项是 DrawingML 的 `baseline`
（上标 `30000`、下标 `-25000`）。

⚠️ **`MT()` 的输入是记号语法，不是纯文本。** 它把每个字符串交给 `mr()`，而
`mr()` 会消费**所有**裸 `_` / `^`：

| 你想写 | 该用 | 为什么 |
|---|---|---|
| 普通正文、文件名、URL、代码 | `R(tok, [...])` | 不做任何记号解析 |
| 记号，如 `UE_a^{s,e}(t)` | `MT(tok, [...])` | 自动拆成上下标 run |
| 记号里的**字面**下划线 | `MT` 里写 `my\_file` | `\_` / `\^` / `\\` 是转义 |

用错的下场：`MT(tok, [("train_loss.png", ...)])` 会渲染成 `train`＋下标 `l`＋`oss.png`
—— 不报错，只是悄悄变了形。

记号留在页上做文字（而非图片）是为了可编辑、可搜索；只有公式本体才渲染成图片（§5.3）。

### 5.2b 从既有材料取材（三个可选脚本）

做一份汇报，手边往往已经有一堆现成材料。三个脚本把常见形态转成管线能吃的中间件。
**都是只读源、不改源，也都可不用** —— 不用就留空目录，构建照样跑。

| 脚本 | 输入（脚本顶部常量） | 输出 | 适用 |
|---|---|---|---|
| `prep_figures.py` | `SOURCE_DIR`：一个插图目录 | `assets/figures/` + manifest | 论文的图、参考 PPT 导出的图；矢量 PDF 会被光栅化 |
| `extract_math.py` | `SOURCE_TEX`：一份 LaTeX 源 | `_math_spec.json` | 论文 / beamer 讲义里已有的公式与伪代码 |
| `prep_refs.py` | `SOURCE_PDF` + 页码范围 | `_refs.json` | 已排好版的参考文献，可复用其著录格式 |

**用法**：把脚本顶部的源路径常量改指到你的文件，跑一次；产物进版本库后就不必再跑。

**它们不试图理解内容** —— 只做搬运与规范化。所以源的组织方式一变，脚本可能要跟着调
（`prep_refs.py` 要页码范围，`extract_math.py` 依赖 LaTeX 环境名）。

**没有 LaTeX 源也行**：公式直接在 `make_formulas.py` 的 `FORMULAS` 里手写 LaTeX，
数学写法与源无关，走的还是同一条渲染管线。

**从 PDF 抽文本的经验**：正文里 `ﬀ/ﬁ/ﬂ` 这类连字要还原成 `ff/fi/fl`；
页眉页脚与页码会混进正文，得按行模式剔掉（`prep_refs.py` 里有示例正则）。

### 5.3 插公式

**不要用 PowerPoint 公式编辑器**（字重/间距/字体都与本设计不搭）。走 LaTeX 渲染成透明 PNG：

1. **批量（推荐，当公式本来就在 LaTeX 里）**：把 `extract_math.py` 顶部的源路径
   `SOURCE_TEX` 指向你的 `.tex`，它会把每个 frame 的 `equation` / `align` /
   `gather` / `algorithmic` 抽出来写成 `_math_spec.json`，`make_formulas.py`
   直接吃这份清单 —— **不手抄**，几十个公式不会抄错。
   源可以是论文的 `.tex`、beamer 讲义，或任何能产出公式文本的地方。

2. **临时加一条**：在 `build/make_formulas.py` 的 `FORMULAS` 里加：

```python
FORMULAS = [
    ("quantile_loss", r"L_{\tau}(y,\hat{y})=\max\left\{\tau\,(y-\hat{y})\right\}", "(3)"),
    ("my_formula",    r"\frac{\partial L}{\partial \theta}=0", "(5)"),   # ← 新增
]
```

3. 在 `slides.py` 里放置（`EQ_K` 是 deck 标准倍率）：

```python
eqimg(slide, tok, "my_formula", x, y, mw, mh)      # 单张，等比缩放到 mw×mh 内
eqblock(slide, tok, ("k1", "k2"), x, y, w, mh)     # 浅绿底 + 左侧绿条，多式堆叠
eqgrid(slide, tok, keys, x, y, w, mh, cols=2)      # 公式墙分栏，统一倍率
```

4. 构建顺序：**先** `python build/make_formulas.py`，**再** `python build/build_pptx.py`。
   放置尺寸由 `manifest.json` 的像素宽高推算，所以 manifest 必须先重生成。

**管线**（了解即可）：

```
pdflatex(standalone, border=2pt)  →  PDF
pdftocairo -png -transp -singlefile -r 400  →  透明 PNG   （400 = DPI 常量）
Pillow alpha bbox  →  裁掉透明边距
```

**要点**：
- 字体在 `make_formulas.py` 的 `_BODY` 里：默认 `\usepackage{lmodern}`
  = **Latin Modern**（LaTeX 默认数学字体）。**想改回 Palatino 数学**就把那一行换成
  `\usepackage{mathpazo}`——只影响公式图片，不动正文
- 模板已加载 `ctex`（`fontset=windows`），因此**公式里可以直接写中文**
  （如 `\overset{\text{强对偶}}{\Longrightarrow}`），伪代码也是中文关键字
- ⚠️ `standalone` **不能把 `align` / `gather` 直接放在顶层**，会报
  `Missing \endgroup inserted`。要改写成 `aligned` / `gathered`，再包进
  `$\displaystyle ...$`（`extract_math.py` 自动做这件事）；`cases` 无此问题
- 算法清单走 `algorithm2e` 的三线式（不用 `algpseudocode` 的裸列表），`varwidth=16cm` —— 细节与切分阈值见 §5.3.2
- 颜色写死 `#1A1A1A`；**改 `tokens.json` 的 `ink` 后要同步 `make_formulas.py` 的 `INK`**（颜色是烤进像素的，不会自动更新）
- 缩放按**统一倍率** `scale = (EQ_K / base_pt) × (96 / dpi)`，不是统一高度 —— 这样公式间字号比例正确，带 `\sum` 上下限的自然更高
- 公式**编号留在页上做文字**（右对齐），所以仍可编辑；只有公式本体是图片
- 公式块样式：`surface-alt` 底 + 左侧 3px `primary` 竖条
- ⚠️ `-singlefile` 不能省，否则 `pdftocairo` 会给文件加 `-1` 后缀
- ⚠️ LaTeX 模板里**不能用 `%`-格式化**（`%` 是 LaTeX 注释符）；模板用 `@@INK@@` / `@@BODY@@` + `.replace()`

#### 5.3.1 公式大小必须一致 —— 返工最多的地方

`make_formulas.py` 用**同一个 DPI、同一个 10pt 基准**渲染所有公式，所以
**一张公式 PNG 里 1 像素 = 固定的物理尺寸**。由此只有一条规则：

> **同一份 deck 里，所有公式按同一个倍率 k 放置。**

很容易做错的反面做法是「把每个公式缩放到填满它自己的框」：同一页上短公式被放大得巨大、
长公式被压得极小，看起来像几套不同的公式。这不是审美偏好，是**看得见的排版错误**。

`slides.py` 就是按这条规则写的：

| 常量 / 函数 | 语义 |
|---|---|
| `EQ_K = 0.40` | **deck 标准倍率**（≈16pt 数学字号）。公式默认用它 |
| `eqimg(key, …, k=None)` | 单张：`min(EQ_K, 框放得下)`，**只在框太小时才缩小，绝不放大** |
| `eqplace(slots)` | 同一页里形状不同的框（宽目标函数 + 窄约束组）**共用一个 k** |
| `eqblock(keys, …)` | 浅绿块内多条公式共用一个 k；装不下就**整体**缩小 |
| `eqgrid(keys, cols=2)` | 公式墙分栏，所有列共用一个 k |
| `eqk(keys, w, …)` | 预先算出 `eqgrid` 会用的 k，好让同页其它公式与之对齐 |
| `alg(key, …, ALG_K)` | 算法清单：上限 `ALG_K = EQ_K`，短清单**不会**被放大到填满栏宽 |

**加公式时的自检**：如果新公式让某页的 k 掉到 `EQ_K` 以下，说明这一页**内容太多**——
把它挪到新页，或者把框开大。**不要只把这一张缩小**，那正是尺寸不一致的来源。

确实放不下的页（整页算法、公式墙）允许整页变小，但必须**整页统一**。

`_audit.py` 会把每页的倍率分布打出来。**每一页只能有一个倍率**；除 `EQ_K` 之外每多出
一个值，都要能说出理由（说不出就是有公式被单独缩放了）。示例 deck 实测：
`0.400 × 2`（公式块、公式墙）+ `0.293 × 1`（整页算法，受正文高度限制整页缩小）。

#### 5.3.2 算法清单：论文三线式

算法**不用** `algpseudocode` 的裸列表，而是 `algorithm2e` 的 ruled 三线式 ——
上下横线 + 自动生成的 `Algorithm N: 名称` + 行号 + 粗体关键字，和论文里一样。
字体会跟着 §5.3 的字体设置走（默认 Latin Modern）。

```python
alg(slide, tok, "demo_pseudo", M, CONTENT_Y, BODY_W, 362)
```

`make_formulas.py` 的 `ALGOS` 每条是 `(key, caption, body)`：`caption` 就是
`Algorithm N:` 后面那句；`body` 用 algorithm2e 语法（`\KwIn` / `\KwOut` /
`\While{cond}{...}` / `\If{cond}{...}` / `\Return`，行尾用 `\;`），
中文关键字可直接写。

**切分只作为最后手段。** 清单作为单张图放置时受正文高度限制，太长就会被缩得很小。
`make_formulas.py` 为此额外导出左右两段（`<key>_a` / `<key>_b`），`alg()` 只在
单张倍率低于 `ALG_MIN_K` 时才改用两段并排：

```python
ALG_MIN_K = 0.16      # 极低阈值：绝大多数清单不切，全屏看很清楚
```

**别轻易调高这个阈值** —— 十几行的算法单张放完全够看，切开会打断阅读顺序。

### 5.4 插图片 / 图表

**真图片**（照片、截图、外部图表）：

```python
img(slide, tok, "nc_network", x, y, mw, mh)     # 等比缩放到 mw×mh 内并居中
cap(slide, tok, x, y, w, "NC-Road 网络拓扑")     # 下方"图 N"绿色前缀图注
```

`img()` 从 `assets/figures/manifest.json` 读像素宽高、**自己算等比**，所以不会拉伸变形。
`K.picture()` 本身**不做等比约束**，绕过 `img()` 直接用它就必须自己算好宽高。

**插图管线**（`build/prep_figures.py`，可选）：

- 把 `SOURCE_DIR` 指向你的插图目录，位图按 ASCII key 复制进 `assets/figures/`
  （避免中文文件名与 `*.png.jpeg` 这类双扩展名）
- **矢量 PDF 用 `pdftocairo -png -r 400` 光栅化** —— python-pptx 不能嵌入 PDF
- 写 `manifest.json` 记下每张图的像素宽高，供 `img()` 等比计算

源可以是论文的图目录、别人给的 PPT 导出的图片、你自己的画图脚本输出。

> 需要"图还没画好"的占位框时，直接用 `K.rect(..., dash='dash')` 画一个——
> 先放一张真图再替换通常比占位框更省事。
>
> 规则：**文档里不出现代码中没有的符号。** 改代码删了函数，顺手搜一遍 SKILL.md。

**图表配色**：用 `tokens.json` 的 `data-series` 色序
`#4D7C2B · #00807C · #C9A227 · #7BA05B · #66B2B2 · #8C6D1F · #9AA09A`
—— 生成图表（matplotlib 等）时显式设成这串，避免默认彩色与绿金主调打架。

### 5.5 插表格

**用 python-pptx 的真表格对象**（可编辑，不是图形拼的）。`slides.py` 的 `tbl()` 已经把它包好：

```python
tbl(slide, tok, heads, rows, widths, y=CONTENT_Y, size=11.5, row_h=32,
    title="表 2　…", note="注：…",
    left=(0,), ctr=(1, 4),     # 左对齐 / 居中的列；其余（数字列）一律右对齐
    emph=(2,))                 # 第 2 行整行强调（primary-dark 粗体）
```

- `widths` 是各列占 `BODY_W` 的比例，和为 1
- `row_h` 是**最小**行高 —— PowerPoint 会为折行内容自动撑高，所以文字长的表要
  给足，否则表格会撑过页脚（`FOOT_RULE_Y = 660`）
- 想让**单个**单元格强调，给它文字加 `"*"` 前缀即可
- 返回表格 + 表注下方的 y，方便接着往下排

底层实现（`tbl()` 内部）：

```python
gf = slide.shapes.add_table(rows, cols, K.px(M), K.px(top), K.px(BODY_W), K.px(30*rows))
table = gf.table
_no_style(table)                       # ① 关掉内置表格样式，才能自己画 booktabs 风格
for i, frac in enumerate(widths):      # ② 列宽按比例
    table.columns[i].width = Emu(int(BODY_W * frac * K.PXE))
_cell(tok, table.cell(0, j), text, size=13.5, color="FFFFFF", bold=True,
      fill=K.c(tok, "primary"), align="r")          # ③ 表头：绿底白字
_cell(tok, table.cell(i, j), val, size=14, color=K.c(tok, "ink"),
      fill=fill, align="r", bottom_rule=K.c(tok, "rule"))   # ④ 表体：细底边
```

**表格样式规范**：

| 元素 | 规格 |
|---|---|
| 表头 | 底 `primary`(#4D7C2B)，文字白、13.5pt 粗体 |
| 表体 | 隔行 `#FFFFFF` / `surface-alt`(#F5F7F1) |
| 分隔线 | **只画底边**（`rule` #E4E4DE，0.75pt），无竖线 |
| 对齐 | 首列左对齐，数字列右对齐 |
| 数字字体 | `Palatino Linotype`（西文/数字），CJK 用 `Noto Sans SC` |
| 强调行 | 该行文字用 `primary-dark`(#3E6623) + 粗体（`emph=(i,)`） |
| 强调单元格 | 文字前加 `"*"`，渲染时去掉星号只留强调 |
| 表题 | 在表格**上方**，11pt `ink-muted`，右下标形式"表 2 …" |
| 表注 | 在表格**下方**，11pt `ink-muted` |

**为什么 `_no_style` 必须调**：PowerPoint 内置表格样式会带来竖线、粗边框、自动 banding，与设计冲突。`slides.py` 里的 `NO_STYLE` GUID 是"No Style, No Grid"。

**CJK 字体必须手动补**：python-pptx 的 `run.font.name` 只设西文。`_cell()` 里额外插了 `<a:ea typeface="Noto Sans SC"/>`。**你自己写单元格代码时别忘了这一步**，否则中文会用主题默认字体。

**多级表头**（源材料里的 `\multirow` + `\cmidrule` 很常见）：`tbl()` 只做单行表头。
需要两级表头时落到 python-pptx 原生合并 —— `_no_style()` / `_cell()` 照样能用：

```python
table.cell(0, 1).merge(table.cell(0, 3))     # 第一行跨 3 列
_cell(tok, table.cell(0, 1), "误差指标", size=11, color="FFFFFF",
      bold=True, fill=K.c(tok, "primary"), align="c")
# 被合并掉的单元格不用再填；第二行照常写子表头
```

**`tbl()` 返回的 y 只是名义值。** 它按 `row_h × 行数` 算，而 `row_h` 是**最小**行高 ——
PowerPoint 会按内容撑高。密集表格后面接着排内容时会被撑高的表压住，
所以要**留余量**（多留 20–40px），并渲染后逐页看图确认。

### 5.6 改配色 / 字体

1. 改 `design/tokens.json`
2. `python design/check_tokens.py` —— **必跑**，它会复算 WCAG 对比度并与文件里声明的值比对
3. 若改了 `ink` → 同步 `build/make_formulas.py` 的 `INK`，重跑 `make_formulas.py`
4. `python build/build_pptx.py`
5. 渲染验证

**对比度约束**（实测值，写在 tokens.json 的 `contrast` 块里）：

| 组合 | 实测 | 结论 |
|---|---|---|
| 白字 / `primary` | 4.96:1 | AA —— 可用，但投影环境建议白字只用于 ≥20pt 标题 |
| 白字 / `secondary` | 4.80:1 | AA，同上 |
| ink / surface | 17.4:1 | AAA |
| ink / surface-card | 15.0:1 | AAA |
| `primary-dark` / surface-card | 5.8:1 | AA |
| `accent`(金) / surface | 2.4:1 | **只作装饰，绝不作文字** |
| `ink-muted` / surface | 5.3:1 | AA |

### 5.7 改 logo

| 位置 | 文件 | 尺寸调节 |
|---|---|---|
| 内容页右上校徽 | `assets/seu-seal.png`（578×578） | `build_pptx.py` 的 `SEAL_CM`（当前 2.91 cm → 110 px） |
| 封面大组合·东大 | `assets/seu-logo.png` | `layout_cover()` 里的 `80`（高，px） |
| 封面大组合·南瑞 | `assets/nari-logo.png` | `layout_cover()` 里的 `50`（高，px） |

**校徽只出现在 `内容 NN` 版式上**（每章一个；封面/目录/过渡/致谢都没有）。改动位置在 `build_pptx.py` 的 `layout_content()` 调用 `seal_sp()`；若要去掉，删掉那一项即可。

---

### 5.8 放不下时的退让顺序（内容密度 vs 版面）

正文区只有 1174 × **406px**（`BODY_H`）。放不下时**按固定顺序退让**，不要各自发明：

| 顺序 | 动作 | 说明 |
|---|---|---|
| 1 | **减字号** | 沿阶梯往下走：`16 → 14 → 13 → 11.5 → 10`。下限 **10pt**（投影可读性） |
| 2 | **分两栏** | 单栏长块改双栏，同样的字能装下约 2 倍内容 |
| 3 | **删次要内容** | 合并同类项、把解释性文字移进讲稿 |
| 4 | **拆页** | 上面都试过还不够，就拆成两页——**优先拆页而不是继续缩字号** |

**字号下限是硬约束**：正文 < 10pt、表格 < 10pt 在投影上就不可读，
`_audit.py` 对表格报 problem。到了下限就只能走 3 或 4。

**整页缩小是允许的，但必须整页统一。** 公式墙/整页算法这类"一张图占满正文"
的页可以整体缩到 `EQ_K` 以下（见 §5.3.1），但页内不能参差。

## 6. 样式规范（照抄，别即兴）

### 字体

| 用途 | 字体 |
|---|---|
| 标题、正文中文 | **Noto Sans SC**（思源黑体） |
| 西文 / 数字 | **Palatino Linotype** |

两个都要显式设置：`"latin": LATIN, "ea": EA`。**只设一个会让另一类字符掉回默认字体。**

⚠️ 字体面写在 `tokens.json` 的 `type.latin-stack` / `type.cjk-stack` 里，代码从那里派生。**不要把 `latin-stack[0]` 改成 `TeX Gyre Pagella`** —— 那个字体只随 TeX Live 分发、Windows 未注册，PowerPoint 会**静默**掉回默认衬线体。pptx 用 `Palatino Linotype`（Windows 自带，与 Pagella 同属 Palatino 族）。

### 字号阶梯（pt）

| 角色 | pt |
|---|---|
| 封面主标题 | 40（模板默认 45；长标题用 40 才能一行放下） |
| 封面副标题（西文） | 22 |
| 封面眉标 / 元信息 | 16 / 16.5 |
| 章节过渡大数字 | 72 |
| 章节过渡标题 / 英文 | 36 / 15 |
| 页面标题（占位符） | 28 |
| 章节眉标 | 13 |
| 正文 | 16 |
| 卡片标题 / 卡片正文 | 16 / 13 |
| 图注 / 注释 | 11 |
| 关键数字 | 56 |
| 表格表头 / 表体 | 13.5 / 14 |
| 页脚 / 页码 | 11 |

### 颜色语义（取自 `tokens.json`）

| key | hex | 用途 |
|---|---|---|
| `surface` | `#FFFFFF` | 页面底色 |
| `surface-alt` | `#F5F7F1` | 表格隔行、公式底、注释区 |
| `surface-card` | `#EAF0E3` | 卡片底 |
| `ink` | `#1A1A1A` | 正文 |
| `ink-muted` | `#6B6B6B` | 注释、页脚、图注 |
| `ink-faint` | `#9AA09A` | 占位文字 |
| `rule` | `#E4E4DE` | 细线 0.75pt |
| `rule-strong` | `#CFD3C8` | 占位框虚线 |
| **`primary`** | `#4D7C2B` | **东大绿**：绿段、眉标、卡片顶条、表头 |
| `primary-dark` | `#3E6623` | 绿底上的正文色 / 正文强调 |
| `primary-light` | `#7BA05B` | 次要绿 |
| **`secondary`** | `#00807C` | **南瑞墨绿**：次要强调、副色卡片 |
| `secondary-dark` | `#00615E` | 辅色正文强调 |
| **`accent`** | `#C9A227` | **金**：仅细线/装饰，**绝不作文字** |
| `alert` | `#B23A2E` | 警示（"待解决"卡） |

上表只列**语义色**。`tokens.json` 里还有一组底色族的色：
`primary-tint` / `secondary-tint`（面板底）、`accent-soft`（浅金装饰）、
`white`（反白文字），以及 `data-series`（图表分类色序，是个列表不是单色）。
要改面板底色就改 tint 那几个，不要在代码里写 hex。

配色逻辑：东大绿为主、南瑞墨绿为辅、金只做细线 —— 两家校色同属绿系，可统一而不打架。金是**东大校徽真实的品牌色**（校徽金环 `#FDD100`），不是随意选的。

### 几何与线

- 外边距 `M = 53px`（0.55in）；正文区宽 `1174px`
- 细线 0.75pt（`9525` EMU）；金色标题线 2px 高、192px 宽
- 卡片顶部色条 3px；公式块左侧竖条 3px
- 封面/致谢的几何菱形：45° 旋转（`rot="2700000"`）

---

## 7. 验证：看什么

自动化只能证明"文件合法"。**必须看图。**

### 7.1 先跑机器审查

```bash
cd build
python _audit.py                     # 直接读 out/*.pptx，不需要 PowerPoint
```

`_audit.py` 查四类**只能靠代码发现**的问题，有 hard problem 时退出码为 1：

| 检查 | 判据 |
|---|---|
| **公式倍率漂移** | 同一页出现了两个不同的 k —— 就是 §5.3.1 那个错 |
| **出血** | 图片 / 表格 / 色块越过页脚线 `y=660` 或右边界 `x=1227`；文本框起点在页脚之下 |
| **看不清** | 表格字号 < 10pt；插图宽 < 180px（180–260px 给提示） |
| **空标题** | 内容页的标题占位符没填 |

它同时打印**公式倍率分布**。主倍率（`EQ_K`）之外的值都应当是有意为之的**整页**缩小，
每多出一个新值就要能说出理由 —— 说不出来就说明有公式被单独缩放了。

### 7.2 交付前逐页必查

机器查完，**再打开联系表逐页看**。下面四项是历次返工最多的地方：

**① 公式字号是否协调**（最容易被忽略）

- [ ] 同一页上所有公式**看起来一样大** —— 短公式没被放大、长公式没被压小
- [ ] 相邻页之间公式没有明显跳字号
- [ ] 公式墙 I–IV 翻页时字号完全不变
- [ ] 确实变小的页是**整页统一小**，不是页内参差

**② 图片大小与可读性**（对照 `错例/反例-IMG_SMALL-*`）

- [ ] 图片等比、无拉伸；南瑞 logo 无白边（透明底 PNG）
- [ ] 图注在图片正下方，未越过 `y=660`
- [ ] 图不是小到看不清 —— **优先删内容或换页，不要把它缩到 200px 宽**。
      正文区只有 1174 × 406px，宽高比大的图放不下就是放不下
- [ ] 图表（matplotlib 等生成的）分类色**取自 `tokens.json` 的 `data-series`**，
      不是默认色 —— 默认的蓝/橙/绿/红与绿金主调打架（对照 `错例/反例-CHART_COLOR-*`）

**③ 表格可读性**

- [ ] 表头 / 表体字号 ≥ 10pt（低于 10pt 投影基本看不清）
- [ ] 表格底边未越过 `y=660` —— `row_h` 只是**最小**行高，PowerPoint 会按内容撑高
- [ ] 列宽比例和为 1；首列左对齐、数字列右对齐；绿头白字、只画底边

**④ 版面硬指标**（对照 `错例/反例-NUM_MISSING-*`、`反例-OVERLAP-*`、`反例-BLEED-*`）

- [ ] 每个内容页**标题不为空**（最阴险的失败模式：不报错，只是看着像没写完）
- [ ] **编号齐全**：图 = 图 N、表 = 表 N、公式 = (N)、文献 = [N]（只能人工查）
- [ ] **没有元素重叠**：文字压图、数字压说明（只能人工查）
- [ ] 正文字号 ≥ 10pt —— `_audit.py` 只查**表内**字号，正文靠人眼
- [ ] **文本框**右边界 ≤ `x=1227`（`_audit.py` 只对图片查出血，文本框要人工查）
- [ ] 进度条绿段数 = 版式号（`内容 NN` → 前 NN 段绿）
- [ ] 金色子进度条长度 = `48 × k/N` px
- [ ] 校徽只在内容页；目录 / 过渡 / 致谢没有
- [ ] 中文标点与断行正常
- [ ] 改过 `tokens.json` → 跑过 `check_tokens.py` 且无 drift

`_contact.py` 的参数：`<src_dir> <out.png> [cols] [cell_w]`，例：`3 560` 得到 3 列、每格 560px。

### 7.3 多 agent 交叉审查（长 deck 用）

机器审查（§7.1）+ 主 agent 逐页看图（§7.2）之后，页数多的 deck（50 页以上）
还会漏 —— 单次审查的注意力是有限的。做法：

1. **分区间且有重叠**：例如 `0–20 / 10–30 / 20–40 / 30–50`。重叠 10 页用来交叉
   验证：两个 agent 都漏报的页才最可能真没问题。
   **按区间出图不要手工切 PNG** —— `_contact.py` 支持第五个参数：

   ```bash
   python _contact.py ../_preview/deck ../_preview/slice_10_30.png 3 560 10-30
   ```

   分区间出图这条经验在 §10 末尾也提过（单元格高度取自第一张图，别把两张不同比例的
   图拼在一起）。
2. 每个 subagent 拿到：**`错例/` 文件夹**、渲染出的页图、以及 §7.2 的检查清单。
3. 要求按 **页码 + 元素名 + 判据 + 期望值** 回报，不接受"不好看"这类描述。
4. 汇总时**优先修多个 agent 都报的**；只被一个报的先回看确认。
5. **发给 subagent 时区分两类结论**：`_audit.py` 的 **problem 必须修**（退出码非 0 就别
   往下走）；**advisory 是提示不是错误**（占位框底边、图偏小之类），发过去时说清哪类是哪类，
   否则 reviewer 会把噪声当缺陷报上来、也会漏掉真正该看的东西。

`错例/` 里是**反例截图 + 判据**（公式倍率不一致、图片太小、文字重叠、表格过小、
编号缺失），以及两张"改成什么样算对"的正例对照。先让 subagent 看反例，
它才知道要找什么 —— 比只给一段文字清单有效得多。

---

## 8. 坑清单（我实际踩过的，按危害排序）

前 5 条都是**"能写出来但 PowerPoint 拒绝打开/静默丢弃"**的静默雷。

### 8.1 `sldLayoutId` 必须 ≥ 2147483649 ⚠️ 最坑

用 `2147483648`（0x80000000）时 PowerPoint **拒绝打开整个文件**，报笼统的"发生意外"，看起来像包损坏但根本不是。

**症状特征：连一个空版式都打不开。** 若遇到"所有版本都失败"，优先怀疑这条。

`_pptx_kit.register_layout()` 的默认值已对，别改回去。

### 8.2 XML 结构必须严格包裹

| 错误 | 后果 |
|---|---|
| `<a:pPr>` / `<a:r>` 没包在 `<a:p>` 里 | PowerPoint **拒绝打开** |
| `<a:fld>` 嵌在 `<a:r>` 里 | 拒绝打开（`<a:r>` 只能含 `<a:rPr>` 和 `<a:t>`；`<a:fld>` 与 `<a:r>` 平级） |
| `<a:rPr>` + `<a:t>` 没包在 `<a:r>` 里 | **静默渲染为空**（我因此让所有页标题空白且不报错） |

### 8.3 旋转是属性，不是子元素

```xml
<!-- 对 -->
<a:xfrm rot="2700000"><a:off .../><a:ext .../></a:xfrm>
<!-- 错：<a:rot> 被静默忽略，形状保持轴对齐 -->
<a:xfrm><a:off .../><a:ext .../><a:rot val="2700000"/></a:xfrm>
```

`2700000` = 45°（1° = 60000）。我因此把封面和致谢的菱形全渲染成了方块。

### 8.4 图片必须在 package 自建的 part 上注册

```python
# 对：在母版（package 自建）上一次注册全部图片
image_parts = K.preload_images(master.part, IMAGE_FILES)

# 错：在手工 load() 出来的 layout part 上调用，会各建一份 ImageParts，
#     产生重名 media 部件（Duplicate name: ppt/media/image2.png）→ 文件损坏
part.get_or_add_image_part(path)
```

版式里的图片用 `K.picture_slot("key", ...)` 占位，`register_layout()` 随后换成真 rId。

### 8.5 `XmlPart.load` 的参数顺序

```python
SlideLayoutPart.load(partname, content_type, package, blob)   # package 在 blob 之前
```

传反了 Package 会被当 XML 解析，报 `ValueError: can only parse strings`。

### 8.6 PowerPoint COM 验证会被"进程残留"污染 —— 但**不要用 taskkill 解决**

`Presentations.Open()` 在**共用或尚未退净的 PowerPoint 实例**上会失败，报同一个泛化错误 `0x80070030`，**与文件无关**。根因很窄：上一个脚本刚 `Quit()` 时进程还在退出，新的 `Dispatch()` 挂到了那个濒死实例上。

我最初的对策是错的、而且会伤到用户：

```bash
taskkill /F /IM POWERPNT.EXE     # ❌ 关掉这台机器上所有人的 PowerPoint 窗口
```

用户看到的是"打开文件几秒后闪退"，而**事件日志里查不到任何 POWERPNT 崩溃记录** —— 因为强杀不留痕。这一点正是判定它不是真崩溃的依据（真崩溃会留 Application Error ID 1000 + WER 报告）。

**正确做法在 `build/_com.py`**：轮询等进程数稳定 + Open 失败退避重试，全程不杀进程。`force_kill()` 只作显式逃生口。

另外：**0 页的 pptx 用 COM 打不开**（stock 模板 0 页时同样失败），验证壳文件时先加一页。

### 8.7 其他

- `PP_PLACEHOLDER.TITLE` 的值是 **1**，不是 13。别硬编码，用枚举。
- **占位符从版式克隆到页时不带文字**（PowerPoint 的正确行为）。所以示例页要显式补标题，否则显示"单击此处添加标题"。
- **截图伪影**：用 `_contact.py` 或 Edge/Chrome 无头截超高页面时，会有未绘制区域（遇到过整块黑色）。看着像缺陷，其实是伪影 —— **先用小视口单独截一页确认**，别急着改代码。
- **无头截图视口太窄**会裁掉内容右边缘，看着像 CSS bug。布局需约 1270px 宽时用 1400px 视口。
- **grep 要找能力，不找名字**。曾 grep 字面量 `progressbar` 就断言"没有进度条"，实际 beamer 的进度条叫 miniframes。
- `pdftocairo` 不加 `-singlefile` 会给单页输出加 `-1` 后缀，文件名对不上。
- LaTeX 模板里别用 `%`-格式化（`%` 是注释符）。
- **`standalone` 里放 `align` / `gather` 要包一层 `aligned` / `gathered`**，
  否则报 `Missing \endgroup`；`algorithm2e` 用 `[H]` 不需要 float。
- **算法切分阈值别调高**：切分只是极长清单的兜底，正常长度切开会打断阅读（§5.3.2）。
- ⚠️ **`pdftocairo` 不能接受「含非 ASCII 的绝对输出路径」**。它把 argv 过一遍 ANSI
  代码页，中文/日文路径会被弄坏，报 `Error opening output file` —— 而**同样含中文的
  输入路径没事**，**相对输出名 + `cwd` 也没事**。项目目录叫 `ppt模板测试`，所以这不是
  假想问题。实测三种组合：绝对+ASCII ✓ / 绝对+中文 ✗ / 相对+cwd(中文) ✓。
  `make_formulas.py` 与 `prep_figures.py` 都已改成"相对名 + cwd"。
- **`out/*.pptx` 是确定性的**：`build_pptx.py` 保存后会重写一遍 zip、把条目时间戳钉死。
  python-pptx 默认给每个条目写当前时间，逐部件字节一致但**文件哈希每次都变** ——
  对跟踪这个二进制的仓库就是每次重建一个 1MB 的假 diff。
- **`tbl()` 的 `note=` 参数会遮蔽模块级的 `note()` 函数**。`slides.py` 里已加
  `_note_fn = note` 别名绕开；你自己在别的模块里写类似签名时注意这个陷阱。

---

## 9. 已知限制（如实告知用户，别假装没有）

1. **页码只有数字，没有"N / M"**。PowerPoint 没有"总页数"原生域。现状：自动编号；章节过渡页另印刻度。
2. **公式本体是图片**，不能在 PowerPoint 里改字形。要改公式得回 `make_formulas.py` 重渲染。编号是文字，可编辑。
3. **金线插页需手动调宽度**（见 §4）—— 用户可选退化为满宽、零维护。
4. **思源黑体 `NotoSansSC-VF.ttf` 是可变字体**，PowerPoint 对 VF 支持不可靠。当前设 `ea="Noto Sans SC"` 且只用 Regular/Bold 具名字重。**渲染看起来正常，但未逐字重严格验证**。若发现字重异常，把样式里的 `Noto Sans SC` 全局换成 `Microsoft YaHei`。
5. **南瑞 logo 是位图**（1088×206）非矢量。当前显示尺寸下约 395 dpi，够用；放很大建议描摹。
6. **东大校徽已换高清版**（`seu-seal.png` 578×578，2.91 cm 显示约 504 dpi，很清晰）。
   但封面用的完整 lockup `seu-logo.png` 仍是旧的 **330×128**（80px 高，约 154 dpi），
   放很大尺寸会软 —— 有更清晰的 lockup 就换掉它。两者都不是矢量。
7. **可随时删除的中间产物**（都会被下次构建重建，不影响源码）：
   `build/_shell.pptx`（A/B 阶段产物）、`build/_formula_work/`（LaTeX 编译临时件，约 5 MB）、
   `build/__pycache__/`、`_preview/`（渲染出来的 PNG）。交付/归档前清掉能省很多体积。
8. `assets/formulas/manifest.json` 里的像素尺寸随渲染 dpi 变化 —— 改了 `DPI`（渲染端）或 `EQ_K`（放置端）后要靠重新渲染/重新构建让它生效，别手改 manifest。
9. **公式密集的页字号会偏小**。整页十几行公式时，为了塞进 16:9 版面只能整页缩小；
   这是内容密度决定的，不是排版失误（改成文字反而更糟）。要现场细讲就拆成两页。
   判据见 §5.3.1：`_audit.py` 会打印每页倍率，新出现的倍率必须能说出理由。
10. **伪代码整页成图时宽高比偏方**。受正文区高度（406px）限制，接近 1:1 的图只能占版面
    约 45% 宽。要么拆页，要么把源 PDF 直接附上。
11. **参考文献可以是渲染文本，不必重排 `.bib`**。从一份**已排好版**的源（编译好的 PDF、
    别人的成品 PPT）里抽文本，能直接复用其著录格式，省掉重实现样式的工作。
    `prep_refs.py` 顺手修掉了 `ﬀ/ﬁ/ﬂ` 连字 —— 从 PDF 抽文本时这类连字很常见。
12. **表格 `row_h` 是"最小行高"**。PowerPoint 会按内容自动撑高，给得太小表格就会越过页脚
    （`FOOT_RULE_Y = 660`）。文字多的表务必渲染后逐页看图。
13. **图片朝向要与框的朝向匹配**。竖构图的图塞进宽扁框会被压得很窄（`_audit.py` 对
    < 180px 宽的图报 problem）。先看图的长宽比，再决定放栏内还是铺满整宽。

---

## 9b. 改了东西要重跑什么

| 改动对象 | 需重跑 | 依赖产物 |
|---|---|---|
| `design/tokens.json`（颜色/字号） | `build_pptx.py` | —— |
| `tokens.json` 的 **`ink`** | `make_formulas.py` → `build_pptx.py` | 公式 PNG 的**颜色烤进像素**，不随 tokens 走 |
| `tokens.json` 的字体栈 | `build_pptx.py` | 主题字体 |
| `make_formulas.py` 的 `FORMULAS` / `ALGOS` | `make_formulas.py` → `build_pptx.py` | `assets/formulas/manifest.json` |
| `make_formulas.py` 的 `DPI`**或** `slides.py` 的 `EQ_K` / `ALG_K` | 同上 | manifest 里的像素尺寸随之变 |
| 换了插图 | `prep_figures.py` → `build_pptx.py` | `assets/figures/manifest.json` |
| `slides.py` 的内容 | `build_pptx.py` | —— |
| `build_pptx.py` 的按版式装饰 | `build_pptx.py` | —— |
| 公式来自 LaTeX 源且源变了 | `extract_math.py` → `make_formulas.py` → `build_pptx.py` | `_math_spec.json` |
| 文献来自源 PDF 且源变了 | `prep_refs.py` → `build_pptx.py` | `_refs.json` |

**口诀**：改了"素材生成端"就先重跑那个脚本；`build_pptx.py` 永远是最后一步。

---

## 10. 快速参考：一条命令看全

> **必须先 `cd build`，且输出文件名写死在代码里。** 两个隐含约定：
> * `slides.py` 用的是 `from build_pptx import ...` 顶层导入，所以要从 `build/` 里跑；
>   在仓库根跑 `python build/build_pptx.py` 会 `ImportError`。
> * 输出路径是 `build_pptx.py` 的 `OUT`（`out/seu_nari_template.pptx`）。要换 deck 名，
>   改 `OUT` **并同步**本文档里的命令。

```bash
cd build
python make_formulas.py && python build_pptx.py \
  && python _audit.py \
  && python _pptopen.py ../out/seu_nari_template.pptx \
  && python _render.py ../out/seu_nari_template.pptx ../_preview/deck \
  && python _contact.py ../_preview/deck ../_preview/sheet.png 3 560 \
  && python ../design/check_tokens.py
```

**取材脚本（§5.2b）不在上面这条链里** —— 它们要读你的源文件，只在源变了的时候跑：

```bash
python prep_figures.py    # 换了插图
python extract_math.py    # 公式来自 LaTeX 源
python prep_refs.py       # 文献来自源 PDF
```

`_audit.py` 退出码非 0 就**不要交付** —— 先按 §5.3.1 和 §7 修。
它过了以后仍然要**看图**（§7.2）：机器只能证明"尺寸和位置对"，证明不了"看得清、协调"。

`_contact.py` 的单元格高度取自**第一张图**：张数多、比例不一时后面的图会被裁。
页数多的时候分几张看（把 PNG 分到几个子目录分别拼），别被裁掉的边误判成版面溢出。

然后打开 `_preview/sheet.png` 看图。
