# seu_nari_ppt_template
This is the ppt template, which is made for the first year course I will take in the SEU. Slide could be automatically made once material is provided to coding agent such as Codex, Claude Code and Deepseek Harness.

---

## 从这里开始

**`SKILL.md` 是完整文档** —— 面向操作这套模板的 coding agent，讲了分层架构、逐项任务手册
（加页 / 插图 / 插公式 / 插表格 / 算法）、样式规范、验证清单、以及一份"实际踩过的坑"。

人可以先看 `out/seu_nari_template.pptx` —— 那是模板自带的 18 页示例（5 章，管线支持的
每种能力各出现一次）。

## 仓库结构

```
SKILL.md                  ← 主文档（先读这个）
design/tokens.json        ← 单一真源：颜色 / 字号 / 栅格
design/check_tokens.py    ← 对比度校验器（改色后必跑）
assets/                   ← 校徽 / 南瑞 logo / 示例插图 / 公式 PNG（含 manifest）
build/                    ← 构建与验证脚本
out/                      ← 产物（模板示例 deck）
错例/                     ← 反例对照图 + README（给格式审查用）
```

## 最小用法

```bash
cd build
python make_formulas.py     # 渲染公式/算法 PNG（需 TeX Live）
python build_pptx.py        # 生成 out/seu_nari_template.pptx
python _audit.py            # 机器审查；退出码非 0 就不要交付
```

`assets/formulas/` 与 `assets/figures/` 里的 PNG **已提交进仓库**，所以没有 TeX Live
也能直接跑 `build_pptx.py`。内容全在 `build/slides.py`（页）与 `build/build_pptx.py`
的 `DECK`（文案）里。

## 特性

- **双品牌**：东南大学校徽（内容页右上，2.91 cm）+ 国网电科院（南瑞）logo；
  配色统一为「东大绿为主 + 南瑞墨绿为辅 + 金色细线」
- **两级进度条**：章节级绿段画在**版式**上（章内插页不会错乱）；章内页级金线画在**页**上
- **公式与算法走 LaTeX 渲染**：公式用 Latin Modern，算法是论文三线式（`Algorithm N: 名称` + 行号）
- **可编辑**：版式上只放图形与占位符，**所有文字都在页上**，普通视图里点得中、改得了
- **交付前机器审查**：`_audit.py` 查公式倍率漂移 / 出血 / 字号过小 / 空标题
